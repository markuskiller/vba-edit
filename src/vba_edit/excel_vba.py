import argparse
import logging
import os
import subprocess
import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.cli_common import (
    add_common_arguments,
    add_encoding_arguments,
    add_excel_specific_arguments,
    add_export_arguments,
    add_folder_organization_arguments,
    add_header_arguments,
    add_metadata_arguments,
    handle_export_with_warnings,
    process_config_file,
    validate_header_options,
)
from vba_edit.exceptions import (
    ApplicationError,
    DocumentClosedError,
    DocumentNotFoundError,
    PathError,
    RPCError,
    VBAAccessError,
    VBAError,
)
from vba_edit.office_vba import ExcelVBAHandler
from vba_edit.path_utils import get_document_paths
from vba_edit.utils import get_active_office_document, get_windows_ansi_codepage, setup_logging

# Configure module logger
logger = logging.getLogger(__name__)


def create_cli_parser() -> argparse.ArgumentParser:
    """Create the command-line interface parser."""
    entry_point_name = "excel-vba"
    package_name_formatted = package_name.replace("_", "-")

    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    parser = argparse.ArgumentParser(
        prog=entry_point_name,
        description=f"""
{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool suite for managing VBA content in MS Office documents.

EXCEL-VBA allows you to edit, import, and export VBA content from Excel workbooks.
If no file is specified, the tool will attempt to use the currently active Excel workbook.

Examples:
    excel-vba edit   <--- uses active Excel workbook and current directory for exported 
                         VBA files (*.bas/*.cls/*.frm) & syncs changes back to the 
                         active Excel workbook on save

    excel-vba import -f "C:/path/to/workbook.xlsm" --vba-directory "path/to/vba/files"
    excel-vba export --file "C:/path/to/workbook.xlsm" --encoding cp850 --save-metadata
    excel-vba export --conf "path/to/conf/file" --in-file-headers --force-overwrite
    excel-vba edit --vba-directory "path/to/vba/files" --logfile "path/to/logfile" --verbose
    excel-vba edit --save-headers --rubberduck-folders --open-folder


IMPORTANT: 
           [!] It's early days. Use with care and backup your important macro-enabled
               MS Office documents before using them with this tool!

           [!] This tool requires "Trust access to the VBA project object model" 
               enabled in Excel.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # Add --version argument to the main parser
    parser.add_argument(
        "--version", "-V", action="version", version=f"{package_name_formatted} v{package_version} ({entry_point_name})"
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Edit command
    edit_parser = subparsers.add_parser("edit", help="Edit VBA content in Excel workbook")
    add_common_arguments(edit_parser)
    add_excel_specific_arguments(edit_parser)
    add_encoding_arguments(edit_parser, default_encoding)
    add_header_arguments(edit_parser)
    add_folder_organization_arguments(edit_parser)  # Issue #21 fix

    # Import command
    import_parser = subparsers.add_parser("import", help="Import VBA content into Excel workbook")
    add_common_arguments(import_parser)
    add_excel_specific_arguments(import_parser)
    add_encoding_arguments(import_parser, default_encoding)
    add_header_arguments(import_parser)
    add_folder_organization_arguments(import_parser)  # Issue #21 fix

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA content from Excel workbook")
    add_common_arguments(export_parser)
    add_excel_specific_arguments(export_parser)
    add_encoding_arguments(export_parser, default_encoding)
    add_header_arguments(export_parser)
    add_metadata_arguments(export_parser)
    add_export_arguments(export_parser)
    add_folder_organization_arguments(export_parser)  # Issue #21 fix

    # Check command - only needs basic args
    check_parser = subparsers.add_parser(
        "check",
        help="Check if 'Trust Access to the MS Excel VBA project object model' is enabled",
    )
    check_subparser = check_parser.add_subparsers(dest="subcommand", required=False)
    check_subparser.add_parser(
        "all", help="Check Trust Access to VBA project model of all suported Office applications"
    )

    return parser


def handle_excel_vba_command(args: argparse.Namespace) -> None:
    """Handle the excel-vba command execution."""
    try:
        # Initialize logging
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
        logger.debug(f"Starting excel-vba command: {args.command}")
        logger.debug(f"Command arguments: {vars(args)}")

        # Ensure paths exist early (creates vba_directory if provided)
        validate_paths(args)

        # Handle xlwings option if present
        if getattr(args, "xlwings", False):
            try:
                import xlwings

                logger.info(f"Using xlwings {xlwings.__version__}")
                handle_xlwings_command(args)
                return
            except ImportError:
                sys.exit("xlwings is not installed. Please install it with: pip install xlwings")

        # Get document path and active document path
        active_doc = None
        if not args.file:
            try:
                active_doc = get_active_office_document("excel")
            except ApplicationError:
                pass

        try:
            doc_path, vba_dir = get_document_paths(args.file, active_doc, args.vba_directory)
            logger.info(f"Using workbook: {doc_path}")
            logger.debug(f"Using VBA directory: {vba_dir}")
        except (DocumentNotFoundError, PathError) as e:
            logger.error(f"Failed to resolve paths: {str(e)}")
            sys.exit(1)

        # Determine encoding
        encoding = None if getattr(args, "detect_encoding", False) else args.encoding
        logger.debug(f"Using encoding: {encoding or 'auto-detect'}")

        # Validate header options
        validate_header_options(args)

        # Create handler instance
        try:
            handler = ExcelVBAHandler(
                doc_path=str(doc_path),
                vba_dir=str(vba_dir),
                encoding=encoding,
                verbose=getattr(args, "verbose", False),
                save_headers=getattr(args, "save_headers", False),
                use_rubberduck_folders=getattr(args, "rubberduck_folders", False),
                open_folder=getattr(args, "open_folder", False),
                in_file_headers=getattr(args, "in_file_headers", True),
            )
        except VBAError as e:
            logger.error(f"Failed to initialize Excel VBA handler: {str(e)}")
            sys.exit(1)

        # Execute requested command
        logger.info(f"Executing command: {args.command}")
        try:
            if args.command == "edit":
                print("NOTE: Deleting a VBA module file will also delete it in the VBA editor!")
                handle_export_with_warnings(handler, overwrite=False, interactive=True)
                try:
                    handler.watch_changes()
                except (DocumentClosedError, RPCError) as e:
                    logger.error(str(e))
                    logger.info("Edit session terminated. Please restart Excel and the tool to continue editing.")
                    sys.exit(1)
            elif args.command == "import":
                handler.import_vba()
            elif args.command == "export":
                handle_export_with_warnings(
                    handler,
                    save_metadata=getattr(args, "save_metadata", False),
                    overwrite=True,
                    interactive=True,
                    force_overwrite=getattr(args, "force_overwrite", False),
                )
        except (DocumentClosedError, RPCError) as e:
            logger.error(str(e))
            sys.exit(1)
        except VBAAccessError as e:
            logger.error(str(e))
            logger.error("Please check Excel Trust Center Settings and try again.")
            sys.exit(1)
        except VBAError as e:
            logger.error(f"VBA operation failed: {str(e)}")
            sys.exit(1)
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}")
            if getattr(args, "verbose", False):
                logger.exception("Detailed error information:")
            sys.exit(1)

    except KeyboardInterrupt:
        logger.info("\nOperation interrupted by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Critical error: {str(e)}")
        if getattr(args, "verbose", False):
            logger.exception("Detailed error information:")
        sys.exit(1)
    finally:
        logger.debug("Command execution completed")


def handle_xlwings_command(args):
    """Handle command execution using xlwings wrapper."""

    # Convert our args to xlwings command format
    cmd = ["xlwings", "vba", args.command]

    if args.file:
        cmd.extend(["-f", args.file])
    if args.verbose:
        cmd.extend(["-v"])

    # Store original directory
    original_dir = None
    try:
        # Change to target directory if specified
        if args.vba_directory:
            original_dir = os.getcwd()
            target_dir = Path(args.vba_directory)
            # Create directory if it doesn't exist
            target_dir.mkdir(parents=True, exist_ok=True)
            os.chdir(str(target_dir))

        # Execute xlwings command
        result = subprocess.run(cmd, check=True)
        sys.exit(result.returncode)

    except subprocess.CalledProcessError as e:
        sys.exit(e.returncode)
    finally:
        # Restore original directory if we changed it
        if original_dir:
            os.chdir(original_dir)


def validate_paths(args: argparse.Namespace) -> None:
    """Validate file and directory paths from command line arguments.
    
    Only validates paths for commands that actually use them (edit, import, export).
    The 'check' command doesn't use file/vba_directory arguments.
    """
    # Skip validation for commands that don't use file paths
    if args.command == "check":
        return
    
    if hasattr(args, 'file') and args.file and not Path(args.file).exists():
        raise FileNotFoundError(f"File not found: {args.file}")

    if hasattr(args, 'vba_directory') and args.vba_directory:
        vba_dir = Path(args.vba_directory)
        if not vba_dir.exists():
            logger.info(f"Creating VBA directory: {vba_dir}")
            vba_dir.mkdir(parents=True, exist_ok=True)


def main() -> None:
    """Main entry point for the excel-vba CLI."""
    try:
        parser = create_cli_parser()
        args = parser.parse_args()

        # Process configuration file BEFORE setting up logging
        args = process_config_file(args)

        # Set up logging first
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))

        # Create target directories and validate inputs early
        validate_paths(args)

        # Run 'check' command (Check if VBA project model is accessible )
        if args.command == "check":
            from vba_edit.utils import check_vba_trust_access

            try:
                if args.subcommand == "all":
                    check_vba_trust_access()  # Check all supported Office applications
                else:
                    check_vba_trust_access("excel")  # Check MS Excel only
            except Exception as e:
                logger.error(f"Failed to check Trust Access to VBA project object model: {str(e)}")
            sys.exit(0)

        # If xlwings option is used, check dependency before proceeding
        elif getattr(args, "xlwings", False):
            try:
                import xlwings

                logger.info(f"Using xlwings {xlwings.__version__}")
                handle_xlwings_command(args)
            except ImportError:
                sys.exit("xlwings is not installed. Please install it with: pip install xlwings")
        else:
            handle_excel_vba_command(args)

    except Exception as e:
        print(f"Critical error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

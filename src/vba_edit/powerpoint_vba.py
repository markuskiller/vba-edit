import argparse
import logging
import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.exceptions import (
    ApplicationError,
    DocumentNotFoundError,
    DocumentClosedError,
    PathError,
    RPCError,
    VBAAccessError,
    VBAError,
)
from vba_edit.office_vba import PowerPointVBAHandler
from vba_edit.path_utils import get_document_paths
from vba_edit.utils import get_active_office_document, get_windows_ansi_codepage, setup_logging
from vba_edit.cli_common import (
    add_common_arguments,
    process_config_file,
    add_encoding_arguments,
    add_header_arguments,
    add_metadata_arguments,
    validate_header_options,
)


# Configure module logger
logger = logging.getLogger(__name__)


def create_cli_parser() -> argparse.ArgumentParser:
    """Create the command-line interface parser."""
    entry_point_name = "powerpoint-vba"
    package_name_formatted = package_name.replace("_", "-")

    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    parser = argparse.ArgumentParser(
        prog=entry_point_name,
        description=f"""
{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool suite for managing VBA content in MS Office documents.

POWERPOINT-VBA allows you to edit, import, and export VBA content from PowerPoint documents.
If no file is specified, the tool will attempt to use the currently active PowerPoint document.

Commands:
    edit    Edit VBA content in MS PowerPoint document
    import  Import VBA content into MS PowerPoint document
    export  Export VBA content from MS PowerPoint document
    check   Check if 'Trust Access to the MS PowerPoint VBA project object model' is enabled

Examples:
    powerpoint-vba edit    <--- uses active PowerPoint document and current directory for exported 
                           VBA files (*.bas/*.cls/*.frm) & syncs changes back to the 
                           active PowerPoint document on save

    powerpoint-vba import -f "C:/path/to/document.docx" --vba-directory "path/to/vba/files"
    powerpoint-vba export --file "C:/path/to/document.docx" --encoding cp850 --save-metadata
    powerpoint-vba edit --vba-directory "path/to/vba/files" --logfile "path/to/logfile" --verbose
    powerpoint-vba edit --save-headers

IMPORTANT: 
           [!] It's early days. Use with care and backup your imortant macro-enabled
               MS Office documents before using them with this tool!

               First tests have been very promissing. Feedback appreciated via
               github issues. 

           [!] This tool requires "Trust access to the VBA project object model" 
               enabled in PowerPoint.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # Add --version argument to the main parser
    parser.add_argument(
        "--version", action="version", version=f"{package_name_formatted} v{package_version} ({entry_point_name})"
    )
    add_common_arguments(parser)

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Edit command
    edit_parser = subparsers.add_parser("edit", help="Edit VBA content in PowerPoint document")
    add_common_arguments(edit_parser)
    add_encoding_arguments(edit_parser, default_encoding)
    add_header_arguments(edit_parser)

    # Import command
    import_parser = subparsers.add_parser("import", help="Import VBA content into PowerPoint document")
    add_common_arguments(import_parser)
    add_encoding_arguments(import_parser, default_encoding)
    add_header_arguments(import_parser)

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA content from PowerPoint document")
    add_common_arguments(export_parser)
    add_encoding_arguments(export_parser, default_encoding)
    add_header_arguments(export_parser)
    add_metadata_arguments(export_parser)

    # Check command
    check_parser = subparsers.add_parser(
        "check",
        help="Check if 'Trust Access to the MS PowerPoint VBA project object model' is enabled",
    )
    add_common_arguments(check_parser)

    return parser


def validate_paths(args: argparse.Namespace) -> None:
    """Validate file and directory paths from command line arguments."""
    if args.file and not Path(args.file).exists():
        raise FileNotFoundError(f"Document not found: {args.file}")

    if args.vba_directory:
        vba_dir = Path(args.vba_directory)
        if not vba_dir.exists():
            logger.info(f"Creating VBA directory: {vba_dir}")
            vba_dir.mkdir(parents=True, exist_ok=True)


def handle_powerpoint_vba_command(args: argparse.Namespace) -> None:
    """Handle the powerpoint-vba command execution."""
    try:
        # Initialize logging
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
        logger.debug(f"Starting powerpoint-vba command: {args.command}")
        logger.debug(f"Command arguments: {vars(args)}")

        # Get document path and active document path
        active_doc = None
        if not args.file:
            try:
                active_doc = get_active_office_document("powerpoint")
            except ApplicationError:
                pass

        try:
            doc_path, vba_dir = get_document_paths(args.file, active_doc, args.vba_directory)
            logger.info(f"Using document: {doc_path}")
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
            handler = PowerPointVBAHandler(
                doc_path=str(doc_path),
                vba_dir=str(vba_dir),
                encoding=encoding,
                verbose=getattr(args, "verbose", False),
                save_headers=getattr(args, "save_headers", False),
                use_rubberduck_folders=args.rubberduck_folders,
                open_folder=args.open_folder,
                in_file_headers=getattr(args, "in_file_headers", True),
            )
        except VBAError as e:
            logger.error(f"Failed to initialize PowerPoint VBA handler: {str(e)}")
            sys.exit(1)

        # Execute requested command
        logger.info(f"Executing command: {args.command}")
        try:
            if args.command == "edit":
                # Display warning about module deletion
                print("NOTE: Deleting a VBA module file will also delete it in the VBA editor!")
                # For edit command, first export without overwriting
                handler.export_vba(overwrite=False)
                try:
                    handler.watch_changes()
                except (DocumentClosedError, RPCError) as e:
                    logger.error(str(e))
                    logger.info("Edit session terminated. Please restart PowerPoint and the tool to continue editing.")
                    sys.exit(1)
            elif args.command == "import":
                handler.import_vba()
            elif args.command == "export":
                # For export command, always overwrite
                handler.export_vba(save_metadata=getattr(args, "save_metadata", False), overwrite=True)
        except (DocumentClosedError, RPCError) as e:
            logger.error(str(e))
            sys.exit(1)
        except VBAAccessError as e:
            logger.error(str(e))
            logger.error("Please check PowerPoint Trust Center Settings and try again.")
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


def main() -> None:
    """Main entry point for the powerpoint-vba CLI."""
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
                    check_vba_trust_access("powerpoint")  # Check MS PowerPoint only
            except Exception as e:
                logger.error(f"Failed to check Trust Access to VBA project object model: {str(e)}")
            sys.exit(0)
        else:
            handle_powerpoint_vba_command(args)

    except Exception as e:
        print(f"Critical error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

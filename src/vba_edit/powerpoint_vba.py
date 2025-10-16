import argparse
import logging
import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.cli_common import (
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
from vba_edit.help_formatter import ColorizedArgumentParser, EnhancedHelpFormatter
from vba_edit.office_vba import PowerPointVBAHandler
from vba_edit.path_utils import get_document_paths
from vba_edit.utils import get_active_office_document, get_windows_ansi_codepage, setup_logging

# Configure module logger
logger = logging.getLogger(__name__)


def create_cli_parser() -> argparse.ArgumentParser:
    """Create the command-line interface parser."""
    entry_point_name = "powerpoint-vba"
    package_name_formatted = package_name.replace("_", "-")

    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    # Create streamlined main help description
    main_description = f"""{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool for managing VBA content in PowerPoint presentations.
Export, import, and edit VBA code in external editor with live sync back to presentation."""

    # Create streamlined examples for main help
    main_examples = f"""
Examples:
  {entry_point_name} edit                           # Edit in external editor, sync changes to presentation
  {entry_point_name} export -f file.pptm            # Export VBA to directory
  {entry_point_name} import --vba-directory src     # Import VBA from directory
  {entry_point_name} check                          # Verify VBA access is enabled

Use '{entry_point_name} <command> --help' for more information on a specific command.

IMPORTANT: Requires "Trust access to VBA project object model" enabled in PowerPoint.
           Early release - backup important files before use!"""

    parser = ColorizedArgumentParser(
        prog=entry_point_name,
        usage=f"{entry_point_name} [--help] [--version] <command> [<args>]",
        description=main_description,
        epilog=main_examples,
        formatter_class=EnhancedHelpFormatter,
    )

    # Add --version argument to the main parser
    parser.add_argument(
        "--version", "-V", action="version", version=f"{package_name_formatted} v{package_version} ({entry_point_name})"
    )

    # Add hidden easter egg flags (not shown in help)
    parser.add_argument("--diagram", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--how-it-works", action="store_true", help=argparse.SUPPRESS)

    subparsers = parser.add_subparsers(
        dest="command",
        required=True,
        title="Commands",
        metavar="<command>",
    )

    # Edit command
    edit_description = """Export VBA code from Office document to filesystem, edit VBA code files in external editor and sync changes back into Office document on save [CTRL+S]

Simple usage:
  powerpoint-vba edit       # Uses active presentation and VBA code files are 
                            # saved in the same location as the presentation
  
Full control usage:
  powerpoint-vba edit -f file.pptm --vba-directory src"""

    edit_usage = """powerpoint-vba edit
    [--file FILE | -f FILE]
    [--vba-directory DIR]
    [--open-folder]
    [--conf FILE | --config FILE]
    [--encoding ENCODING | -e ENCODING | --detect-encoding | -d]
    [--save-headers | --in-file-headers]
    [--save-metadata | -m]
    [--rubberduck-folders]
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    edit_parser = subparsers.add_parser(
        "edit",
        usage=edit_usage,
        help="Edit in external editor, sync changes back to presentation",
        description=edit_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,  # Suppress default help to add it manually in Common Options
    )

    # File Options group
    file_group = edit_parser.add_argument_group("File Options")
    file_group.add_argument(
        "--file",
        "-f",
        dest="file",
        help="Path to Office document (default: active document)",
    )
    file_group.add_argument(
        "--vba-directory",
        dest="vba_directory",
        metavar="DIR",
        help="Directory to export VBA files (default: same directory as document)",
    )
    file_group.add_argument(
        "--open-folder",
        dest="open_folder",
        action="store_true",
        help="Open export directory in file explorer after export",
    )

    # Configuration group
    config_group = edit_parser.add_argument_group("Configuration")
    config_group.add_argument(
        "--conf",
        "--config",
        metavar="FILE",
        help="Path to configuration file (TOML format)",
    )

    # Encoding Options group (mutually exclusive)
    encoding_group = edit_parser.add_argument_group("Encoding Options (mutually exclusive)")
    encoding_mutex = encoding_group.add_mutually_exclusive_group()
    encoding_mutex.add_argument(
        "--encoding",
        "-e",
        dest="encoding",
        metavar="ENCODING",
        default=default_encoding,
        help=f"Encoding for writing VBA files (default: {default_encoding})",
    )
    encoding_mutex.add_argument(
        "--detect-encoding",
        "-d",
        dest="detect_encoding",
        action="store_true",
        help="Auto-detect file encoding for VBA files",
    )

    # Header Options group (mutually exclusive)
    header_group = edit_parser.add_argument_group("Header Options (mutually exclusive)")
    header_mutex = header_group.add_mutually_exclusive_group()
    header_mutex.add_argument(
        "--save-headers",
        dest="save_headers",
        action="store_true",
        help="Save VBA component headers to separate .header files",
    )
    header_mutex.add_argument(
        "--in-file-headers",
        dest="in_file_headers",
        action="store_true",
        help="Include VBA headers directly in code files",
    )

    # Edit Options group
    edit_options_group = edit_parser.add_argument_group("Edit Options")
    edit_options_group.add_argument(
        "--save-metadata",
        "-m",
        dest="save_metadata",
        action="store_true",
        help="Save metadata to file",
    )
    edit_options_group.add_argument(
        "--rubberduck-folders",
        dest="rubberduck_folders",
        action="store_true",
        help="Organize folders per RubberduckVBA @Folder annotations",
    )

    # Common Options group
    common_group = edit_parser.add_argument_group("Common Options")
    common_group.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    common_group.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    common_group.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    common_group.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # Import command
    import_description = """Import VBA code from filesystem into Office document

Header handling is automatic - no flags needed:
  • Detects inline headers (VERSION/BEGIN/Attribute at file start)
  • Falls back to separate .header files if present
  • Creates minimal headers if neither exists

Simple usage:
  powerpoint-vba import     # Uses active document and imports from same directory

Full control usage:
  powerpoint-vba import -f presentation.pptm --vba-directory src"""

    import_usage = """powerpoint-vba import
    [--file FILE | -f FILE]
    [--vba-directory DIR]
    [--conf FILE | --config FILE]
    [--encoding ENCODING | -e ENCODING | --detect-encoding | -d]
    [--rubberduck-folders]
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    import_parser = subparsers.add_parser(
        "import",
        usage=import_usage,
        help="Import VBA from filesystem into presentation",
        description=import_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # File Options group
    file_group = import_parser.add_argument_group("File Options")
    file_group.add_argument(
        "--file",
        "-f",
        dest="file",
        help="Path to Office document (default: active document)",
    )
    file_group.add_argument(
        "--vba-directory",
        dest="vba_directory",
        metavar="DIR",
        help="Directory to import VBA files from (default: same directory as document)",
    )

    # Configuration group
    config_group = import_parser.add_argument_group("Configuration")
    config_group.add_argument(
        "--conf",
        "--config",
        metavar="FILE",
        help="Path to configuration file (TOML format)",
    )

    # Encoding Options group (mutually exclusive)
    encoding_group = import_parser.add_argument_group("Encoding Options (mutually exclusive)")
    encoding_mutex = encoding_group.add_mutually_exclusive_group()
    encoding_mutex.add_argument(
        "--encoding",
        "-e",
        dest="encoding",
        metavar="ENCODING",
        default=default_encoding,
        help=f"Encoding for reading VBA files (default: {default_encoding})",
    )
    encoding_mutex.add_argument(
        "--detect-encoding",
        "-d",
        dest="detect_encoding",
        action="store_true",
        help="Auto-detect file encoding for VBA files",
    )

    # Import Options group
    import_options_group = import_parser.add_argument_group("Import Options")
    import_options_group.add_argument(
        "--rubberduck-folders",
        dest="rubberduck_folders",
        action="store_true",
        help="Organize folders per RubberduckVBA @Folder annotations",
    )

    # Common Options group
    common_group = import_parser.add_argument_group("Common Options")
    common_group.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    common_group.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    common_group.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    common_group.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # Export command
    export_description = """Export VBA code from Office document to filesystem

Simple usage:
  powerpoint-vba export     # Uses active presentation and exports to same directory
  
Full control usage:
  powerpoint-vba export -f file.pptm --vba-directory src"""

    export_usage = """powerpoint-vba export
    [--file FILE | -f FILE]
    [--keep-open]
    [--vba-directory DIR]
    [--open-folder]
    [--conf FILE | --config FILE]
    [--encoding ENCODING | -e ENCODING | --detect-encoding | -d]
    [--save-headers | --in-file-headers]
    [--save-metadata | -m]
    [--force-overwrite]
    [--rubberduck-folders]
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    export_parser = subparsers.add_parser(
        "export",
        usage=export_usage,
        help="Export VBA from presentation to filesystem",
        description=export_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # File Options group
    file_group = export_parser.add_argument_group("File Options")
    file_group.add_argument(
        "--file",
        "-f",
        dest="file",
        help="Path to Office document (default: active document)",
    )
    file_group.add_argument(
        "--vba-directory",
        dest="vba_directory",
        metavar="DIR",
        help="Directory to export VBA files (default: same directory as document)",
    )
    file_group.add_argument(
        "--open-folder",
        dest="open_folder",
        action="store_true",
        help="Open export directory in file explorer after export",
    )

    # Configuration group
    config_group = export_parser.add_argument_group("Configuration")
    config_group.add_argument(
        "--conf",
        "--config",
        metavar="FILE",
        help="Path to configuration file (TOML format)",
    )

    # Encoding Options group (mutually exclusive)
    encoding_group = export_parser.add_argument_group("Encoding Options (mutually exclusive)")
    encoding_mutex = encoding_group.add_mutually_exclusive_group()
    encoding_mutex.add_argument(
        "--encoding",
        "-e",
        dest="encoding",
        metavar="ENCODING",
        default=default_encoding,
        help=f"Encoding for writing VBA files (default: {default_encoding})",
    )
    encoding_mutex.add_argument(
        "--detect-encoding",
        "-d",
        dest="detect_encoding",
        action="store_true",
        help="Auto-detect file encoding for VBA files",
    )

    # Header Options group (mutually exclusive)
    header_group = export_parser.add_argument_group("Header Options (mutually exclusive)")
    header_mutex = header_group.add_mutually_exclusive_group()
    header_mutex.add_argument(
        "--save-headers",
        dest="save_headers",
        action="store_true",
        help="Save VBA component headers to separate .header files",
    )
    header_mutex.add_argument(
        "--in-file-headers",
        dest="in_file_headers",
        action="store_true",
        help="Include VBA headers directly in code files",
    )

    # Export Options group
    export_options_group = export_parser.add_argument_group("Export Options")
    export_options_group.add_argument(
        "--save-metadata",
        "-m",
        dest="save_metadata",
        action="store_true",
        help="Save metadata to file",
    )
    export_options_group.add_argument(
        "--force-overwrite",
        dest="force_overwrite",
        action="store_true",
        help="Force overwrite of existing files without prompting",
    )
    export_options_group.add_argument(
        "--keep-open",
        dest="keep_open",
        action="store_true",
        help="Keep document open after export (default: close after export)",
    )
    export_options_group.add_argument(
        "--rubberduck-folders",
        dest="rubberduck_folders",
        action="store_true",
        help="Organize folders per RubberduckVBA @Folder annotations",
    )

    # Common Options group
    common_group = export_parser.add_argument_group("Common Options")
    common_group.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    common_group.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    common_group.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    common_group.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # Check command
    check_description = """Check if 'Trust access to the VBA project object model' is enabled

Simple usage:
  powerpoint-vba check      # Check PowerPoint VBA access
  powerpoint-vba check all  # Check all Office applications"""

    check_usage = """powerpoint-vba check
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    check_parser = subparsers.add_parser(
        "check",
        usage=check_usage,
        help="Check VBA project access settings",
        description=check_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # Common Options group
    common_group = check_parser.add_argument_group("Common Options")
    common_group.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    common_group.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    common_group.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    common_group.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # Subcommand for checking all applications
    # Note: Using custom usage above, so subparser metavar doesn't affect main help
    check_subparser = check_parser.add_subparsers(
        dest="subcommand",
        required=False,
        title="Subcommands",
        metavar="",  # Empty metavar to avoid space in usage line
    )
    check_all_parser = check_subparser.add_parser(
        "all",
        help="Check Trust Access to VBA project model of all supported Office applications",
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # Common Options group for 'check all' subcommand
    check_all_common = check_all_parser.add_argument_group("Common Options")
    check_all_common.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    check_all_common.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    check_all_common.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    return parser


def validate_paths(args: argparse.Namespace) -> None:
    """Validate file and directory paths from command line arguments.

    Only validates paths for commands that actually use them (edit, import, export).
    The 'check' command doesn't use file/vba_directory arguments.
    """
    # Skip validation for commands that don't use file paths
    if args.command == "check":
        return

    if hasattr(args, "file") and args.file and not Path(args.file).exists():
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
                use_rubberduck_folders=getattr(args, "rubberduck_folders", False),
                open_folder=getattr(args, "open_folder", False),
                in_file_headers=getattr(args, "in_file_headers", True),
            )
        except VBAError as e:
            logger.error(f"Failed to initialize PowerPoint VBA handler: {str(e)}")
            sys.exit(1)

        # Execute requested command
        logger.info(f"Executing command: {args.command}")
        try:
            if args.command == "edit":
                logger.info("NOTE: Deleting a VBA module file will also delete it in the VBA editor!")
                handle_export_with_warnings(
                    handler,
                    save_metadata=getattr(args, "save_metadata", False),
                    overwrite=False,
                    interactive=True,
                    keep_open=True,  # CRITICAL: Must keep document open for edit mode
                )
                try:
                    handler.watch_changes()
                except (DocumentClosedError, RPCError) as e:
                    logger.error(str(e))
                    logger.info("Edit session terminated. Please restart PowerPoint and the tool to continue editing.")
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
                    keep_open=getattr(args, "keep_open", False),
                )
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
        # Check for --no-color flag BEFORE creating parser
        # This ensures help messages honor the flag
        if "--no-color" in sys.argv or "--no-colour" in sys.argv:
            from vba_edit.console import disable_colors

            disable_colors()

        # Handle easter egg flags first (before argparse validation)
        # This allows them to work without requiring a command
        if "--diagram" in sys.argv or "--how-it-works" in sys.argv:
            from vba_edit.utils import show_workflow_diagram

            show_workflow_diagram()

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
        from vba_edit.console import error

        error(f"Critical error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

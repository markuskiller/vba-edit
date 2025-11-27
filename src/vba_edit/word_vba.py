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
from vba_edit.office_vba import WordVBAHandler
from vba_edit.path_utils import get_document_paths
from vba_edit.utils import get_active_office_document, get_windows_ansi_codepage, setup_logging

# Configure module logger
logger = logging.getLogger(__name__)


def create_cli_parser() -> argparse.ArgumentParser:
    """Create the command-line interface parser."""
    entry_point_name = "word-vba"
    package_name_formatted = package_name.replace("_", "-")

    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    # Create streamlined main help description
    main_description = f"""{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool for managing VBA content in Word documents.
Export, import, and edit VBA code in external editor with live sync back to document."""

    # Create streamlined examples for main help
    main_examples = f"""
Examples:
  {entry_point_name} edit                           # Edit in external editor, sync changes to document
  {entry_point_name} export -f file.docm            # Export VBA to directory
  {entry_point_name} import --vba-directory src     # Import VBA from directory
  {entry_point_name} check                          # Verify VBA access is enabled

Use '{entry_point_name} <command> --help' for more information on a specific command.

IMPORTANT: Requires "Trust access to VBA project object model" enabled in Word.
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
  word-vba edit             # Uses active document and VBA code files are 
                            # saved in the same location as the document
  
Full control usage:
  word-vba edit -f file.docm --vba-directory src"""

    edit_usage = """word-vba edit
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
        help="Edit in external editor, sync changes back to document",
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
  word-vba import           # Uses active document and imports from same directory

Full control usage:
  word-vba import -f document.docm --vba-directory src"""

    import_usage = """word-vba import
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
        help="Import VBA from filesystem into document",
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
  word-vba export           # Uses active document and exports to same directory
  
Full control usage:
  word-vba export -f file.docm --vba-directory src"""

    export_usage = """word-vba export
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
        help="Export VBA from document to filesystem",
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
  word-vba check            # Check Word VBA access
  word-vba check all        # Check all Office applications"""

    check_usage = """word-vba check
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
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    check_all_common.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # References command
    references_description = """Manage VBA references in Word documents

VBA references are links to external type libraries (COM objects, DLLs, other Office documents)
that provide additional functionality to your VBA code.

Simple usage:
  word-vba references list                    # List refs in active document
  word-vba references export                  # Export to {document}_refs.toml
  word-vba references import -r refs.toml     # Import from TOML file"""

    references_usage = """word-vba references <command> [options]"""

    references_parser = subparsers.add_parser(
        "references",
        usage=references_usage,
        help="Manage VBA references (list/export/import)",
        description=references_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # Subparsers for references subcommands (must come before common options for proper display order)
    references_subparsers = references_parser.add_subparsers(
        dest="refs_subcommand",
        required=True,
        title="Commands",
        metavar="<command>",
    )

    # References list subcommand
    list_refs_description = """List all VBA references in a Word document

Shows reference details including:
  • Name and description
  • GUID and version
  • File path (for file-based references)
  • Status (built-in, broken, etc.)

Examples:
  word-vba references list              # List refs in active document
  word-vba references list -f file.docm # List refs in specific file"""

    list_refs_usage = """word-vba references list
    [--file FILE | -f FILE]
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    list_refs_parser = references_subparsers.add_parser(
        "list",
        usage=list_refs_usage,
        help="List all references in document",
        description=list_refs_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # File Options group for list
    list_file_group = list_refs_parser.add_argument_group("File Options")
    list_file_group.add_argument(
        "--file",
        "-f",
        dest="file",
        help="Path to Word document (default: active document)",
    )

    # Common Options group for list
    list_common = list_refs_parser.add_argument_group("Common Options")
    list_common.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    list_common.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    list_common.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    list_common.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # References export subcommand
    export_refs_description = """Export VBA references to TOML configuration file

Saves all references (except built-in ones) to a TOML file that can be:
  • Shared with team members
  • Version controlled
  • Used to replicate reference setup in other documents

The TOML file contains reference details (name, GUID, version, path) and can
be edited manually if needed.

Examples:
  word-vba references export                        # Export to {document}_refs.toml
  word-vba references export -r custom_refs.toml    # Export to custom file
  word-vba references export -f template.dotm -r refs.toml"""

    export_refs_usage = """word-vba references export
    [--file FILE | -f FILE]
    [--refs-file FILE | -r FILE]
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    export_refs_parser = references_subparsers.add_parser(
        "export",
        usage=export_refs_usage,
        help="Export references to TOML file",
        description=export_refs_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # File Options group for export
    export_refs_file_group = export_refs_parser.add_argument_group("File Options")
    export_refs_file_group.add_argument(
        "--file",
        "-f",
        dest="file",
        help="Path to Word document (default: active document)",
    )
    export_refs_file_group.add_argument(
        "--refs-file",
        "-r",
        dest="refs_file",
        metavar="FILE",
        help="Output TOML file path (default: {document}_refs.toml)",
    )

    # Common Options group for export
    export_refs_common = export_refs_parser.add_argument_group("Common Options")
    export_refs_common.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    export_refs_common.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    export_refs_common.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    export_refs_common.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # References import subcommand
    import_refs_description = """Import VBA references from TOML configuration file

Reads reference definitions from a TOML file and adds them to a Word document.
Handles:
  • Duplicate detection (skips already present references)
  • Path validation (for file-based references)
  • Error reporting (missing files, invalid GUIDs, etc.)

Use this to:
  • Set up references in new documents
  • Replicate reference configuration across team
  • Restore references after document recreation

Examples:
  word-vba references import -r refs.toml           # Import to active document
  word-vba references import -r refs.toml -f new.docm"""

    import_refs_usage = """word-vba references import
    --refs-file FILE | -r FILE
    [--file FILE | -f FILE]
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    import_refs_parser = references_subparsers.add_parser(
        "import",
        usage=import_refs_usage,
        help="Import references from TOML file",
        description=import_refs_description,
        formatter_class=EnhancedHelpFormatter,
        add_help=False,
    )

    # File Options group for import
    import_refs_file_group = import_refs_parser.add_argument_group("File Options")
    import_refs_file_group.add_argument(
        "--file",
        "-f",
        dest="file",
        help="Path to Word document (default: active document)",
    )
    import_refs_file_group.add_argument(
        "--refs-file",
        "-r",
        dest="refs_file",
        metavar="FILE",
        required=True,
        help="Input TOML file with reference definitions",
    )

    # Common Options group for import
    import_refs_common = import_refs_parser.add_argument_group("Common Options")
    import_refs_common.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    import_refs_common.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    import_refs_common.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    import_refs_common.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )

    # Common Options group for references command (added after subcommands for proper display order)
    references_common = references_parser.add_argument_group("Common Options")
    references_common.add_argument(
        "--verbose",
        "-v",
        dest="verbose",
        action="store_true",
        help="Enable verbose logging output",
    )
    references_common.add_argument(
        "--logfile",
        "-l",
        dest="logfile",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)",
    )
    references_common.add_argument(
        "--no-color",
        "--no-colour",
        dest="no_color",
        action="store_true",
        help="Disable colored output",
    )
    references_common.add_argument(
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

    if hasattr(args, "vba_directory") and args.vba_directory:
        vba_dir = Path(args.vba_directory)
        if not vba_dir.exists():
            logger.info(f"Creating VBA directory: {vba_dir}")
            vba_dir.mkdir(parents=True, exist_ok=True)


def handle_references_command(args: argparse.Namespace, doc_path: Path) -> None:
    """Handle the references command and its subcommands."""
    import win32com.client
    from vba_edit.console import error, info, warning
    from vba_edit.reference_manager import ReferenceManager

    logger.debug(f"Handling references subcommand: {args.refs_subcommand}")

    word = None
    doc = None
    try:
        # Open Word and the document
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        logger.debug(f"Opening document: {doc_path}")
        doc = word.Documents.Open(str(doc_path))

        # Create ReferenceManager instance with the document object
        manager = ReferenceManager(doc)
        logger.debug(f"Created ReferenceManager for: {doc_path}")

        if args.refs_subcommand == "list":
            # List all references
            info(f"Listing references in: {doc_path.name}")
            refs = manager.list_references()

            if not refs:
                warning("No references found in document")
                return

            info(f"Found {len(refs)} reference(s):")

            # Display each reference with simple formatting (avoid Unicode issues)
            for i, ref in enumerate(refs, 1):
                print(f"\n{i}. {ref['name']}")
                if ref.get("description"):
                    print(f"   Description: {ref['description']}")
                if ref.get("guid"):
                    print(f"   GUID: {ref['guid']}")
                if ref.get("version"):
                    print(f"   Version: {ref['version']}")
                if ref.get("path"):
                    print(f"   Path: {ref['path']}")
                print(f"   Built-in: {ref.get('builtin', False)}")
                print(f"   Broken: {ref.get('isbroken', False)}")

        elif args.refs_subcommand == "export":
            # Export references to TOML
            refs_file = args.refs_file
            if not refs_file:
                # Default: {document_name}_refs.toml in same directory
                refs_file = doc_path.parent / f"{doc_path.stem}_refs.toml"
            else:
                refs_file = Path(refs_file)

            info(f"Exporting references from: {doc_path.name}")
            manager.export_to_toml(refs_file)
            info(f"Exported to: {refs_file}")

        elif args.refs_subcommand == "import":
            # Import references from TOML
            refs_file = Path(args.refs_file)

            if not refs_file.exists():
                error(f"Reference file not found: {refs_file}")
                sys.exit(1)

            info(f"Importing references to: {doc_path.name}")
            info(f"   From file: {refs_file}")

            result = manager.import_from_toml(refs_file)

            # Report results
            if result["added"] > 0:
                info(f"Added {result['added']} reference(s)")
            if result["skipped"] > 0:
                info(f"Skipped {result['skipped']} (already present)")
            if result["failed"] > 0:
                warning(f"Failed {result['failed']} reference(s)")

            if result["added"] == 0 and result["failed"] == 0:
                info("All references were already present")

    except Exception as e:
        error(f"References operation failed: {str(e)}")
        logger.exception("Reference operation error")
        sys.exit(1)
    finally:
        # Cleanup - close document and Word
        try:
            if doc:
                doc.Close(SaveChanges=False)
                logger.debug("Closed document")
        except Exception as e:
            logger.debug(f"Error closing document: {e}")

        try:
            if word:
                word.Quit()
                logger.debug("Closed Word application")
        except Exception as e:
            logger.debug(f"Error closing Word: {e}")


def handle_word_vba_command(args: argparse.Namespace) -> None:
    """Handle the word-vba command execution."""
    try:
        # Initialize logging
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
        logger.debug(f"Starting word-vba command: {args.command}")
        logger.debug(f"Command arguments: {vars(args)}")

        # Get document path and active document path
        active_doc = None
        if not args.file:
            try:
                active_doc = get_active_office_document("word")
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
            handler = WordVBAHandler(
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
            logger.error(f"Failed to initialize Word VBA handler: {str(e)}")
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
                    logger.info("Edit session terminated. Please restart Word and the tool to continue editing.")
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
            logger.error("Please check Word Trust Center Settings and try again.")
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
    """Main entry point for the word-vba CLI."""
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
                    check_vba_trust_access("word")  # Check MS Word only
            except Exception as e:
                logger.error(f"Failed to check Trust Access to VBA project object model: {str(e)}")
            sys.exit(0)
        elif args.command == "references":
            # Handle references command separately (doesn't need vba_directory)
            active_doc = None
            if not args.file:
                try:
                    from vba_edit.utils import get_active_office_document

                    active_doc = get_active_office_document("word")
                except ApplicationError:
                    pass

            # Get document path (references command only needs document, not vba_dir)
            if args.file:
                doc_path = Path(args.file)
                if not doc_path.exists():
                    from vba_edit.console import error

                    error(f"Document not found: {doc_path}")
                    sys.exit(1)
            elif active_doc:
                doc_path = Path(active_doc)
            else:
                from vba_edit.console import error

                error("No document specified and no active Word document found")
                error("Use --file/-f to specify a document or open a document in Word")
                sys.exit(1)

            handle_references_command(args, doc_path)
        else:
            handle_word_vba_command(args)

    except Exception as e:
        from vba_edit.console import print_exception

        # Use print_exception to properly render Rich markup in exception messages
        print_exception(e)
        sys.exit(1)


if __name__ == "__main__":
    main()

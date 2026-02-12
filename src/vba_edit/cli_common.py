"""Common CLI argument definitions for all Office VBA handlers."""

import argparse
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, Optional

from vba_edit.console import error, info, warning
from vba_edit.exceptions import VBAExportWarning
from vba_edit.utils import confirm_action, get_windows_ansi_codepage

# Prefer stdlib tomllib (Py 3.11+), fallback to tomli for older envs
try:
    import tomllib as toml_lib  # Python 3.11+ includes tomllib in stdlib
except ImportError:
    try:
        import tomli as toml_lib  # tomli is the recommended TOML parser for Python <3.11 # type: ignore[import-not-found]
    except ImportError:
        import toml as toml_lib  # Fall back to toml package if tomli isn't available # type: ignore[import-not-found]

logger = logging.getLogger(__name__)


# region PRcompat
# for PR integration only, until test modules are updated
def add_folder_organization_arguments(parser: argparse.ArgumentParser) -> None:
    add_vba_files_arguments(parser)
    add_exporting_arguments(parser)

def add_export_arguments(parser: argparse.ArgumentParser) -> None:
    add_exporting_arguments(parser)

def add_common_arguments(parser: argparse.ArgumentParser) -> None:
    add_config_arguments(parser)
    add_command_arguments(parser)
    parser.add_argument(
        "--verbose",
        "-v",
        dest=CONFIG_KEY_VERBOSE,
        action="store_true",
        help="Enable verbose logging output",
    )


PLACEHOLDER_FILE_NAME_LEGACY = "{general.file.name}"
PLACEHOLDER_FILE_FULLNAME_LEGACY = "{general.file.fullname}"
PLACEHOLDER_FILE_PATH_LEGACY = "{general.file.path}"
PLACEHOLDER_VBA_PROJECT_LEGACY = "{vbaproject}"
PLACEHOLDER_VBA_PROJECT = PLACEHOLDER_VBA_PROJECT_LEGACY

def _placeholder_compatibility_stub(placeholders: Dict[str, str]) -> Dict[str, str]:
    placeholders[PLACEHOLDER_FILE_NAME_LEGACY] = placeholders[PLACEHOLDER_FILE_NAME]
    placeholders[PLACEHOLDER_FILE_FULLNAME_LEGACY] = placeholders[PLACEHOLDER_FILE_FULLNAME]
    placeholders[PLACEHOLDER_FILE_PATH_LEGACY] = placeholders[PLACEHOLDER_FILE_PATH]
    placeholders[PLACEHOLDER_VBA_PROJECT] = placeholders[PLACEHOLDER_FILE_VBAPROJECT]
    placeholders[PLACEHOLDER_VBA_PROJECT_LEGACY] = placeholders[PLACEHOLDER_FILE_VBAPROJECT]
    return placeholders
# endregion


def handle_export_with_warnings(
    handler,
    save_metadata: bool = False,
    overwrite: bool = True,
    interactive: bool = True,
    force_overwrite: bool = False,
    keep_open: bool = False,
):
    """Handle VBA export with user confirmation for warnings.

    This helper wraps the export_vba() call and handles VBAExportWarning exceptions
    by prompting the user for confirmation. This centralizes the warning handling
    logic that would otherwise be duplicated across all CLI entry points.

    Args:
        handler: The VBA handler instance (WordVBAHandler, ExcelVBAHandler, etc.)
        save_metadata: Whether to save metadata file
        overwrite: Whether to overwrite existing files
        interactive: Whether to show warnings and prompt for confirmation
        force_overwrite: If True, skip all confirmation prompts (use with caution)
        keep_open: If True, keep document open after export (default: False = close)

    Returns:
        None

    Raises:
        SystemExit: If user cancels the export or an error occurs
    """
    # If force_overwrite is set, skip all interactive prompts
    if force_overwrite:
        logger.info("--force-overwrite flag is set: skipping all confirmation prompts")
        interactive = False

    try:
        handler.export_vba(
            save_metadata=save_metadata, overwrite=overwrite, interactive=interactive, keep_open=keep_open
        )
    except VBAExportWarning as warning_exc:
        if warning_exc.warning_type == "existing_files":
            file_count = warning_exc.context["file_count"]
            warning(f"Found {file_count} existing VBA file(s) in the VBA directory.")
            info("Continuing will overwrite these files with content from the document.")
            if not confirm_action("Do you want to continue?", default=False):
                info("Export cancelled by user.")
                import sys

                sys.exit(0)
            # User confirmed, retry with interactive=False to skip further prompts
            handler.export_vba(save_metadata=save_metadata, overwrite=True, interactive=False, keep_open=keep_open)

        elif warning_exc.warning_type == "header_mode_changed":
            old_mode = warning_exc.context["old_mode"]
            new_mode = warning_exc.context["new_mode"]
            warning(f"Header storage mode has changed from {old_mode} to {new_mode}.")
            info("Continuing will re-export all forms and clean up old .header files if needed.")
            if not confirm_action("Do you want to continue?", default=True):
                info("Export cancelled by user.")
                import sys

                sys.exit(0)
            # User confirmed, retry with overwrite=True and interactive=False
            handler.export_vba(save_metadata=save_metadata, overwrite=True, interactive=False, keep_open=keep_open)


# TOML configuration section constants
CONFIG_SECTION_GENERAL = "general"
CONFIG_SECTION_OFFICE = "office"
CONFIG_SECTION_EXCEL = "excel"
CONFIG_SECTION_WORD = "word"
CONFIG_SECTION_ACCESS = "access"
CONFIG_SECTION_POWERPOINT = "powerpoint"
CONFIG_SECTION_ADVANCED = "advanced"

# TOML configuration keys (general section) and argument namespace attributes
CONFIG_KEY_FILE = "file"
CONFIG_KEY_VBA_DIRECTORY = "vba_directory"
CONFIG_KEY_PQ_DIRECTORY = "pq_directory"
CONFIG_KEY_ENCODING = "encoding"
CONFIG_KEY_DETECT_ENCODING = "detect_encoding"
CONFIG_KEY_SAVE_HEADERS = "save_headers"
CONFIG_KEY_IN_FILE_HEADERS = "in_file_headers"
CONFIG_KEY_SAVE_METADATA = "save_metadata"
CONFIG_KEY_VERBOSE = "verbose"
CONFIG_KEY_LOGFILE = "logfile"
CONFIG_KEY_RUBBERDUCK_FOLDERS = "rubberduck_folders"
CONFIG_KEY_INVISIBLE_MODE = "invisible_mode"
CONFIG_KEY_OPEN_FOLDER = "open_folder"
CONFIG_KEY_FORCE_OVERWRITE = "force_overwrite"
CONFIG_KEY_KEEP_OPEN = "keep_open"
CONFIG_KEY_NO_COLOR = "no_color"
CONFIG_KEY_XLWINGS = "xlwings"

# Placeholder constants for use in arguments and configuration values
PLACEHOLDER_CONFIG_PATH = "{config.path}"
PLACEHOLDER_FILE_NAME = "{file.name}"
PLACEHOLDER_FILE_FULLNAME = "{file.fullname}"
PLACEHOLDER_FILE_PATH = "{file.path}"
PLACEHOLDER_FILE_VBAPROJECT = "{file.vbaproject}"

# Valid placeholders by config key (exported for other modules)
CONFIG_KEY_PLACEHOLDERS: Dict[str, tuple[str, ...]] = {
    CONFIG_KEY_VBA_DIRECTORY: (
        PLACEHOLDER_FILE_NAME,
        PLACEHOLDER_FILE_FULLNAME,
        PLACEHOLDER_FILE_PATH,
        PLACEHOLDER_CONFIG_PATH,
        PLACEHOLDER_FILE_VBAPROJECT,
    ),
    CONFIG_KEY_LOGFILE: (
        PLACEHOLDER_FILE_NAME,
        PLACEHOLDER_FILE_FULLNAME,
        PLACEHOLDER_FILE_PATH,
        PLACEHOLDER_CONFIG_PATH,
    ),
}


def get_placeholders_for_config_key(config_key: str) -> tuple[str, ...]:
    """Return the valid placeholders for a given config key."""
    return CONFIG_KEY_PLACEHOLDERS.get(config_key, ())


# Office application CLI configurations
OFFICE_CLI_CONFIGS = {
    "excel": {
        "entry_point": "excel-vba",
        "app_name": "Excel",
        "file_type": "workbook",
        "file_extensions": "*.bas/*.cls/*.frm",
        "example_filename": "workbook.xlsm",
    },
    "word": {
        "entry_point": "word-vba",
        "app_name": "Word",
        "file_type": "document",
        "file_extensions": "*.bas/*.cls/*.frm",
        "example_filename": "document.docx",
    },
    "access": {
        "entry_point": "access-vba",
        "app_name": "MS Access",
        "file_type": "database",
        "file_extensions": "*.bas/*.cls",  # Access only supports modules and class modules
        "example_filename": "database.accdb",
    },
    "powerpoint": {
        "entry_point": "powerpoint-vba",
        "app_name": "PowerPoint",
        "file_type": "presentation",
        "file_extensions": "*.bas/*.cls/*.frm",
        "example_filename": "presentation.pptx",
    },
}

# Centralized help strings
CLI_HELP_STRINGS = {
    "edit": "Edit VBA in external editor, sync changes back to {file_type}",
    "import": "Import VBA from filesystem into {file_type}",
    "export": "Export VBA from {file_type} to filesystem",
    "check": "Check VBA project access settings in {app_name}",
}


def resolve_placeholders_in_value(value: str, placeholders: Dict[str, str]) -> str:
    """Resolve placeholders in a single string value.

    Args:
        value: String that may contain placeholders
        placeholders: Dictionary mapping placeholder names to values

    Returns:
        String with placeholders resolved
    """
    if not isinstance(value, str):
        return value

    resolved_value = value
    for placeholder, replacement in placeholders.items():
        if replacement:  # Only replace if we have a value
            resolved_value = resolved_value.replace(placeholder, replacement)

    return resolved_value


def get_placeholder_values(config_file_path: Optional[str] = None, file_path: Optional[str] = None) -> Dict[str, str]:
    """Get placeholder values based on config file and file paths.

    Args:
        config_file_path: Path to the TOML config file (optional)
        file_path: Path to the Office document (optional)

    Returns:
        Dictionary mapping placeholder names to their values
    """
    placeholders = {
        PLACEHOLDER_CONFIG_PATH: "",
        PLACEHOLDER_FILE_NAME: "",
        PLACEHOLDER_FILE_FULLNAME: "",
        PLACEHOLDER_FILE_PATH: "",
        PLACEHOLDER_FILE_VBAPROJECT: "",  # Resolved later
    }

    # Get config file directory for relative path resolution
    if config_file_path:
        config_dir = Path(config_file_path).parent
        placeholders[PLACEHOLDER_CONFIG_PATH] = str(config_dir)

    # Extract file information if file path is available
    if file_path:
        # Handle case where file_path might contain unresolved placeholders
        if "{" not in file_path:  # Only process if no placeholders remain
            resolved_file_path = Path(file_path)

            # If relative path and we have config directory, resolve relative to config
            if not resolved_file_path.is_absolute() and config_file_path:
                config_dir = Path(config_file_path).parent
                resolved_file_path = config_dir / file_path

            file_name = resolved_file_path.stem  # filename without extension
            file_fullname = resolved_file_path.name  # filename with extension
            file_path_str = str(resolved_file_path.parent)

            placeholders[PLACEHOLDER_FILE_NAME] = file_name
            placeholders[PLACEHOLDER_FILE_FULLNAME] = file_fullname
            placeholders[PLACEHOLDER_FILE_PATH] = file_path_str

    placeholders = _placeholder_compatibility_stub(placeholders)
    return placeholders



def resolve_all_placeholders(args: argparse.Namespace, config_file_path: Optional[str] = None) -> argparse.Namespace:
    """Resolve all placeholders in arguments after config and CLI have been merged.

    Args:
        args: Command-line arguments namespace with merged config values
        config_file_path: Path to config file if one was used

    Returns:
        Updated arguments with placeholders resolved
    """
    args_dict = vars(args).copy()

    # Get file path from args for placeholder resolution
    file_path = args_dict.get("file")
    # Get placeholder values

    placeholders = get_placeholder_values(config_file_path, file_path)

    # Resolve placeholders only in arguments that support them
    placeholder_enabled_args = {CONFIG_KEY_VBA_DIRECTORY, CONFIG_KEY_LOGFILE}
    for key, value in args_dict.items():
        if key in placeholder_enabled_args and isinstance(value, str):
            args_dict[key] = resolve_placeholders_in_value(value, placeholders)

    # Store config file path for later VBA project placeholder resolution
    if config_file_path:
        args_dict["_config_file_path"] = config_file_path

    return argparse.Namespace(**args_dict)


def _enhance_toml_error_message(config_path: str, text: str, err: Exception) -> str:
    """Produce a helpful error for common Windows path mistakes in TOML."""
    # Base message with location if available
    base = str(err)
    if hasattr(err, "lineno") and hasattr(err, "colno"):
        base = f"{base} (at line {getattr(err, 'lineno', None)}, column {getattr(err, 'colno', None)})"

    # Look for suspicious backslashes in double-quoted values of known path keys
    keys = (CONFIG_KEY_FILE, CONFIG_KEY_VBA_DIRECTORY, CONFIG_KEY_PQ_DIRECTORY, CONFIG_KEY_LOGFILE)
    pattern = re.compile(
        r"^(\s*(?:" + "|".join(re.escape(k) for k in keys) + r')\s*=\s*)"([^"\r\n]*)"',
        re.IGNORECASE | re.MULTILINE,
    )

    hints = []
    for m in pattern.finditer(text):
        key, val = m.group(1).strip().split("=")[0].strip(), m.group(2)
        if "\\" in val:
            hints.append(f"- {key} has backslashes in a double-quoted string: {val!r}")

    if hints:
        guidance = (
            "TOML basic strings treat backslashes as escapes. For Windows paths, use one of:\n"
            "- Literal string (single quotes): 'C:\\Users\\me\\doc.xlsm'\n"
            '- Escaped backslashes: "C:\\\\Users\\\\me\\\\doc.xlsm"\n'
            '- Forward slashes: "C:/Users/me/doc.xlsm"\n'
            "Spec: https://toml.io/en/v1.0.0#string"
        )
        return (
            f"Failed to load config '{config_path}': {base}\nPossible issues:\n" + "\n".join(hints) + "\n\n" + guidance
        )

    return f"Failed to load config '{config_path}': {base}"


def load_config_file(config_path: str) -> Dict[str, Any]:
    """Load configuration from a TOML file.

    Args:
        config_path: Path to the TOML configuration file

    Returns:
        Dictionary containing the configuration

    Raises:
        FileNotFoundError: If the configuration file doesn't exist
        ValueError: If the configuration file isn't valid TOML
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    text = Path(config_path).read_text(encoding="utf-8")
    try:
        # Use loads() so we can re-use the same text for better diagnostics
        return toml_lib.loads(text)
    except Exception as e:
        # Raise a clear message explaining how to write Windows paths in TOML
        raise ValueError(_enhance_toml_error_message(config_path, text, e)) from e


def merge_config_with_args(args: argparse.Namespace, config: Dict[str, Any]) -> argparse.Namespace:
    """Merge configuration from a file with command-line arguments.

    Command-line arguments take precedence over configuration file values.
    Configuration structure is preserved (e.g., general.file remains as nested structure).

    Args:
        args: Command-line arguments
        config: Configuration from file

    Returns:
        Updated arguments with values from configuration
    """
    # Create a copy of the args namespace as a dictionary
    args_dict = vars(args).copy()

    # Handle 'general' section - these map directly to CLI args
    if CONFIG_SECTION_GENERAL in config:
        general_config = config[CONFIG_SECTION_GENERAL]
        for key, value in general_config.items():
            # Convert dashes to underscores for argument names
            arg_key = key.replace("-", "_")

            # Only update if the arg wasn't explicitly set (is None)
            if arg_key in args_dict and args_dict[arg_key] is None:
                args_dict[arg_key] = value

    # Store the full config for later access by handlers if needed
    args_dict["_config"] = config
    args_dict["_config_file_path"] = getattr(args, "_config_file_path", None)

    # Convert back to a Namespace
    return argparse.Namespace(**args_dict)


def process_config_file(args: argparse.Namespace) -> argparse.Namespace:
    """Load configuration file if specified and merge with command-line arguments.
    Also resolves placeholders in both config and CLI arguments.

    Args:
        args: Command-line arguments

    Returns:
        Updated arguments with values from configuration file and placeholders resolved
    """
    config_file_path = None

    # Process config file if specified
    if hasattr(args, "conf") and args.conf:
        config_file_path = args.conf

        try:
            config = load_config_file(config_file_path)
            args = merge_config_with_args(args, config)
        except Exception as e:
            error(f"Error loading configuration file: {e}")
            return args

    # Resolve all placeholders once after merging (except {vbaproject})
    args = resolve_all_placeholders(args, config_file_path)

    return args


def add_config_arguments(parser: argparse.ArgumentParser) -> None:
    """Add configuration file arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    config_group = parser.add_argument_group("Configuration")
    config_group.add_argument(
        "--conf",
        "--config",
        metavar="CONFIG_FILE",
        help=(
            "Path to configuration file (TOML format)\n"
            "Command-line arguments override config file values. "
            "Values in the configuration file support the same placeholders as the corresponding command-line arguments."
        ),
    )


def add_command_arguments(parser: argparse.ArgumentParser) -> None:
    """Add common command arguments to a parser.

    These are arguments common to edit/import/export commands.
    Global options (--version, --help) are added at the main parser level.

    Args:
        parser: The argument parser to add arguments to
    """
    # File Options group
    file_group = parser.add_argument_group("File Options")
    file_group.add_argument(
        "--file",
        "-f",
        dest=CONFIG_KEY_FILE,
        help="Path to Office document (default: active document).",
    )
    file_group.add_argument(
        "--vba-directory",
        dest=CONFIG_KEY_VBA_DIRECTORY,
        metavar="DIR",
        help="Directory to export VBA files to (default: same directory as document).\n"
        f"Supports placeholders: {', '.join(get_placeholders_for_config_key(CONFIG_KEY_VBA_DIRECTORY))}",
    )


def add_vba_files_arguments(parser: argparse.ArgumentParser) -> None:
    """Add folder organization and encoding/decodingarguments to a parser.

    These arguments only make sense for commands that handle VBA code
    (edit, export, import) and should not be available globally.

    Args:
        parser: The argument parser to add arguments to
    """
    add_encoding_arguments(parser)
    vba_src_file_group = parser.add_argument_group("Source File Organization")
    vba_src_file_group.add_argument(
        "--rubberduck-folders",
        dest=CONFIG_KEY_RUBBERDUCK_FOLDERS,
        action="store_true",
        help="Organize folders per RubberduckVBA @Folder annotations",
    )


def add_exporting_arguments(parser: argparse.ArgumentParser) -> None:
    """Add export-specific arguments to a parser.

    These arguments only make sense for commands that export VBA code
    (edit, export) and should not be available globally.

    Args:
        parser: The argument parser to add arguments to
    """
    add_header_arguments(parser)
    add_metadata_arguments(parser)
    exporting_group = parser.add_argument_group("Export Options")
    exporting_group.add_argument(
        "--force-overwrite",
        dest=CONFIG_KEY_FORCE_OVERWRITE,
        action="store_true",
        help="Force overwrite of existing files without prompting",
    )
    exporting_group.add_argument(
        "--open-folder",
        dest=CONFIG_KEY_OPEN_FOLDER,
        action="store_true",
        help="Open export directory in file explorer after export",
    )


def add_after_export_arguments(parser: argparse.ArgumentParser) -> None:
    """Add office file-specific arguments to a parser.

    These arguments only make sense for the export command

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument(
        "--keep-open",
        dest=CONFIG_KEY_KEEP_OPEN,
        action="store_true",
        help="Keep document open after export (default: close after export)",
    )


def add_common_option_group(parser: argparse.ArgumentParser) -> argparse._ArgumentGroup:
    """Add common CLI options to a parser as an organized group.

    This helper reduces boilerplate by providing a consistent way to add
    common options (verbose, logfile, no-color, help) across all commands.

    Use this in subparser setup to avoid repeating the same argument definitions.
    This creates a "Common Options" group with the standard set of options.

    Args:
        parser: The argument parser to add the group to

    Returns:
        The argument group that was created (for further customization if needed)

    Example:
        >>> import_parser = subparsers.add_parser("import")
        >>> add_common_option_group(import_parser)
    """
    # Common Options group
    common_group = parser.add_argument_group("Common Options")
    common_group.add_argument(
        "--verbose",
        "-v",
        dest=CONFIG_KEY_VERBOSE,
        action="store_true",
        help="Enable verbose logging output",
    )
    common_group.add_argument(
        "--logfile",
        "-l",
        dest=CONFIG_KEY_LOGFILE,
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file (default: vba_edit.log)\n"
        f"Supports placeholders: {', '.join(get_placeholders_for_config_key(CONFIG_KEY_LOGFILE))}",
    )
    common_group.add_argument(
        "--no-color",
        "--no-colour",
        dest=CONFIG_KEY_NO_COLOR,
        action="store_true",
        help="Disable colored output",
    )
    common_group.add_argument(
        "--help",
        "-h",
        action="help",
        help="Show this help message and exit",
    )


def add_excel_specific_arguments(parser: argparse.ArgumentParser) -> None:
    """Add Excel-specific arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    excel_group = parser.add_argument_group("Excel-Specific Options")
    excel_group.add_argument(
        "--xlwings",
        "-x",
        dest=CONFIG_KEY_XLWINGS,
        action="store_true",
        help="""Use xlwings backend with vba-edit enhancements. Adds custom --vba-directory 
        support (automatically changes directory and creates it if needed). Useful for 
        comparing implementations, validation testing, or as fallback option. 
        Note: Advanced features (headers, Rubberduck folders, config files) require native mode.""",
    )


def add_encoding_arguments(parser: argparse.ArgumentParser) -> None:
    """Add encoding-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    encoding_group = parser.add_argument_group("Encoding Options (mutually exclusive)")
    encoding_mutex = encoding_group.add_mutually_exclusive_group()
    encoding_mutex.add_argument(
        "--encoding",
        "-e",
        dest=CONFIG_KEY_ENCODING,
        metavar="ENCODING",
        default=default_encoding,
        help=f"Encoding used for reading/writing VBA files (e.g. 'utf-8', 'windows-1252', default: {default_encoding})",
    )
    encoding_mutex.add_argument(
        "--detect-encoding",
        "-d",
        dest=CONFIG_KEY_DETECT_ENCODING,
        action="store_true",
        help="Auto-detect file encoding for VBA files",
    )


def add_header_arguments(parser: argparse.ArgumentParser) -> None:
    """Add header-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    header_group = parser.add_argument_group("Header Options (mutually exclusive)")
    header_mutex = header_group.add_mutually_exclusive_group()
    header_mutex.add_argument(
        "--save-headers",
        dest=CONFIG_KEY_SAVE_HEADERS,
        action="store_true",
        help="Save VBA component headers to separate .header files",
    )
    header_mutex.add_argument(
        "--in-file-headers",
        dest=CONFIG_KEY_IN_FILE_HEADERS,
        action="store_true",
        help="Include VBA headers directly in code files",
    )


def validate_header_options(args: argparse.Namespace) -> None:
    """Validate that header options are not conflicting."""
    if getattr(args, CONFIG_KEY_SAVE_HEADERS, False) and getattr(args, CONFIG_KEY_IN_FILE_HEADERS, False):
        raise argparse.ArgumentTypeError(
            "Options --save-headers and --in-file-headers are mutually exclusive. "
            "Choose either separate header files or embedded headers."
        )


def add_metadata_arguments(parser: argparse.ArgumentParser) -> None:
    """Add metadata-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    metadata_group = parser.add_argument_group("Metadata Options")
    metadata_group.add_argument(
        "--save-metadata",
        "-m",
        dest=CONFIG_KEY_SAVE_METADATA,
        action="store_true",
        help="Save metadata to file",
    )


def get_office_config(office_app: str) -> Dict[str, str]:
    """Get configuration for specified Office application.

    Args:
        office_app: Office application name (excel, word, access, powerpoint)

    Returns:
        Configuration dictionary

    Raises:
        KeyError: If office_app is not supported
    """
    if office_app not in OFFICE_CLI_CONFIGS:
        raise KeyError(f"Unsupported Office application: {office_app}")
    return OFFICE_CLI_CONFIGS[office_app]


def get_help_string(command: str, office_app: str) -> str:
    """Get help string for a command and Office application.

    Args:
        command: Command name (edit, import, export, check)
        office_app: Office application name

    Returns:
        Formatted help string
    """
    config = get_office_config(office_app)
    template = CLI_HELP_STRINGS.get(command, f"{command.title()} VBA content")
    return template.format(**config)


def get_command_description(command: str, office_app: str) -> str:
    """Get description string for a command and Office application.

    Args:
        command: Command name (edit, import, export, check)
        office_app: Office application name
    Returns:
        Formatted description string
    """
    config = get_office_config(office_app)

    match command:
        case "import":
            command_description = f"""Import VBA code from filesystem into {config["file_type"]}

Header handling is automatic - no flags needed:
  • Detects inline headers (VERSION/BEGIN/Attribute at file start)
  • Falls back to separate .header files if present
  • Creates minimal headers if neither exists

Simple usage:
  {config["entry_point"]} {command}          # Uses active {config["file_type"]} and imports from same directory

Full control usage:
  {config["entry_point"]} {command} -f {config["example_filename"]} --vba-directory src"""

        case "export":
            command_description = f"""Export VBA code from {config["file_type"]} to filesystem

Simple usage:
  {config["entry_point"]} {command}           # Uses active {config["file_type"]} and exports to same directory
  
Full control usage:
  {config["entry_point"]} {command} -f {config["example_filename"]} --vba-directory src"""

        case "edit":
            command_description = f"""Export VBA code from {config["file_type"]} to filesystem, edit in external editor and sync changes back into {config["file_type"]} on save [CTRL+S]

Simple usage:
  {config["entry_point"]} {command}           # Uses active {config["file_type"]} and VBA code files are 
                            # saved in the same location as the {config["file_type"]}
  
Full control usage:
  {config["entry_point"]} {command} -f {config["example_filename"]} --vba-directory src"""

        case "check":
            command_description = f"""Check if 'Trust access to the VBA project object model' is enabled

Simple usage:
  {config["entry_point"]} {command}          # Check {config["app_name"]} VBA access
  {config["entry_point"]} {command} all      # Check all Office applications"""
        case _:
            command_description = "dummy"

    return command_description


def get_command_usage(command: str, office_app: str) -> str:
    """Get usage string for a command and Office application.

    Args:
        command: Command name (edit, import, export, check)
        office_app: Office application name
    Returns:
        Formatted usage string
    """
    common_command_options1 = """
    [--file FILE | -f FILE]
    [--vba-directory DIR]"""

    edit_and_export_options = """
    [--save-headers | --in-file-headers]
    [--save-metadata | -m]
    [--open-folder]
    [--force-overwrite]"""

    after_export_options = """
    [--keep-open]"""

    vba_handling_options = """
    [--encoding ENCODING | -e ENCODING | --detect-encoding | -d]
    [--rubberduck-folders]"""

    config_options = """
    [--conf FILE | --config FILE]"""

    app_excel_specific_options = """
    [--xlwings | -x]"""

    common_command_options2 = """
    [--verbose | -v]
    [--logfile | -l]
    [--no-color | --no-colour]
    [--help | -h]"""

    config = get_office_config(office_app)

    command_usage = f"{config['entry_point']} {command}"
    match command:
        case "export" | "edit" | "import":
            command_usage = f"{command_usage}{common_command_options1}"

            if command in ("edit", "export"):
                command_usage = f"{command_usage}{edit_and_export_options}"

            if command == "export":
                command_usage = f"{command_usage}{after_export_options}"

            command_usage = f"{command_usage}{vba_handling_options}"
            if office_app == "excel":
                command_usage = f"{command_usage}{app_excel_specific_options}"
            command_usage = f"{command_usage}{config_options}"

    command_usage = f"{command_usage}{common_command_options2}"

    return command_usage


EXAMPLE_FILENAME_MAX_LEN = 24
EXAMPLE_COMMAND_COL_WIDTH = 46


def _cap_example_filename(example_filename: str, max_len: int = EXAMPLE_FILENAME_MAX_LEN) -> str:
    """Cap example filename length to keep example alignment stable."""
    if len(example_filename) <= max_len:
        return example_filename
    return example_filename[: max_len - 1] + "…"


def _format_example_line(command: str, comment: str, col_width: int = EXAMPLE_COMMAND_COL_WIDTH) -> str:
    """Format example line with aligned comments."""
    return f"{command.ljust(col_width)}# {comment}"


def create_cli_description(
    entry_point_name: str,
    package_name_formatted: str,
    package_version: str,
    app_name: str,
    file_type: str,
    file_extensions: str,
    example_filename: str,
) -> str:
    """Create standardized CLI description for Office VBA tools."""
    return f"""{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool for managing VBA content in {app_name} {file_type}s.
Export, import, and edit VBA code in external editor with live sync back to the {file_type}."""


def create_cli_examples(
    entry_point_name: str,
    package_name_formatted: str,
    package_version: str,
    app_name: str,
    file_type: str,
    file_extensions: str,
    example_filename: str,
) -> str:
    """Create standardized CLI description for Office VBA tools."""
    example_filename = _cap_example_filename(example_filename)

    edit_cmd = f"{entry_point_name} edit"
    export_cmd = f"{entry_point_name} export -f {example_filename}"
    import_cmd = f"{entry_point_name} import --vba-directory src"
    check_cmd = f"{entry_point_name} check"

    return f"""
Examples:
  {_format_example_line(edit_cmd, f"Edit in external editor, sync changes to {file_type}")}
  {_format_example_line(export_cmd, "Export VBA to directory")}
  {_format_example_line(import_cmd, "Import VBA from directory")}
  {_format_example_line(check_cmd, "Verify that access to VBA is enabled")}

Use '{entry_point_name} <command> --help' for more information on a specific command.
    
IMPORTANT: Requires "Trust access to the VBA project object model" enabled in {app_name}.
           Early release - backup important files before use!"""


def create_office_cli_description(office_app: str, package_name_formatted: str, package_version: str) -> str:
    """Create CLI description for specified Office application.

    Args:
        office_app: Office application name (excel, word, access, powerpoint)
        package_name_formatted: Package name for display (e.g., "excel-vba")
        package_version: Version string

    Returns:
        Formatted description string
    """
    config = get_office_config(office_app)
    return create_cli_description(
        entry_point_name=config["entry_point"],
        package_name_formatted=package_name_formatted,
        package_version=package_version,
        app_name=config["app_name"],
        file_type=config["file_type"],
        file_extensions=config["file_extensions"],
        example_filename=config["example_filename"],
    )


def create_office_cli_examples(office_app: str, package_name_formatted: str, package_version: str) -> str:
    """Create CLI epilog for specified Office application.

    Args:
        office_app: Office application name (excel, word, access, powerpoint)
        package_name_formatted: Package name for display (e.g., "excel-vba")
        package_version: Version string

    Returns:
        Formatted epilog string
    """
    config = get_office_config(office_app)
    return create_cli_examples(
        entry_point_name=config["entry_point"],
        package_name_formatted=package_name_formatted,
        package_version=package_version,
        app_name=config["app_name"],
        file_type=config["file_type"],
        file_extensions=config["file_extensions"],
        example_filename=config["example_filename"],
    )

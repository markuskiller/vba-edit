"""Common CLI argument definitions for all Office VBA handlers."""

import argparse
import logging
import re
import os
from pathlib import Path
from typing import Dict, Any, Optional
from vba_edit.utils import get_windows_ansi_codepage

# Prefer stdlib tomllib (Py 3.11+), fallback to tomli for older envs
try:
    import tomllib as toml_lib  # Python 3.11+ includes tomllib in stdlib
except ImportError:
    try:
        import tomli as toml_lib  # tomli is the recommended TOML parser for Python <3.11
    except ImportError:
        import toml as toml_lib  # Fall back to toml package if tomli isn't available

logger = logging.getLogger(__name__)

# Placeholder constants
PLACEHOLDER_CONFIG_PATH = "{config.path}"
PLACEHOLDER_FILE_NAME = "{general.file.name}"
PLACEHOLDER_FILE_FULLNAME = "{general.file.fullname}"
PLACEHOLDER_FILE_PATH = "{general.file.path}"
PLACEHOLDER_VBA_PROJECT = "{vbaproject}"

# TOML configuration section constants
CONFIG_SECTION_GENERAL = "general"
CONFIG_SECTION_OFFICE = "office"
CONFIG_SECTION_EXCEL = "excel"
CONFIG_SECTION_WORD = "word"
CONFIG_SECTION_ACCESS = "access"
CONFIG_SECTION_POWERPOINT = "powerpoint"
CONFIG_SECTION_ADVANCED = "advanced"

# TOML configuration key constants (for general section)
CONFIG_KEY_FILE = "file"
CONFIG_KEY_VBA_DIRECTORY = "vba_directory"
CONFIG_KEY_PQ_DIRECTORY = "pq_directory"
CONFIG_KEY_ENCODING = "encoding"
CONFIG_KEY_DETECT_ENCODING = "detect_encoding"
CONFIG_KEY_SAVE_HEADERS = "save_headers"
CONFIG_KEY_VERBOSE = "verbose"
CONFIG_KEY_LOGFILE = "logfile"
CONFIG_KEY_RUBBERDUCK_FOLDERS = "rubberduck_folders"
CONFIG_KEY_INVISIBLE_MODE = "invisible_mode"
CONFIG_KEY_MODE = "mode"
CONFIG_KEY_OPEN_FOLDER = "open_folder"


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
        # {vbaproject} will be resolved later when we have access to the Office file
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

            placeholders[PLACEHOLDER_FILE_NAME] = resolved_file_path.stem  # filename without extension
            placeholders[PLACEHOLDER_FILE_FULLNAME] = resolved_file_path.name  # filename with extension
            placeholders[PLACEHOLDER_FILE_PATH] = str(resolved_file_path.parent)

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

    # Resolve placeholders in all string arguments
    for key, value in args_dict.items():
        if isinstance(value, str):
            args_dict[key] = resolve_placeholders_in_value(value, placeholders)

    # Store config file path for later VBA project placeholder resolution
    if config_file_path:
        args_dict["_config_file_path"] = config_file_path

    return argparse.Namespace(**args_dict)


def resolve_vbaproject_placeholder_in_args(args: argparse.Namespace, vba_project_name: str) -> argparse.Namespace:
    """Resolve the {vbaproject} placeholder in arguments after VBA project name is known.

    Args:
        args: Command-line arguments
        vba_project_name: Name of the VBA project

    Returns:
        Arguments with {vbaproject} placeholder resolved
    """
    args_dict = vars(args).copy()

    # Resolve {vbaproject} placeholder in all string arguments
    for key, value in args_dict.items():
        if isinstance(value, str):
            args_dict[key] = value.replace(PLACEHOLDER_VBA_PROJECT, vba_project_name)

    return argparse.Namespace(**args_dict)


def resolve_config_placeholders_recursive(value, placeholders: Dict[str, str]):
    """Recursively resolve placeholders in nested configuration structures.

    Args:
        value: Value to process (can be dict, list, or string)
        placeholders: Dictionary mapping placeholder names to values

    Returns:
        Value with placeholders resolved
    """
    if isinstance(value, str):
        return resolve_placeholders_in_value(value, placeholders)
    elif isinstance(value, dict):
        return {k: resolve_config_placeholders_recursive(v, placeholders) for k, v in value.items()}
    elif isinstance(value, list):
        return [resolve_config_placeholders_recursive(item, placeholders) for item in value]
    else:
        return value


def resolve_vbaproject_placeholder(config: Dict[str, Any], vba_project_name: str) -> Dict[str, Any]:
    """Resolve the {vbaproject} placeholder after VBA project name is known.

    Args:
        config: Configuration dictionary
        vba_project_name: Name of the VBA project

    Returns:
        Configuration with {vbaproject} placeholder resolved
    """
    import copy

    resolved_config = copy.deepcopy(config)

    placeholders = {PLACEHOLDER_VBA_PROJECT: vba_project_name}

    return resolve_config_placeholders_recursive(resolved_config, placeholders)


def _enhance_toml_error_message(config_path: str, text: str, err: Exception) -> str:
    """Produce a helpful error for common Windows path mistakes in TOML."""
    # Base message with location if available
    base = str(err)
    if hasattr(err, "lineno") and hasattr(err, "colno"):
        base = f"{base} (at line {getattr(err, 'lineno', None)}, column {getattr(err, 'colno', None)})"

    # Look for suspicious backslashes in double-quoted values of known path keys
    keys = ("file", "vba_directory", "pq_directory", "logfile")
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
            print(f"Error loading configuration file: {e}")
            return args

    # Resolve all placeholders once after merging (except {vbaproject})
    args = resolve_all_placeholders(args, config_file_path)

    return args


def add_config_arguments(parser: argparse.ArgumentParser) -> None:
    """Add configuration file arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument(
        "--conf",
        "--config",
        metavar="CONFIG_FILE",
        help="Path to configuration file (TOML format) with argument values. "
        "Command-line arguments override config file values. "
        "Configuration values support placeholders: {config.path}, {general.file.name}, {general.file.fullname}, {general.file.path}, {vbaproject}",
    )


def add_common_arguments(parser: argparse.ArgumentParser) -> None:
    """Add common arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    add_config_arguments(parser)
    parser.add_argument(
        "--file",
        "-f",
        help="Path to Office document (optional, defaults to active document). "
        "Supports placeholders: {general.file.name}, {general.file.fullname}, {general.file.path}, {vbaproject}",
    )
    parser.add_argument(
        "--vba-directory",
        help="Directory to export VBA files to (optional, defaults to current directory) "
        "Supports placeholders: {general.file.name}, {general.file.fullname}, {general.file.path}, {vbaproject}",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging output")
    parser.add_argument(
        "--logfile",
        "-l",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file. Optional path can be specified (default: vba_edit.log)"
        "Supports placeholders: {general.file.name}, {general.file.fullname}, {general.file.path}, {vbaproject}",
    )
    parser.add_argument(
        "--rubberduck-folders",
        action="store_true",
        default=None,
        help="If a module contains a RubberduckVBA '@Folder annotation, organize folders in the file system accordingly",
    )
    parser.add_argument(
        "--open-folder",
        action="store_true",
        default=None,
        help="Open the export directory in file explorer after successful export",
    )


def add_excel_specific_arguments(parser: argparse.ArgumentParser) -> None:
    """Add Excel-specific arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument(
        "--xlwings", "-x", action="store_true", help="Use wrapper for xlwings vba edit|import|export commands"
    )


def add_encoding_arguments(parser: argparse.ArgumentParser, default_encoding: str = None) -> None:
    """Add encoding-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
        default_encoding: Default encoding to use
    """
    if default_encoding is None:
        default_encoding = get_windows_ansi_codepage() or "cp1252"

    encoding_group = parser.add_mutually_exclusive_group()
    encoding_group.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when reading/writing VBA files (e.g., 'utf-8', 'windows-1252', default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding", "-d", action="store_true", default=None, help="Auto-detect file encoding for VBA files"
    )


def add_header_arguments(parser: argparse.ArgumentParser) -> None:
    """Add header-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument(
        "--save-headers",
        action="store_true",
        default=None,
        help="Save VBA component headers to separate .header files (default: False)",
    )
    parser.add_argument(
        "--in-file-headers",
        action="store_true",
        default=None,
        help="Include VBA headers directly in code files instead of separate .header files (default: True)",
    )


def validate_header_options(args: argparse.Namespace) -> None:
    """Validate that header options are not conflicting."""
    if getattr(args, "save_headers", False) and getattr(args, "in_file_headers", False):
        raise argparse.ArgumentTypeError(
            "Options --save-headers and --in-file-headers are mutually exclusive. "
            "Choose either separate header files or embedded headers."
        )


def add_metadata_arguments(parser: argparse.ArgumentParser) -> None:
    """Add metadata-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument(
        "--save-metadata",
        "-m",
        action="store_true",
        default=None,
        help="Save metadata file with character encoding information (default: False)",
    )

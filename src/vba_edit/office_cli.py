"""Generic Office VBA CLI factory for creating standardized CLI interfaces.

This module provides a centralized, generic CLI factory for all Office VBA command-line tools (Excel, Word, Access, PowerPoint).

## Key Benefits
- **Consistent Behavior**: All Office tools behave identically with unified error handling
- **Easy Maintenance**: Bug fixes and new features automatically apply to all Office apps
- **Extensibility**: Adding new Office applications requires minimal effort
- **Clean Special Handling**: Office-specific features (xlwings, Access multi-DB) are isolated

## Architecture

The `OfficeVBACLI` class provides a generic implementation that can be configured for any
Office application. Special handling for unique features is managed through:

- **Hook Functions**: Pre-command processing (e.g., Access multi-database checks)
- **Handler Functions**: Special command processing (e.g., Excel xlwings integration)
- **Configuration**: Office-specific settings and arguments

## Usage

```python
# Create simplified Office modules
from vba_edit.office_cli import create_office_main

# Word VBA module becomes just:
main = create_office_main("word")

if __name__ == "__main__":
    main()
```

Other modules with special handling (like Excel) can still implement their unique features without affecting the overall structure or other Office applications.

## Supported Office Applications

- **Excel**: Full support including xlwings integration
- **Word**: Standard VBA operations
- **Access**: Standard operations with multi-database handling
- **PowerPoint**: Standard VBA operations

## Special Handling

### Excel-Specific Features
- xlwings wrapper integration (`--xlwings` flag)
- Excel-specific argument handling

### Access-Specific Features
- Multi-database detection and handling
- Database-specific warning messages

### Standard Features (Word, PowerPoint)
- No special handling required
"""

import argparse
import logging

import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.cli_common import (
    CONFIG_KEY_VBA_DIRECTORY,
    CONFIG_SECTION_GENERAL,
    PLACEHOLDER_FILE_VBAPROJECT,
    add_after_export_arguments,
    add_common_option_group,
    add_config_arguments,
    get_placeholders_for_config_key,
    handle_export_with_warnings,
    load_config_file,
    merge_config_with_args,
    resolve_all_placeholders,
    validate_header_options,
    get_command_usage,
    get_command_description,
    add_command_arguments,
    add_vba_files_arguments,
    add_exporting_arguments,
    add_excel_specific_arguments,
    create_office_cli_description,
    create_office_cli_examples,
    get_help_string,
    get_office_config,
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
from vba_edit.office_vba import (
    ExcelVBAHandler,
    WordVBAHandler,
    AccessVBAHandler,
    PowerPointVBAHandler,
)
from vba_edit.path_utils import get_document_paths
from vba_edit.utils import get_active_office_document, setup_logging
from vba_edit.console import error

from typing import Callable, Type


# Eager mapping (backwards-compatible; values are classes)
OFFICE_HANDLERS: dict[str, Type] = {
    "excel": ExcelVBAHandler,
    "word": WordVBAHandler,
    "access": AccessVBAHandler,
    "powerpoint": PowerPointVBAHandler,
}

# Lazy mapping of Office applications to their handler *factories*
# NOTE: Factories are used so tests can patch vba_edit.office_cli.ExcelVBAHandler
# and have it take effect without needing to patch a pre-built dict.
OFFICE_HANDLER_FACTORIES: dict[str, Callable[[], Type]] = {
    "excel": lambda: ExcelVBAHandler,
    "word": lambda: WordVBAHandler,
    "access": lambda: AccessVBAHandler,
    "powerpoint": lambda: PowerPointVBAHandler,
}

# Configure module logger
logger = logging.getLogger(__name__)


def _get_office_function(office_app: str, function_name: str):
    """Dynamically import and return an office-specific function.

    Args:
        office_app: Office application name (excel, word, access, powerpoint)
        function_name: Name of the function to import

    Returns:
        The requested function, or None if not found
    """
    try:
        if office_app == "excel":
            from vba_edit import excel_vba

            return getattr(excel_vba, function_name, None)
        elif office_app == "word":
            from vba_edit import word_vba

            return getattr(word_vba, function_name, None)
        elif office_app == "access":
            from vba_edit import access_vba

            return getattr(access_vba, function_name, None)
        elif office_app == "powerpoint":
            from vba_edit import powerpoint_vba

            return getattr(powerpoint_vba, function_name, None)
    except (ImportError, AttributeError):
        return None

    return None


# Office-specific configuration with dynamic function names
OFFICE_SPECIAL_HANDLING = {
    "access": {
        "pre_command_hook": "access_pre_command_hook",
        "extra_notes": ["NOTE: The database will remain open - close it manually when finished."],
    },
    "excel": {"xlwings_handler": "excel_xlwings_handler", "extra_arguments": add_excel_specific_arguments},
    "word": {},
    "powerpoint": {},
}


class OfficeVBACLI:
    """Generic Office VBA CLI that can be configured for any Office application."""

    def __init__(self, office_app: str):
        """Initialize CLI for specific Office application.

        Args:
            office_app: Office application name (excel, word, access, powerpoint)
        """
        self.office_app = office_app
        self.config = get_office_config(office_app)

        # Store a factory, not the class
        self._handler_factory = OFFICE_HANDLER_FACTORIES[office_app]

        self.special_config = OFFICE_SPECIAL_HANDLING.get(office_app, {})
        self.logger = logging.getLogger(f"{__name__}.{office_app}")

    def _get_special_function(self, function_key: str):
        """Get a special function for this office app."""
        function_name = self.special_config.get(function_key)
        if isinstance(function_name, str):
            return _get_office_function(self.office_app, function_name)
        return function_name

    def _collect_parser_dests(self, parser: argparse.ArgumentParser) -> set:
        dests = set()
        for action in parser._actions:
            if action.dest and action.dest is not argparse.SUPPRESS:
                dests.add(action.dest)
            if isinstance(action, argparse._SubParsersAction):
                for subparser in action.choices.values():
                    dests.update(self._collect_parser_dests(subparser))
        return dests

    def _get_subparser(self, parser: argparse.ArgumentParser, command: str) -> argparse.ArgumentParser | None:
        for action in parser._actions:
            if isinstance(action, argparse._SubParsersAction):
                return action.choices.get(command)
        return None

    def _get_config_defaults(self, parser: argparse.ArgumentParser, config: dict) -> dict:
        defaults = {}
        general_config = config.get(CONFIG_SECTION_GENERAL, {})
        if not isinstance(general_config, dict):
            return defaults

        known_dests = self._collect_parser_dests(parser)
        for key, value in general_config.items():
            arg_key = key.replace("-", "_")
            if arg_key in known_dests:
                defaults[arg_key] = value

        return defaults

    def create_cli_parser(self) -> argparse.ArgumentParser:
        """Create the command-line interface parser."""
        entry_point_name = self.config["entry_point"]
        package_name_formatted = package_name.replace("_", "-")

        # Create streamlined main help description
        main_description = create_office_cli_description(self.office_app, package_name_formatted, package_version)

        # Create streamlined examples for main help
        main_examples = create_office_cli_examples(self.office_app, package_name_formatted, package_version)

        parser = ColorizedArgumentParser(
            prog=entry_point_name,
            usage=f"{entry_point_name} [--help] [--version] <command> [<args>]",
            description=main_description,
            epilog=main_examples,
            formatter_class=EnhancedHelpFormatter,
        )

        # Add --version argument to the main parser
        parser.add_argument(
            "--version",
            "-V",
            action="version",
            version=f"{package_name_formatted} v{package_version} ({entry_point_name})",
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

        # Import command
        import_parser = subparsers.add_parser(
            "import",
            usage=get_command_usage("import", self.office_app),
            help=get_help_string("import", self.office_app),
            description=get_command_description("import", self.office_app),
            formatter_class=EnhancedHelpFormatter,
            add_help=False,
        )
        add_command_arguments(import_parser)
        add_vba_files_arguments(import_parser)
        add_config_arguments(import_parser)  # Add config file options to import command
        add_common_option_group(import_parser)  # Add common options to import command

        # Export command
        export_parser = subparsers.add_parser(
            "export",
            usage=get_command_usage("export", self.office_app),
            help=get_help_string("export", self.office_app),
            description=get_command_description("export", self.office_app),
            formatter_class=EnhancedHelpFormatter,
            add_help=False,
        )
        add_command_arguments(export_parser)
        add_exporting_arguments(export_parser)
        add_after_export_arguments(export_parser)
        add_vba_files_arguments(export_parser)
        add_config_arguments(export_parser)  # Add config file options to export command
        add_common_option_group(export_parser)  # Add common options to export command

        # Edit command
        edit_parser = subparsers.add_parser(
            "edit",
            usage=get_command_usage("edit", self.office_app),
            help=get_help_string("edit", self.office_app),
            description=get_command_description("edit", self.office_app),
            formatter_class=EnhancedHelpFormatter,
            add_help=False,  # Suppress default help to add it manually in Common Options
        )
        add_command_arguments(edit_parser)
        add_exporting_arguments(edit_parser)
        add_vba_files_arguments(edit_parser)
        add_config_arguments(edit_parser)  # Add config file options to edit command
        add_common_option_group(edit_parser)  # Add common options to edit command

        # Check command
        check_parser = subparsers.add_parser(
            "check",
            usage=get_command_usage("check", self.office_app),
            help=get_help_string("check", self.office_app),
            description=get_command_description("check", self.office_app),
            formatter_class=EnhancedHelpFormatter,
            add_help=False,
        )
        add_common_option_group(check_parser)  # Add common options to check command

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
        add_common_option_group(check_all_parser)  # Add common options to check all subcommand

        # Add application specific arguments
        extra_args_func = self._get_special_function("extra_arguments")
        if extra_args_func:
            extra_args_func(edit_parser)
            extra_args_func(import_parser)
            extra_args_func(export_parser)

        return parser

    def validate_paths(self, args: argparse.Namespace) -> None:
        """Validate file and directory paths from command line arguments.

        Only validates paths for commands that actually use them (edit, import, export).
        The 'check' command doesn't use file/vba_directory arguments.
        """
        # Skip validation for commands that don't use file paths
        if args.command == "check":
            return

        if args.file and not Path(args.file).exists():
            file_type = self.config["file_type"]
            raise FileNotFoundError(f"{file_type.title()} not found: {args.file}")

        if args.vba_directory:
            # Only create the VBA directory if there's no PLACEHOLDER_FILE_VBAPROJECT value, or if it is already resolved
            if PLACEHOLDER_FILE_VBAPROJECT not in args.vba_directory:
                vba_dir = Path(args.vba_directory)
                if not vba_dir.exists():
                    self.logger.info(f"Creating VBA directory: {vba_dir}")
                    vba_dir.mkdir(parents=True, exist_ok=True)

    def _call_handle_export_with_warnings(
        self,
        handler,
        args,
        *,
        overwrite: bool,
        interactive: bool,
        keep_open: bool = False,
        save_metadata: bool | None = None,
        force_overwrite: bool | None = None,
    ) -> None:
        """Call the module-level handle_export_with_warnings via runtime lookup.

        NOTE: This indirection ensures test patches applied to vba_edit.excel_vba.handle_export_with_warnings
        (and then synced into this module) are honored at call time.
        """
        module = sys.modules[__name__]
        module.handle_export_with_warnings(
            handler,
            save_metadata=getattr(args, "save_metadata", False) if save_metadata is None else save_metadata,
            overwrite=overwrite,
            interactive=interactive,
            force_overwrite=getattr(args, "force_overwrite", False) if force_overwrite is None else force_overwrite,
            keep_open=keep_open,
        )

    def handle_office_vba_command(self, args: argparse.Namespace) -> None:
        """Handle the office-vba command execution."""
        try:
            # Initialize logging
            setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
            self.logger.debug(f"Starting {self.office_app}-vba command: {args.command}")
            self.logger.debug(f"Command arguments: {vars(args)}")

            # Ensure paths exist early (creates vba_directory if provided)
            self.validate_paths(args)

            # Run application-specific pre-command hook
            pre_hook = self._get_special_function("pre_command_hook")
            if pre_hook:
                pre_hook(args)

            # Handle xlwings option if present (Excel only)
            xlwings_handler = self._get_special_function("xlwings_handler")
            if xlwings_handler and xlwings_handler(self, args):
                return  # xlwings handled the command

            # Get document path and active document path
            active_file = None
            if not args.file:
                try:
                    active_file = get_active_office_document(self.office_app)
                except ApplicationError:
                    pass

            try:
                file_path, vba_dir = get_document_paths(
                    args.file,
                    active_file,
                    args.vba_directory,
                    get_placeholders_for_config_key(CONFIG_KEY_VBA_DIRECTORY),
                )
                file_type = self.config["file_type"]
                self.logger.info(f"Using {file_type}: {file_path}")
                self.logger.debug(f"Using VBA directory: {vba_dir}")
            except (DocumentNotFoundError, PathError) as e:
                self.logger.error(f"Failed to resolve paths: {str(e)}")
                sys.exit(1)

            # Determine encoding
            encoding = None if getattr(args, "detect_encoding", False) else args.encoding
            self.logger.debug(f"Using encoding: {encoding or 'auto-detect'}")

            # Validate header options
            validate_header_options(args)

            # Create handler instance
            try:
                handler_class = self._handler_factory()
                handler = handler_class(
                    doc_path=str(file_path),
                    vba_dir=str(vba_dir),
                    encoding=encoding,
                    verbose=getattr(args, "verbose", False),
                    save_headers=getattr(args, "save_headers", False),
                    use_rubberduck_folders=getattr(args, "rubberduck_folders", False),
                    open_folder=getattr(args, "open_folder", False),
                    in_file_headers=getattr(args, "in_file_headers", True),
                )
            except VBAError as e:
                app_name = self.config["app_name"]
                self.logger.error(f"Failed to initialize {app_name} VBA handler: {str(e)}")
                sys.exit(1)

            # Execute requested command
            self.logger.info(f"Executing command: {args.command}")
            try:
                if args.command == "edit":
                    print("NOTE: Deleting a VBA module file will also delete it in the VBA editor!")

                    # Add office-specific notes
                    extra_notes = self.special_config.get("extra_notes", [])
                    for note in extra_notes:
                        print(note)
                    # region Indirection to ensure test patches to handle_export_with_warnings are honored
                    # once test_excel_vba_cli.py::test_save_metadata_passed_to_handler_edit is adapted, this can be simplified to a direct call to handle_export_with_warnings
                    self._call_handle_export_with_warnings(
                        handler,
                        args,
                        overwrite=False,
                        interactive=True,
                        keep_open=True,  # CRITICAL: Must keep document open for edit mode
                    )
                    # endregion
                    try:
                        handler.watch_changes()
                    except (DocumentClosedError, RPCError) as e:
                        self.logger.error(str(e))
                        app_name = self.config["app_name"]
                        self.logger.info(
                            f"Edit session terminated. Please restart {app_name} and this tool to continue editing."
                        )
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
                self.logger.error(str(e))
                sys.exit(1)
            except VBAAccessError as e:
                self.logger.error(str(e))
                app_name = self.config["app_name"]
                self.logger.error(f"Please check {app_name} Trust Center Settings and try again.")
                sys.exit(1)
            except VBAError as e:
                self.logger.error(f"VBA operation failed: {str(e)}")
                sys.exit(1)
            except Exception as e:
                self.logger.error(f"Unexpected error: {str(e)}")
                if getattr(args, "verbose", False):
                    self.logger.exception("Detailed error information:")
                sys.exit(1)

        except KeyboardInterrupt:
            self.logger.info("\nOperation interrupted by user")
            sys.exit(0)
        except Exception as e:
            self.logger.error(f"Critical error: {str(e)}")
            if getattr(args, "verbose", False):
                self.logger.exception("Detailed error information:")
            sys.exit(1)
        finally:
            self.logger.debug("Command execution completed")

    def main(self) -> None:
        """Main entry point for the Office VBA CLI."""
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

            parser = self.create_cli_parser()
            pre_args, _ = parser.parse_known_args()
            config = None
            config_load_failed = False
            command = getattr(pre_args, "command", None)
            config_path = getattr(pre_args, "conf", None)
            if command in {"edit", "import", "export"} and config_path:
                try:
                    config = load_config_file(config_path)
                except Exception as e:
                    error(f"Error loading configuration file: {e}")
                    config_load_failed = True
                else:
                    subparser = self._get_subparser(parser, command)
                    target_parser = subparser or parser
                    config_defaults = self._get_config_defaults(target_parser, config)
                    if config_defaults:
                        target_parser.set_defaults(**config_defaults)

            args = parser.parse_args()

            # Apply configuration and resolve placeholders BEFORE setting up logging
            if not config_load_failed:
                if config and getattr(args, "conf", None):
                    args = merge_config_with_args(args, config)
                    args = resolve_all_placeholders(args, args.conf)
                else:
                    args = resolve_all_placeholders(args, None)

            # Set up logging first
            setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))

            # Create target directories and validate inputs early
            self.validate_paths(args)

            # Run 'check' command (Check if VBA project model is accessible)
            if args.command == "check":
                from vba_edit.utils import check_vba_trust_access

                try:
                    if getattr(args, "subcommand", None) == "all":
                        check_vba_trust_access()  # Check all supported Office applications
                    else:
                        check_vba_trust_access(self.office_app)  # Check specific Office app only
                except Exception as e:
                    self.logger.error(f"Failed to check Trust Access to VBA project object model: {str(e)}")
                sys.exit(0)
            else:
                self.handle_office_vba_command(args)

        except Exception as e:
            from vba_edit.console import print_exception

            # Use print_exception to properly render Rich markup in exception messages
            print_exception(e)
            sys.exit(1)


def create_office_main(office_app: str):
    """Create a main function for a specific Office application.

    Args:
        office_app: Office application name (excel, word, access, powerpoint)

    Returns:
        Main function for the Office application
    """

    def main():
        cli = OfficeVBACLI(office_app)
        cli.main()

    return main

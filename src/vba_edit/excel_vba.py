import argparse
import os
import subprocess
import sys
from pathlib import Path

from vba_edit.office_cli import create_office_main

# region UnifiedCLI_PRcompat
# TODO: Remove once test modules no longer depend on these module-level functions.
#       These shims exist solely to avoid breaking existing tests when refactoring the CLI.
from vba_edit.office_cli import OfficeVBACLI

# Re-export symbols that older tests patch at vba_edit.excel_vba.*.
# NOTE: OfficeVBACLI uses the implementations imported in vba_edit.office_cli, so we
# also "sync" these re-exports back into that module at runtime to make patching work.
import vba_edit.office_cli as office_cli
from vba_edit.cli_common import handle_export_with_warnings as handle_export_with_warnings
from vba_edit.office_vba import ExcelVBAHandler as ExcelVBAHandler
from vba_edit.path_utils import get_document_paths as get_document_paths
from vba_edit.utils import setup_logging as setup_logging


# endregion

MODULE_NAME = "excel"


# region UnifiedCLI_PRcompat
# TODO: Remove once test modules no longer depend on these module-level functions.
#       These shims exist solely to avoid breaking existing tests when refactoring the CLI.
def _sync_prcompat_patches_to_office_cli() -> None:
    """Ensure patches applied to vba_edit.excel_vba.* affect vba_edit.office_cli.* too."""
    office_cli.setup_logging = setup_logging
    office_cli.get_document_paths = get_document_paths
    office_cli.handle_export_with_warnings = handle_export_with_warnings
    office_cli.ExcelVBAHandler = ExcelVBAHandler


def validate_paths(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for path validation."""
    _sync_prcompat_patches_to_office_cli()
    OfficeVBACLI(MODULE_NAME).validate_paths(args)


def handle_excel_vba_command(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for command handling."""
    _sync_prcompat_patches_to_office_cli()
    OfficeVBACLI(MODULE_NAME).handle_office_vba_command(args)


def create_cli_parser() -> argparse.ArgumentParser:
    """Backwards-compatible wrapper.

    NOTE: The CLI is now centralized in OfficeVBACLI (office_cli.py). This function
    remains to preserve the historical excel_vba module API and entry-point tests.
    """
    _sync_prcompat_patches_to_office_cli()
    return OfficeVBACLI(MODULE_NAME).create_cli_parser()


# endregion


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


# If xlwings option is used, check dependency before proceeding
def excel_xlwings_handler(cli_instance, args: argparse.Namespace) -> bool:
    """Handle Excel xlwings functionality. Returns True if handled, exits otherwise."""
    if not getattr(args, "xlwings", False):
        return False

    try:
        import xlwings  # type: ignore[import-not-found]

        cli_instance.logger.info(f"Using xlwings {xlwings.__version__}")
        handle_xlwings_command(args)
        return True

    except ImportError:
        sys.exit("xlwings is not installed. Please install it with: pip install xlwings")

    except Exception as e:
        from vba_edit.console import print_exception

        # Use print_exception to properly render Rich markup in exception messages
        print_exception(e)
        sys.exit(1)


# Create the main function for Excel
main = create_office_main(MODULE_NAME)

if __name__ == "__main__":
    main()

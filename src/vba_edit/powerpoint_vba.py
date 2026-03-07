"""PowerPoint VBA CLI module."""

from vba_edit.office_cli import create_office_main

# region UnifiedCLI_PRcompat
# TODO: Remove once test modules no longer depend on these module-level functions.
#       These shims exist solely to avoid breaking existing tests when refactoring the CLI.
import argparse
from vba_edit.office_cli import OfficeVBACLI

# endregion

MODULE_NAME = "powerpoint"


# region UnifiedCLI_PRcompat
# TODO: Remove once test modules no longer depend on these module-level functions.
#       These shims exist solely to avoid breaking existing tests when refactoring the CLI.
def validate_paths(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for path validation."""
    OfficeVBACLI(MODULE_NAME).validate_paths(args)


def handle_powerpoint_vba_command(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for command handling."""
    OfficeVBACLI(MODULE_NAME).handle_office_vba_command(args)


def create_cli_parser() -> argparse.ArgumentParser:
    """Backwards-compatible wrapper.

    NOTE: The CLI is now centralized in OfficeVBACLI (office_cli.py). This function
    remains to preserve the historical powerpoint_vba module API and entry-point tests.
    """
    return OfficeVBACLI(MODULE_NAME).create_cli_parser()


# endregion

# Create the main function for PowerPoint
main = create_office_main(MODULE_NAME)

if __name__ == "__main__":
    main()

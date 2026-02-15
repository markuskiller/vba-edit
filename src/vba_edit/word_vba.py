"""Word VBA CLI module."""

from vba_edit.office_cli import create_office_main

# region PRcompat
# for PR integration only, until test modules are updated
import argparse
from vba_edit.office_cli import OfficeVBACLI

# endregion
# region PRcompat
MODULE_NAME = "word"


def validate_paths(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for path validation."""
    OfficeVBACLI(MODULE_NAME).validate_paths(args)


def handle_word_vba_command(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for command handling."""
    OfficeVBACLI(MODULE_NAME).handle_office_vba_command(args)


def create_cli_parser() -> argparse.ArgumentParser:
    """Backwards-compatible wrapper.

    NOTE: The CLI is now centralized in OfficeVBACLI (office_cli.py). This function
    remains to preserve the historical word_vba module API and entry-point tests.
    """
    return OfficeVBACLI(MODULE_NAME).create_cli_parser()


# endregion

# Create the main function for Word
main = create_office_main("word")

if __name__ == "__main__":
    main()

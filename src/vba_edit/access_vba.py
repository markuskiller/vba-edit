import argparse
import logging
import sys

from vba_edit.exceptions import VBAError
from vba_edit.office_cli import create_office_main

# region PRcompat
# for PR integration only, until test modules are updated
from vba_edit.office_cli import OfficeVBACLI
# endregion
# region PRcompat


MODULE_NAME = "access"


def validate_paths(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for path validation."""
    OfficeVBACLI(MODULE_NAME).validate_paths(args)


def handle_access_vba_command(args: argparse.Namespace) -> None:
    """Backwards-compatible wrapper for command handling."""
    OfficeVBACLI(MODULE_NAME).handle_office_vba_command(args)


def create_cli_parser() -> argparse.ArgumentParser:
    """Backwards-compatible wrapper.

    NOTE: The CLI is now centralized in OfficeVBACLI (office_cli.py). This function
    remains to preserve the historical excel_vba module API and entry-point tests.
    """
    return OfficeVBACLI(MODULE_NAME).create_cli_parser()


# endregion

logger = logging.getLogger(__name__)


def check_multiple_databases(file_path: str = None) -> None:
    """Check for multiple open databases and handle appropriately.

    Args:
        file_path: Optional path to specific database file

    Raises:
        VBAError: If multiple databases are open and no specific file is provided
    """
    try:
        import win32com.client

        app = win32com.client.GetObject("Access.Application")
        try:
            # Get current database
            current_db = app.CurrentDb()
            if not current_db:
                # No database open
                return

            # Check for other open databases using a more reliable method
            try:
                current_name = current_db.Name
                open_dbs = []

                # Check each database connection directly
                for i in range(app.DBEngine.Workspaces(0).Databases.Count):
                    try:
                        db = app.DBEngine.Workspaces(0).Databases(i)
                        if db and db.Name != current_name:
                            open_dbs.append(db.Name)
                    except Exception:
                        continue

                if open_dbs and not file_path:
                    raise VBAError(
                        "Multiple Access databases are open. Please specify the target "
                        "database using the --file option."
                    )
            except AttributeError:
                # DBEngine or Workspaces not accessible, consider it a single database
                logger.debug("Could not enumerate databases, assuming single database")
                return

        except Exception as e:
            logger.debug(f"Error checking current database: {e}")
            return

    except Exception:
        # Only log at debug level since this is a non-critical check
        logger.debug("Could not check for multiple databases - Access may not be running")
        return


def access_pre_command_hook(args: argparse.Namespace) -> None:
    """Access-specific pre-command processing."""
    try:
        check_multiple_databases(args.file)
    except VBAError as e:
        logger.error(str(e))
        sys.exit(1)


# Create the main function for Access
main = create_office_main("access")

if __name__ == "__main__":
    main()

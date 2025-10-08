"""Tests for GitHub Issues #20 and #21 - CLI Interface Issues.

Issue #20: word-vba.exe check command fails with AttributeError 'Namespace' object has no attribute 'file'
Issue #21: CLI options should be command-specific, not global (e.g., --rubberduck-folders)

Test Strategy:
- Issue #20: Verify that 'check' command works without file/vba_directory arguments
- Issue #21: Verify that folder organization options only exist on appropriate commands
"""

import argparse
import subprocess
import sys
from io import StringIO
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import pytest

from vba_edit.cli_common import (
    add_common_arguments,
    add_folder_organization_arguments,
)


class TestIssue20_CheckCommandAttributeError:
    """Tests for Issue #20 - check command AttributeError fix.

    The 'check' command doesn't use file/vba_directory arguments but validate_paths()
    was trying to access them, causing AttributeError.
    """

    def test_word_vba_check_command_parser(self):
        """Test that word-vba check command parser works without file argument."""
        from vba_edit.word_vba import create_cli_parser

        parser = create_cli_parser()

        # Should parse successfully without error
        args = parser.parse_args(["check"])
        assert args.command == "check"

        # Should not have file or vba_directory attributes (not defined for check command)
        assert not hasattr(args, "file") or args.file is None

    def test_word_vba_check_all_command_parser(self):
        """Test that word-vba check all command parser works."""
        from vba_edit.word_vba import create_cli_parser

        parser = create_cli_parser()

        # Should parse successfully
        args = parser.parse_args(["check", "all"])
        assert args.command == "check"
        assert args.subcommand == "all"

    def test_word_vba_validate_paths_skips_check_command(self):
        """Test that validate_paths() skips validation for check command."""
        from vba_edit.word_vba import validate_paths

        # Create a mock args namespace for check command
        args = argparse.Namespace(command="check")

        # Should not raise AttributeError even though file/vba_directory don't exist
        try:
            validate_paths(args)
        except AttributeError as e:
            pytest.fail(f"validate_paths raised AttributeError for check command: {e}")

    def test_excel_vba_check_command_parser(self):
        """Test that excel-vba check command parser works."""
        from vba_edit.excel_vba import create_cli_parser

        parser = create_cli_parser()
        args = parser.parse_args(["check"])
        assert args.command == "check"

    def test_excel_vba_validate_paths_skips_check_command(self):
        """Test that excel validate_paths() skips validation for check command."""
        from vba_edit.excel_vba import validate_paths

        args = argparse.Namespace(command="check")

        try:
            validate_paths(args)
        except AttributeError as e:
            pytest.fail(f"validate_paths raised AttributeError for check command: {e}")

    def test_access_vba_check_command_parser(self):
        """Test that access-vba check command parser works."""
        from vba_edit.access_vba import create_cli_parser

        parser = create_cli_parser()
        args = parser.parse_args(["check"])
        assert args.command == "check"

    def test_access_vba_validate_paths_skips_check_command(self):
        """Test that access validate_paths() skips validation for check command."""
        from vba_edit.access_vba import validate_paths

        args = argparse.Namespace(command="check")

        try:
            validate_paths(args)
        except AttributeError as e:
            pytest.fail(f"validate_paths raised AttributeError for check command: {e}")

    def test_powerpoint_vba_check_command_parser(self):
        """Test that powerpoint-vba check command parser works."""
        from vba_edit.powerpoint_vba import create_cli_parser

        parser = create_cli_parser()
        args = parser.parse_args(["check"])
        assert args.command == "check"

    def test_powerpoint_vba_validate_paths_skips_check_command(self):
        """Test that powerpoint validate_paths() skips validation for check command."""
        from vba_edit.powerpoint_vba import validate_paths

        args = argparse.Namespace(command="check")

        try:
            validate_paths(args)
        except AttributeError as e:
            pytest.fail(f"validate_paths raised AttributeError for check command: {e}")

    def test_validate_paths_works_for_edit_command_with_file(self, tmp_path):
        """Test that validate_paths() still works correctly for edit command."""
        from vba_edit.word_vba import validate_paths

        # Create a temporary file
        test_file = tmp_path / "test.docm"
        test_file.touch()

        args = argparse.Namespace(command="edit", file=str(test_file), vba_directory=None)

        # Should not raise any errors
        validate_paths(args)

    def test_validate_paths_fails_for_nonexistent_file(self):
        """Test that validate_paths() still catches nonexistent files for edit/import/export."""
        from vba_edit.word_vba import validate_paths

        args = argparse.Namespace(command="export", file="nonexistent_file.docm", vba_directory=None)

        # Should raise FileNotFoundError
        with pytest.raises(FileNotFoundError, match="Document not found"):
            validate_paths(args)


class TestIssue21_CommandSpecificOptions:
    """Tests for Issue #21 - CLI options should be command-specific.

    Options like --rubberduck-folders and --open-folder should only be available
    for commands that export VBA code (edit, import, export), not globally.
    """

    def test_folder_organization_arguments_function_exists(self):
        """Test that add_folder_organization_arguments function exists."""
        # The function should be imported and available
        assert callable(add_folder_organization_arguments)

    def test_folder_organization_arguments_adds_rubberduck_folders(self):
        """Test that add_folder_organization_arguments adds --rubberduck-folders option."""
        parser = argparse.ArgumentParser()
        add_folder_organization_arguments(parser)

        # Should parse --rubberduck-folders flag
        args = parser.parse_args(["--rubberduck-folders"])
        assert hasattr(args, "rubberduck_folders")
        assert args.rubberduck_folders is True

        # Default should be False
        args = parser.parse_args([])
        assert args.rubberduck_folders is False

    def test_folder_organization_arguments_adds_open_folder(self):
        """Test that add_folder_organization_arguments adds --open-folder option."""
        parser = argparse.ArgumentParser()
        add_folder_organization_arguments(parser)

        # Should parse --open-folder flag
        args = parser.parse_args(["--open-folder"])
        assert hasattr(args, "open_folder")
        assert args.open_folder is True

        # Default should be False
        args = parser.parse_args([])
        assert args.open_folder is False

    def test_word_vba_edit_has_folder_options(self):
        """Test that word-vba edit command has folder organization options."""
        from vba_edit.word_vba import create_cli_parser

        parser = create_cli_parser()

        # Should accept --rubberduck-folders and --open-folder
        args = parser.parse_args(["edit", "--rubberduck-folders", "--open-folder"])
        assert args.command == "edit"
        assert args.rubberduck_folders is True
        assert args.open_folder is True

    def test_word_vba_import_has_folder_options(self):
        """Test that word-vba import command has folder organization options."""
        from vba_edit.word_vba import create_cli_parser

        parser = create_cli_parser()

        args = parser.parse_args(["import", "--rubberduck-folders"])
        assert args.command == "import"
        assert args.rubberduck_folders is True

    def test_word_vba_export_has_folder_options(self):
        """Test that word-vba export command has folder organization options."""
        from vba_edit.word_vba import create_cli_parser

        parser = create_cli_parser()

        args = parser.parse_args(["export", "--open-folder"])
        assert args.command == "export"
        assert args.open_folder is True

    def test_word_vba_check_rejects_folder_options(self):
        """Test that word-vba check command rejects folder organization options."""
        from vba_edit.word_vba import create_cli_parser

        parser = create_cli_parser()

        # Should fail when trying to use --rubberduck-folders with check command
        with pytest.raises(SystemExit):
            parser.parse_args(["check", "--rubberduck-folders"])

        # Should fail with --open-folder too
        with pytest.raises(SystemExit):
            parser.parse_args(["check", "--open-folder"])

    def test_excel_vba_edit_has_folder_options(self):
        """Test that excel-vba edit command has folder organization options."""
        from vba_edit.excel_vba import create_cli_parser

        parser = create_cli_parser()

        args = parser.parse_args(["edit", "--rubberduck-folders", "--open-folder"])
        assert args.command == "edit"
        assert args.rubberduck_folders is True
        assert args.open_folder is True

    def test_excel_vba_check_rejects_folder_options(self):
        """Test that excel-vba check command rejects folder organization options."""
        from vba_edit.excel_vba import create_cli_parser

        parser = create_cli_parser()

        with pytest.raises(SystemExit):
            parser.parse_args(["check", "--rubberduck-folders"])

    def test_access_vba_export_has_folder_options(self):
        """Test that access-vba export command has folder organization options."""
        from vba_edit.access_vba import create_cli_parser

        parser = create_cli_parser()

        args = parser.parse_args(["export", "--rubberduck-folders", "--open-folder"])
        assert args.command == "export"
        assert args.rubberduck_folders is True
        assert args.open_folder is True

    def test_access_vba_check_rejects_folder_options(self):
        """Test that access-vba check command rejects folder organization options."""
        from vba_edit.access_vba import create_cli_parser

        parser = create_cli_parser()

        with pytest.raises(SystemExit):
            parser.parse_args(["check", "--rubberduck-folders"])

    def test_powerpoint_vba_import_has_folder_options(self):
        """Test that powerpoint-vba import command has folder organization options."""
        from vba_edit.powerpoint_vba import create_cli_parser

        parser = create_cli_parser()

        args = parser.parse_args(["import", "--rubberduck-folders"])
        assert args.command == "import"
        assert args.rubberduck_folders is True

    def test_powerpoint_vba_check_rejects_folder_options(self):
        """Test that powerpoint-vba check command rejects folder organization options."""
        from vba_edit.powerpoint_vba import create_cli_parser

        parser = create_cli_parser()

        with pytest.raises(SystemExit):
            parser.parse_args(["check", "--open-folder"])

    def test_common_arguments_no_longer_include_folder_options(self):
        """Test that add_common_arguments doesn't include folder organization options."""
        parser = argparse.ArgumentParser()
        add_common_arguments(parser)

        # Should fail to parse --rubberduck-folders (not in common arguments)
        with pytest.raises(SystemExit):
            parser.parse_args(["--rubberduck-folders"])

        # Should fail to parse --open-folder (not in common arguments)
        with pytest.raises(SystemExit):
            parser.parse_args(["--open-folder"])

    def test_common_arguments_still_include_file_and_verbose(self):
        """Test that add_common_arguments still includes appropriate common options."""
        parser = argparse.ArgumentParser()
        add_common_arguments(parser)

        # These should still be in common arguments
        args = parser.parse_args(["--file", "test.docm", "--verbose"])
        assert args.file == "test.docm"
        assert args.verbose is True


class TestIssue20And21_Integration:
    """Integration tests combining both issues."""

    def test_word_vba_check_command_full_flow(self):
        """Test complete flow of word-vba check command without errors."""
        from vba_edit.word_vba import create_cli_parser, validate_paths

        parser = create_cli_parser()
        args = parser.parse_args(["check"])

        # Should parse successfully
        assert args.command == "check"

        # validate_paths should not raise AttributeError
        try:
            validate_paths(args)
        except AttributeError:
            pytest.fail("validate_paths raised AttributeError for check command")

    def test_all_entry_points_check_command_works(self):
        """Test that check command works for all entry points."""
        entry_points = [
            ("vba_edit.word_vba", "word-vba"),
            ("vba_edit.excel_vba", "excel-vba"),
            ("vba_edit.access_vba", "access-vba"),
            ("vba_edit.powerpoint_vba", "powerpoint-vba"),
        ]

        for module_name, entry_point_name in entry_points:
            module = __import__(module_name, fromlist=["create_cli_parser", "validate_paths"])
            parser = module.create_cli_parser()

            # Parse check command
            args = parser.parse_args(["check"])
            assert args.command == "check", f"{entry_point_name} check command parsing failed"

            # Validate paths should not raise AttributeError
            try:
                module.validate_paths(args)
            except AttributeError as e:
                pytest.fail(f"{entry_point_name} validate_paths raised AttributeError: {e}")

    def test_all_entry_points_folder_options_on_edit_only(self):
        """Test that folder options work on edit command for all entry points."""
        entry_points = [
            "vba_edit.word_vba",
            "vba_edit.excel_vba",
            "vba_edit.access_vba",
            "vba_edit.powerpoint_vba",
        ]

        for module_name in entry_points:
            module = __import__(module_name, fromlist=["create_cli_parser"])
            parser = module.create_cli_parser()

            # Should work with edit command
            args = parser.parse_args(["edit", "--rubberduck-folders", "--open-folder"])
            assert args.rubberduck_folders is True
            assert args.open_folder is True

            # Should fail with check command
            with pytest.raises(SystemExit):
                parser.parse_args(["check", "--rubberduck-folders"])


class TestRegressionPrevention:
    """Tests to ensure fixes don't break existing functionality."""

    def test_edit_command_still_requires_handler_args(self):
        """Test that edit command still properly uses folder organization args."""
        from vba_edit.word_vba import create_cli_parser

        parser = create_cli_parser()

        # Parse edit command with options
        args = parser.parse_args(["edit", "--rubberduck-folders", "--open-folder", "--verbose"])

        # Verify all arguments are present
        assert hasattr(args, "rubberduck_folders")
        assert hasattr(args, "open_folder")
        assert hasattr(args, "verbose")
        assert args.rubberduck_folders is True
        assert args.open_folder is True
        assert args.verbose is True

    def test_export_command_with_all_options(self):
        """Test that export command can use all its options without conflicts."""
        from vba_edit.excel_vba import create_cli_parser

        parser = create_cli_parser()

        args = parser.parse_args(
            [
                "export",
                "--file",
                "test.xlsm",
                "--vba-directory",
                "src",
                "--verbose",
                "--save-metadata",
                "--force-overwrite",
                "--rubberduck-folders",
                "--open-folder",
                "--in-file-headers",
            ]
        )

        assert args.command == "export"
        assert args.file == "test.xlsm"
        assert args.vba_directory == "src"
        assert args.verbose is True
        assert args.save_metadata is True
        assert args.force_overwrite is True
        assert args.rubberduck_folders is True
        assert args.open_folder is True
        assert args.in_file_headers is True

    def test_handler_creation_with_getattr_defaults(self):
        """Test that handler can be created even when folder args are missing."""
        # Simulate args namespace without folder organization options (like check command)
        args = argparse.Namespace(command="check", verbose=False)

        # These getattr calls should work with defaults
        rubberduck = getattr(args, "rubberduck_folders", False)
        open_folder = getattr(args, "open_folder", False)

        assert rubberduck is False
        assert open_folder is False


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])

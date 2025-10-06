"""Tests for Excel VBA CLI functionality."""

import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

from vba_edit.excel_vba import create_cli_parser, handle_excel_vba_command


def test_rubberduck_folders_option():
    """Test that the CLI parser includes Rubberduck folders option."""
    parser = create_cli_parser()

    # Test edit command with rubberduck folders
    args = parser.parse_args(["edit", "--rubberduck-folders"])
    assert args.rubberduck_folders is True

    # Test export command with rubberduck folders
    args = parser.parse_args(["export", "--rubberduck-folders"])
    assert args.rubberduck_folders is True

    # Test import command with rubberduck folders
    args = parser.parse_args(["import", "--rubberduck-folders"])
    assert args.rubberduck_folders is True

    # Test that check command doesn't have rubberduck option
    args = parser.parse_args(["check"])
    assert not hasattr(args, "--rubberduck_folders")


@patch("vba_edit.excel_vba.ExcelVBAHandler")
@patch("vba_edit.excel_vba.get_document_paths")
@patch("vba_edit.excel_vba.setup_logging")
def test_rubberduck_folders_passed_to_handler(mock_logging, mock_get_paths, mock_handler):
    """Test that rubberduck_folders option is passed to the handler."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        doc_path = tmp_path / "test.xlsm"
        doc_path.touch()

        # Mock the path resolution
        mock_get_paths.return_value = (doc_path, tmp_path)

        # Create mock handler instance
        mock_handler_instance = Mock()
        mock_handler.return_value = mock_handler_instance

        # Create args with rubberduck_folders enabled
        parser = create_cli_parser()
        args = parser.parse_args(["export", "--rubberduck-folders", "--file", str(doc_path)])

        # Handle the command
        handle_excel_vba_command(args)

        # Verify handler was called with use_rubberduck_folders=True
        mock_handler.assert_called_once()
        call_kwargs = mock_handler.call_args[1]
        assert call_kwargs["use_rubberduck_folders"] is True


def test_watchfiles_integration():
    pass


#     """Test that watchfiles is properly integrated."""
#     try:
#         from watchfiles import watch, Change
#         assert True, "watchfiles imported successfully"
#     except ImportError:
#         pytest.fail("watchfiles not available - please update dependencies")

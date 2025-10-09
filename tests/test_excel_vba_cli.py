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


def test_save_metadata_option():
    """Test that the CLI parser includes save-metadata option for edit and export commands."""
    parser = create_cli_parser()

    # Test edit command with save-metadata
    args = parser.parse_args(["edit", "--save-metadata"])
    assert args.save_metadata is True

    # Test edit command with -m shorthand
    args = parser.parse_args(["edit", "-m"])
    assert args.save_metadata is True

    # Test export command with save-metadata
    args = parser.parse_args(["export", "--save-metadata"])
    assert args.save_metadata is True

    # Test export command with -m shorthand
    args = parser.parse_args(["export", "-m"])
    assert args.save_metadata is True

    # Test that import command also has save-metadata (if it does)
    # Note: This will fail if import doesn't support it yet
    # args = parser.parse_args(["import", "--save-metadata"])
    # assert args.save_metadata is True


@patch("vba_edit.excel_vba.handle_export_with_warnings")
@patch("vba_edit.excel_vba.ExcelVBAHandler")
@patch("vba_edit.excel_vba.get_document_paths")
@patch("vba_edit.excel_vba.setup_logging")
def test_save_metadata_passed_to_handler_edit(mock_logging, mock_get_paths, mock_handler, mock_export):
    """Test that save_metadata option is passed to handle_export_with_warnings for edit command."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        doc_path = tmp_path / "test.xlsm"
        doc_path.touch()

        # Mock the path resolution
        mock_get_paths.return_value = (doc_path, tmp_path)

        # Create mock handler instance
        mock_handler_instance = Mock()
        mock_handler_instance.watch_changes = Mock(side_effect=KeyboardInterrupt)  # Exit gracefully
        mock_handler.return_value = mock_handler_instance

        # Create args with save_metadata enabled
        parser = create_cli_parser()
        args = parser.parse_args(["edit", "--save-metadata", "--file", str(doc_path)])

        # Handle the command (will be interrupted by KeyboardInterrupt)
        try:
            handle_excel_vba_command(args)
        except (KeyboardInterrupt, SystemExit):
            pass

        # Verify handle_export_with_warnings was called with save_metadata=True
        mock_export.assert_called()
        call_kwargs = mock_export.call_args[1]
        assert call_kwargs["save_metadata"] is True
        assert call_kwargs["overwrite"] is False
        assert call_kwargs["interactive"] is True


@patch("vba_edit.excel_vba.handle_export_with_warnings")
@patch("vba_edit.excel_vba.ExcelVBAHandler")
@patch("vba_edit.excel_vba.get_document_paths")
@patch("vba_edit.excel_vba.setup_logging")
def test_save_metadata_passed_to_handler_export(mock_logging, mock_get_paths, mock_handler, mock_export):
    """Test that save_metadata option is passed to handle_export_with_warnings for export command."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        doc_path = tmp_path / "test.xlsm"
        doc_path.touch()

        # Mock the path resolution
        mock_get_paths.return_value = (doc_path, tmp_path)

        # Create mock handler instance
        mock_handler_instance = Mock()
        mock_handler.return_value = mock_handler_instance

        # Create args with save_metadata enabled
        parser = create_cli_parser()
        args = parser.parse_args(["export", "--save-metadata", "--file", str(doc_path)])

        # Handle the command
        try:
            handle_excel_vba_command(args)
        except SystemExit:
            pass

        # Verify handle_export_with_warnings was called with save_metadata=True
        mock_export.assert_called_once()
        call_kwargs = mock_export.call_args[1]
        assert call_kwargs["save_metadata"] is True


@patch("vba_edit.excel_vba.handle_export_with_warnings")
@patch("vba_edit.excel_vba.ExcelVBAHandler")
@patch("vba_edit.excel_vba.get_document_paths")
@patch("vba_edit.excel_vba.setup_logging")
def test_save_metadata_default_false(mock_logging, mock_get_paths, mock_handler, mock_export):
    """Test that save_metadata defaults to False when not specified."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        doc_path = tmp_path / "test.xlsm"
        doc_path.touch()

        # Mock the path resolution
        mock_get_paths.return_value = (doc_path, tmp_path)

        # Create mock handler instance
        mock_handler_instance = Mock()
        mock_handler.return_value = mock_handler_instance

        # Create args WITHOUT save_metadata
        parser = create_cli_parser()
        args = parser.parse_args(["export", "--file", str(doc_path)])

        # Handle the command
        try:
            handle_excel_vba_command(args)
        except SystemExit:
            pass

        # Verify handle_export_with_warnings was called with save_metadata=False
        mock_export.assert_called_once()
        call_kwargs = mock_export.call_args[1]
        assert call_kwargs["save_metadata"] is False


#     """Test that watchfiles is properly integrated."""
#     try:
#         from watchfiles import watch, Change
#         assert True, "watchfiles imported successfully"
#     except ImportError:
#         pytest.fail("watchfiles not available - please update dependencies")

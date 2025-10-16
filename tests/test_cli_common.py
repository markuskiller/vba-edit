"""Tests for common CLI utilities."""

import argparse
import sys
from io import StringIO
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import pytest

from vba_edit.cli_common import (
    add_export_arguments,
    handle_export_with_warnings,
)
from vba_edit.exceptions import VBAExportWarning


class TestAddExportArguments:
    """Tests for add_export_arguments function."""

    def test_adds_force_overwrite_flag(self):
        """Test that --force-overwrite flag is added to parser."""
        parser = argparse.ArgumentParser()
        add_export_arguments(parser)

        # Parse with the flag
        args = parser.parse_args(["--force-overwrite"])
        assert hasattr(args, "force_overwrite")
        assert args.force_overwrite is True

    def test_force_overwrite_default_false(self):
        """Test that --force-overwrite defaults to False."""
        parser = argparse.ArgumentParser()
        add_export_arguments(parser)

        # Parse without the flag
        args = parser.parse_args([])
        assert hasattr(args, "force_overwrite")
        assert args.force_overwrite is False

    def test_force_overwrite_no_short_option(self):
        """Test that -f is not used (reserved for --file)."""
        parser = argparse.ArgumentParser()
        add_export_arguments(parser)

        # -f should not work for force_overwrite
        with pytest.raises(SystemExit):
            parser.parse_args(["-f"])


class TestHandleExportWithWarnings:
    """Tests for handle_export_with_warnings helper function."""

    def test_successful_export_no_warnings(self):
        """Test successful export without any warnings."""
        mock_handler = Mock()
        mock_handler.export_vba = Mock()

        handle_export_with_warnings(
            mock_handler, save_metadata=True, overwrite=True, interactive=True, force_overwrite=False
        )

        # Should call export_vba once
        mock_handler.export_vba.assert_called_once_with(
            save_metadata=True, overwrite=True, interactive=True, keep_open=False
        )

    def test_existing_files_warning_user_confirms(self):
        """Test handling of existing_files warning when user confirms."""
        mock_handler = Mock()

        # First call raises warning, second call succeeds
        warning = VBAExportWarning("existing_files", {"file_count": 5, "files": ["Module1.bas", "Class1.cls"]})
        mock_handler.export_vba = Mock(side_effect=[warning, None])

        with patch("vba_edit.cli_common.confirm_action", return_value=True):
            handle_export_with_warnings(
                mock_handler, save_metadata=False, overwrite=True, interactive=True, force_overwrite=False
            )

        # Should call export_vba twice (first attempt + retry)
        assert mock_handler.export_vba.call_count == 2

        # Second call should have interactive=False
        second_call_kwargs = mock_handler.export_vba.call_args_list[1][1]
        assert second_call_kwargs["interactive"] is False

    def test_existing_files_warning_user_cancels(self):
        """Test handling of existing_files warning when user cancels."""
        mock_handler = Mock()

        warning = VBAExportWarning("existing_files", {"file_count": 5, "files": ["Module1.bas"]})
        mock_handler.export_vba = Mock(side_effect=warning)

        with patch("vba_edit.cli_common.confirm_action", return_value=False):
            with pytest.raises(SystemExit) as exc_info:
                handle_export_with_warnings(
                    mock_handler, save_metadata=False, overwrite=True, interactive=True, force_overwrite=False
                )

            # Should exit with code 0 (user cancellation, not error)
            assert exc_info.value.code == 0

        # Should only call export_vba once (no retry)
        mock_handler.export_vba.assert_called_once()

    def test_header_mode_changed_warning_user_confirms(self):
        """Test handling of header_mode_changed warning when user confirms."""
        mock_handler = Mock()

        warning = VBAExportWarning(
            "header_mode_changed", {"old_mode": "separate header files", "new_mode": "inline headers"}
        )
        mock_handler.export_vba = Mock(side_effect=[warning, None])

        with patch("vba_edit.cli_common.confirm_action", return_value=True):
            handle_export_with_warnings(
                mock_handler, save_metadata=True, overwrite=True, interactive=True, force_overwrite=False
            )

        # Should call export_vba twice
        assert mock_handler.export_vba.call_count == 2

        # Second call should have overwrite=True and interactive=False
        second_call_kwargs = mock_handler.export_vba.call_args_list[1][1]
        assert second_call_kwargs["overwrite"] is True
        assert second_call_kwargs["interactive"] is False

    def test_header_mode_changed_warning_user_cancels(self):
        """Test handling of header_mode_changed warning when user cancels."""
        mock_handler = Mock()

        warning = VBAExportWarning("header_mode_changed", {"old_mode": "inline", "new_mode": "separate"})
        mock_handler.export_vba = Mock(side_effect=warning)

        with patch("vba_edit.cli_common.confirm_action", return_value=False):
            with pytest.raises(SystemExit) as exc_info:
                handle_export_with_warnings(
                    mock_handler, save_metadata=False, overwrite=True, interactive=True, force_overwrite=False
                )

            assert exc_info.value.code == 0

        mock_handler.export_vba.assert_called_once()

    def test_force_overwrite_skips_prompts(self):
        """Test that force_overwrite=True skips all confirmation prompts."""
        mock_handler = Mock()
        mock_handler.export_vba = Mock()

        with patch("vba_edit.cli_common.confirm_action") as mock_confirm:
            handle_export_with_warnings(
                mock_handler, save_metadata=True, overwrite=True, interactive=True, force_overwrite=True
            )

        # confirm_action should never be called
        mock_confirm.assert_not_called()

        # export_vba should be called with interactive=False and keep_open=False
        mock_handler.export_vba.assert_called_once_with(
            save_metadata=True, overwrite=True, interactive=False, keep_open=False
        )

    def test_force_overwrite_logs_usage(self):
        """Test that using force_overwrite is logged."""
        mock_handler = Mock()
        mock_handler.export_vba = Mock()

        with patch("vba_edit.cli_common.logger") as mock_logger:
            handle_export_with_warnings(
                mock_handler, save_metadata=False, overwrite=True, interactive=True, force_overwrite=True
            )

            # Should log that force_overwrite is being used
            mock_logger.info.assert_called_once()
            log_message = mock_logger.info.call_args[0][0]
            assert "force-overwrite" in log_message.lower()

    def test_force_overwrite_with_warning_no_prompt(self):
        """Test force_overwrite bypasses warnings without prompting."""
        mock_handler = Mock()
        mock_handler.export_vba = Mock()  # No warning raised when interactive=False

        with patch("vba_edit.cli_common.confirm_action") as mock_confirm:
            handle_export_with_warnings(
                mock_handler, save_metadata=True, overwrite=True, interactive=True, force_overwrite=True
            )

        # No confirmation should be requested
        mock_confirm.assert_not_called()

        # Should call with interactive=False (due to force_overwrite)
        call_kwargs = mock_handler.export_vba.call_args[1]
        assert call_kwargs["interactive"] is False

    def test_unknown_warning_type_ignored(self):
        """Test that unknown warning types are ignored (function completes successfully)."""
        mock_handler = Mock()

        # Unknown warning type raised once, then succeeds
        warning = VBAExportWarning("unknown_type", {"data": "test"})
        mock_handler.export_vba = Mock(side_effect=[warning, None])

        # Should not raise an exception - unknown warnings are silently ignored
        # and the function completes (though it doesn't retry like known types)
        try:
            handle_export_with_warnings(
                mock_handler, save_metadata=False, overwrite=True, interactive=True, force_overwrite=False
            )
            # If we get here without exception, the warning was caught and ignored
            # This is acceptable behavior - only call once since no retry happens
        except VBAExportWarning:
            # This would be unexpected - warning should be caught
            pytest.fail("Unknown warning type should be caught by except handler")

        # Should only call export_vba once (no retry for unknown types)
        assert mock_handler.export_vba.call_count == 1

    def test_output_messages_for_existing_files(self):
        """Test that appropriate messages are printed for existing_files warning."""
        from io import StringIO
        from vba_edit.console import console

        mock_handler = Mock()

        warning = VBAExportWarning(
            "existing_files", {"file_count": 3, "files": ["Module1.bas", "Class1.cls", "Form1.frm"]}
        )
        mock_handler.export_vba = Mock(side_effect=[warning, None])

        # Capture console output
        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        try:
            with patch("vba_edit.cli_common.confirm_action", return_value=True):
                handle_export_with_warnings(
                    mock_handler, save_metadata=False, overwrite=True, interactive=True, force_overwrite=False
                )
        finally:
            console.file = original_file

        output = buffer.getvalue()
        assert "⚠" in output  # warning symbol
        assert "3" in output  # file count
        assert "existing VBA file" in output

    def test_output_messages_for_header_mode_change(self):
        """Test that appropriate messages are printed for header_mode_changed warning."""
        from io import StringIO
        from vba_edit.console import console

        mock_handler = Mock()

        warning = VBAExportWarning(
            "header_mode_changed", {"old_mode": "separate header files", "new_mode": "inline headers"}
        )
        mock_handler.export_vba = Mock(side_effect=[warning, None])

        # Capture console output
        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        try:
            with patch("vba_edit.cli_common.confirm_action", return_value=True):
                handle_export_with_warnings(
                    mock_handler, save_metadata=True, overwrite=True, interactive=True, force_overwrite=False
                )
        finally:
            console.file = original_file

        output = buffer.getvalue()
        assert "⚠" in output  # warning symbol
        assert "Header storage mode" in output
        assert "separate header files" in output
        assert "inline headers" in output

    def test_cancellation_message_printed(self):
        """Test that cancellation message is printed when user cancels."""
        from io import StringIO
        from vba_edit.console import console

        mock_handler = Mock()

        warning = VBAExportWarning("existing_files", {"file_count": 1})
        mock_handler.export_vba = Mock(side_effect=warning)

        # Capture console output
        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        try:
            with patch("vba_edit.cli_common.confirm_action", return_value=False):
                with pytest.raises(SystemExit):
                    handle_export_with_warnings(
                        mock_handler, save_metadata=False, overwrite=True, interactive=True, force_overwrite=False
                    )
        finally:
            console.file = original_file

        output = buffer.getvalue().lower()
        assert "cancelled" in output

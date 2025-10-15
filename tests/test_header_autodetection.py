"""Tests for automatic header detection during import.

This module tests the automatic header format detection that eliminates
the need for --save-headers and --in-file-headers flags on import commands.
"""

import pytest
from pathlib import Path
from vba_edit.office_vba import VBAComponentHandler


class TestHeaderAutoDetection:
    """Test automatic detection of header formats."""

    @pytest.fixture
    def handler(self):
        """Create a VBAComponentHandler instance."""
        return VBAComponentHandler()

    def test_detects_inline_headers_with_version(self, tmp_path, handler):
        """Test detection of inline headers starting with VERSION."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Module1"

Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is True

    def test_detects_inline_headers_with_begin(self, tmp_path, handler):
        """Test detection of inline headers starting with BEGIN."""
        test_file = tmp_path / "UserForm1.frm"
        test_file.write_text(
            """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1
   Caption         =   "Test Form"
   ClientHeight    =   3000
End
Attribute VB_Name = "UserForm1"

Private Sub UserForm_Initialize()
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is True

    def test_detects_inline_headers_with_attributes(self, tmp_path, handler):
        """Test detection of inline headers with Attribute VB_."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """Attribute VB_Name = "Module1"
Attribute VB_GlobalNameSpace = False

Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is True

    def test_no_inline_headers_code_only(self, tmp_path, handler):
        """Test detection correctly identifies code-only files."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """Sub Test()
    MsgBox "Hello World"
End Sub

Function Calculate(x As Integer) As Integer
    Calculate = x * 2
End Function
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is False

    def test_no_inline_headers_with_comments(self, tmp_path, handler):
        """Test detection works with files that have comments at top."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """' This is a comment
' Author: Test User
' Date: 2025-10-14

Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is False

    def test_ignores_version_in_comments(self, tmp_path, handler):
        """Test that VERSION mentioned in comments doesn't trigger detection."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """' VERSION 1.0 released on 2025-01-01
' This module handles data processing

Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is False

    def test_ignores_begin_in_comments(self, tmp_path, handler):
        """Test that BEGIN mentioned in comments doesn't trigger detection."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """' BEGIN with basic setup
' REM BEGIN the initialization
' Initialize system

Sub Test()
    ' BEGIN processing
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is False

    def test_ignores_attribute_in_comments(self, tmp_path, handler):
        """Test that Attribute mentioned in comments doesn't trigger detection."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """' Set the Attribute VB_Name property
' Attribute configuration follows

Sub Test()
    ' Attribute VB_Name should be set
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is False

    def test_detects_version_after_blank_lines(self, tmp_path, handler):
        """Test that VERSION is detected even if there are blank lines at top."""
        test_file = tmp_path / "Module1.cls"
        test_file.write_text(
            """

VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "Module1"

Sub Test()
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is True

    def test_detects_attribute_after_blank_lines(self, tmp_path, handler):
        """Test that Attribute VB_ is detected after blank lines."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """


Attribute VB_Name = "Module1"

Sub Test()
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is True

    def test_handles_empty_file(self, tmp_path, handler):
        """Test detection handles empty files gracefully."""
        test_file = tmp_path / "Empty.bas"
        test_file.write_text("", encoding="utf-8")

        assert handler.has_inline_headers(test_file) is False

    def test_handles_file_with_only_whitespace(self, tmp_path, handler):
        """Test detection handles files with only whitespace."""
        test_file = tmp_path / "Whitespace.bas"
        test_file.write_text("\n\n  \n\n", encoding="utf-8")

        assert handler.has_inline_headers(test_file) is False

    def test_case_insensitive_version_detection(self, tmp_path, handler):
        """Test that VERSION detection is case-insensitive."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """version 1.0 CLASS
BEGIN
END
Attribute VB_Name = "Module1"

Sub Test()
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is True

    def test_case_insensitive_begin_detection(self, tmp_path, handler):
        """Test that BEGIN detection is case-insensitive."""
        test_file = tmp_path / "UserForm1.frm"
        test_file.write_text(
            """VERSION 5.00
begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1
end
Attribute VB_Name = "UserForm1"

Private Sub UserForm_Initialize()
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is True

    def test_attribute_must_be_vb_prefix(self, tmp_path, handler):
        """Test that only 'Attribute VB_' lines are detected, not other attributes."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """' Comment with word Attribute in it

Sub Test()
    ' Set attribute value
    Dim attribute As String
End Sub
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is False

    def test_header_must_be_at_start(self, tmp_path, handler):
        """Test that headers are only detected at file start, not in middle."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """' Comment at top

Sub Test()
    MsgBox "Hello"
End Sub

' VERSION 1.0 CLASS - this is in a comment in the middle
' Attribute VB_Name - also in middle
""",
            encoding="utf-8",
        )

        assert handler.has_inline_headers(test_file) is False

    def test_handles_nonexistent_file(self, tmp_path, handler):
        """Test detection handles nonexistent files gracefully."""
        test_file = tmp_path / "DoesNotExist.bas"

        assert handler.has_inline_headers(test_file) is False

    def test_handles_binary_file(self, tmp_path, handler):
        """Test detection handles binary files gracefully."""
        test_file = tmp_path / "Binary.frx"
        test_file.write_bytes(b"\x00\x01\x02\x03\xff\xfe\xfd")

        # Should return False for binary files without crashing
        result = handler.has_inline_headers(test_file)
        assert result is False


class TestHeaderAutoDetectionIntegration:
    """Integration tests for automatic header detection with import logic."""

    @pytest.fixture
    def handler(self):
        """Create a VBAComponentHandler instance."""
        return VBAComponentHandler()

    def test_scenario_inline_headers(self, tmp_path, handler):
        """Test Scenario 1: File with inline headers."""
        test_file = tmp_path / "Module1.bas"
        test_file.write_text(
            """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Module1"

Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        # Should detect inline headers
        assert handler.has_inline_headers(test_file) is True

        # Should split header from code
        content = test_file.read_text(encoding="utf-8")
        header, code = handler.split_vba_content(content)

        assert "VERSION 1.0 CLASS" in header
        assert "Attribute VB_Name" in header
        assert "Sub Test()" in code
        assert "VERSION" not in code

    def test_scenario_separate_header_file(self, tmp_path, handler):
        """Test Scenario 2: Code file + separate .header file."""
        # Create code file (no inline headers)
        code_file = tmp_path / "Module1.bas"
        code_file.write_text(
            """Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        # Create separate header file
        header_file = tmp_path / "Module1.header"
        header_file.write_text(
            """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Module1"
""",
            encoding="utf-8",
        )

        # Should NOT detect inline headers (they're in separate file)
        assert handler.has_inline_headers(code_file) is False

        # Header file should exist for fallback
        assert header_file.exists()

    def test_scenario_no_headers(self, tmp_path, handler):
        """Test Scenario 3: Code only, no headers anywhere."""
        code_file = tmp_path / "Module1.bas"
        code_file.write_text(
            """Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        # Should NOT detect inline headers
        assert handler.has_inline_headers(code_file) is False

        # No separate header file either
        header_file = tmp_path / "Module1.header"
        assert not header_file.exists()

        # Import logic should create minimal header in this case


class TestHeaderDetectionPerformance:
    """Test performance characteristics of header detection."""

    @pytest.fixture
    def handler(self):
        """Create a VBAComponentHandler instance."""
        return VBAComponentHandler()

    def test_only_reads_first_10_lines(self, tmp_path, handler):
        """Test that detection only reads first 10 lines for efficiency."""
        # Create a file with 1000 lines but headers only at top
        lines = [
            "VERSION 1.0 CLASS",
            "BEGIN",
            "END",
            'Attribute VB_Name = "Module1"',
            "",
        ]

        # Add 995 more lines of code
        lines.extend(f"' Comment line {i}" for i in range(995))

        test_file = tmp_path / "Large.bas"
        test_file.write_text("\n".join(lines), encoding="utf-8")

        # Should still detect headers (they're in first 10 lines)
        assert handler.has_inline_headers(test_file) is True

    def test_ignores_headers_after_line_10(self, tmp_path, handler):
        """Test that headers after line 10 are ignored."""
        # Create file with headers starting at line 20
        lines = [f"' Comment line {i}" for i in range(15)]
        lines.extend(("VERSION 1.0 CLASS", "BEGIN", "END"))

        test_file = tmp_path / "LateHeaders.bas"
        test_file.write_text("\n".join(lines), encoding="utf-8")

        # Should NOT detect headers (they're after line 10)
        assert handler.has_inline_headers(test_file) is False


class TestConflictingHeaders:
    """Test handling of conflicting header formats (inline + separate .header file)."""

    @pytest.fixture
    def handler(self):
        """Create a VBAComponentHandler instance."""
        return VBAComponentHandler()

    def test_warns_when_both_formats_exist(self, tmp_path, handler, caplog):
        """Test that a warning is logged when both inline and separate headers exist."""
        import logging
        from unittest.mock import Mock, patch
        from vba_edit.excel_vba import ExcelVBAHandler

        # Create a code file with inline headers
        code_file = tmp_path / "Module1.bas"
        code_file.write_text(
            """VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "Module1"

Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        # Create a separate header file (conflicting situation)
        header_file = tmp_path / "Module1.bas.header"
        header_file.write_text(
            """VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "Module1"
""",
            encoding="utf-8",
        )

        # Verify both exist
        assert code_file.exists()
        assert header_file.exists()
        assert handler.has_inline_headers(code_file) is True

        # Create test Excel file
        excel_file = tmp_path / "test.xlsm"
        excel_file.touch()

        # Mock the Excel handler and test import_component
        with patch("vba_edit.office_vba.OfficeVBAHandler._import_with_in_file_headers"):
            with patch("vba_edit.office_vba.OfficeVBAHandler.save_document"):
                with patch("vba_edit.excel_vba.ExcelVBAHandler._open_document_impl"):
                    office_handler = ExcelVBAHandler(doc_path=str(excel_file), vba_dir=str(tmp_path))

                    # Mock components collection
                    mock_components = Mock()

                    # Capture log messages
                    with caplog.at_level(logging.WARNING):
                        office_handler.import_component(code_file, mock_components)

                    # Check that warning was logged
                    warning_messages = [record.message for record in caplog.records if record.levelname == "WARNING"]
                    assert warning_messages
                    assert any("Both inline headers and separate header file found" in msg for msg in warning_messages)
                    assert any("Using inline headers" in msg for msg in warning_messages)
                    assert any(".header file will be ignored" in msg for msg in warning_messages)

    def test_inline_headers_take_precedence(self, tmp_path, handler):
        """Test that inline headers are used when both formats exist."""
        from unittest.mock import Mock, patch
        from vba_edit.excel_vba import ExcelVBAHandler

        # Create a code file with inline headers
        code_file = tmp_path / "TestModule.bas"
        code_file.write_text(
            """VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "TestModule"

Sub InlineTest()
End Sub
""",
            encoding="utf-8",
        )

        # Create a separate header file (should be ignored)
        header_file = tmp_path / "TestModule.bas.header"
        header_file.write_text("VERSION 5.0 CLASS\n", encoding="utf-8")

        # Verify setup
        assert code_file.exists()
        assert header_file.exists()

        # Create test Excel file
        excel_file = tmp_path / "test.xlsm"
        excel_file.touch()

        # Mock and test that inline header path is taken
        with patch("vba_edit.office_vba.OfficeVBAHandler._import_with_in_file_headers") as mock_inline:
            with patch("vba_edit.office_vba.OfficeVBAHandler._import_with_separate_headers") as mock_separate:
                with patch("vba_edit.office_vba.OfficeVBAHandler.save_document"):
                    with patch("vba_edit.excel_vba.ExcelVBAHandler._open_document_impl"):
                        office_handler = ExcelVBAHandler(doc_path=str(excel_file), vba_dir=str(tmp_path))
                        mock_components = Mock()

                        office_handler.import_component(code_file, mock_components)

                        # Verify inline import was called, separate was NOT called
                        mock_inline.assert_called_once()
                        mock_separate.assert_not_called()

    def test_no_warning_with_only_inline_headers(self, tmp_path, handler, caplog):
        """Test that no warning is logged when only inline headers exist."""
        import logging
        from unittest.mock import Mock, patch
        from vba_edit.excel_vba import ExcelVBAHandler

        # Create a code file with inline headers only
        code_file = tmp_path / "Module1.bas"
        code_file.write_text(
            """VERSION 1.0 CLASS
Attribute VB_Name = "Module1"

Sub Test()
End Sub
""",
            encoding="utf-8",
        )

        # Verify NO separate header file exists
        header_file = tmp_path / "Module1.bas.header"
        assert not header_file.exists()

        # Create test Excel file
        excel_file = tmp_path / "test.xlsm"
        excel_file.touch()

        with patch("vba_edit.office_vba.OfficeVBAHandler._import_with_in_file_headers"):
            with patch("vba_edit.office_vba.OfficeVBAHandler.save_document"):
                with patch("vba_edit.excel_vba.ExcelVBAHandler._open_document_impl"):
                    office_handler = ExcelVBAHandler(doc_path=str(excel_file), vba_dir=str(tmp_path))
                    mock_components = Mock()

                    with caplog.at_level(logging.WARNING):
                        office_handler.import_component(code_file, mock_components)

                    # Check that NO warning about conflicting headers was logged
                    warning_messages = [record.message for record in caplog.records if record.levelname == "WARNING"]
                    assert all(
                        "Both inline headers and separate header file found" not in msg for msg in warning_messages
                    )

    def test_no_warning_with_only_separate_headers(self, tmp_path, handler, caplog):
        """Test that no warning is logged when only separate header file exists."""
        import logging
        from unittest.mock import Mock, patch
        from vba_edit.excel_vba import ExcelVBAHandler

        # Create a code file WITHOUT inline headers
        code_file = tmp_path / "Module1.bas"
        code_file.write_text(
            """Sub Test()
    MsgBox "Hello"
End Sub
""",
            encoding="utf-8",
        )

        # Create separate header file
        header_file = tmp_path / "Module1.bas.header"
        header_file.write_text('Attribute VB_Name = "Module1"\n', encoding="utf-8")

        # Verify setup
        assert not handler.has_inline_headers(code_file)
        assert header_file.exists()

        # Create test Excel file
        excel_file = tmp_path / "test.xlsm"
        excel_file.touch()

        with patch("vba_edit.office_vba.OfficeVBAHandler._import_with_separate_headers"):
            with patch("vba_edit.office_vba.OfficeVBAHandler.save_document"):
                with patch("vba_edit.excel_vba.ExcelVBAHandler._open_document_impl"):
                    office_handler = ExcelVBAHandler(doc_path=str(excel_file), vba_dir=str(tmp_path))
                    mock_components = Mock()

                    with caplog.at_level(logging.WARNING):
                        office_handler.import_component(code_file, mock_components)

                    # Check that NO warning about conflicting headers was logged
                    warning_messages = [record.message for record in caplog.records if record.levelname == "WARNING"]
                    assert all(
                        "Both inline headers and separate header file found" not in msg for msg in warning_messages
                    )

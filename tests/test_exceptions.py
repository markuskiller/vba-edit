"""Tests for custom exceptions."""

import pytest

from vba_edit.exceptions import VBAError, VBAExportWarning


class TestVBAExportWarning:
    """Tests for VBAExportWarning exception."""

    def test_vbaexport_warning_creation(self):
        """Test creating VBAExportWarning with warning_type and context."""
        warning = VBAExportWarning(
            warning_type="existing_files", context={"file_count": 5, "files": ["file1.bas", "file2.cls"]}
        )

        assert warning.warning_type == "existing_files"
        assert warning.context["file_count"] == 5
        assert len(warning.context["files"]) == 2

    def test_vbaexport_warning_different_types(self):
        """Test VBAExportWarning with different warning types."""
        # Test existing_files type
        warning1 = VBAExportWarning(warning_type="existing_files", context={"file_count": 3})
        assert warning1.warning_type == "existing_files"

        # Test header_mode_changed type
        warning2 = VBAExportWarning(
            warning_type="header_mode_changed", context={"old_mode": "separate", "new_mode": "inline"}
        )
        assert warning2.warning_type == "header_mode_changed"
        assert warning2.context["old_mode"] == "separate"
        assert warning2.context["new_mode"] == "inline"

    def test_vbaexport_warning_not_vba_error(self):
        """Test that VBAExportWarning does not inherit from VBAError."""
        warning = VBAExportWarning("test_type", {})

        # Should be an Exception
        assert isinstance(warning, Exception)

        # Should NOT be a VBAError
        assert not isinstance(warning, VBAError)

    def test_vbaexport_warning_context_attribute(self):
        """Test that context attribute is stored correctly."""
        context_data = {"file_count": 10, "directory": "c:\\test\\vba", "files": ["Module1.bas", "Class1.cls"]}
        warning = VBAExportWarning("existing_files", context_data)

        assert warning.context == context_data
        assert warning.context["file_count"] == 10
        assert warning.context["directory"] == "c:\\test\\vba"

    def test_vbaexport_warning_empty_context(self):
        """Test VBAExportWarning with empty context."""
        warning = VBAExportWarning("test_type", {})
        assert warning.warning_type == "test_type"
        assert warning.context == {}

    def test_vbaexport_warning_str_representation(self):
        """Test string representation of VBAExportWarning."""
        warning = VBAExportWarning("existing_files", {"file_count": 5})
        # Should contain warning type and context info
        warning_str = str(warning)
        assert isinstance(warning_str, str)

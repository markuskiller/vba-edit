"""Unit tests for utility functions."""

import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

import pytest
from pywintypes import com_error as COM_ERROR

from vba_edit.exceptions import DocumentNotFoundError, EncodingError
from vba_edit.path_utils import get_document_paths
from vba_edit.utils import VBAFileChangeHandler, detect_vba_encoding, is_office_app_installed


def test_get_document_paths(tmp_path):
    """Test document path resolution with different inputs."""
    # Test with explicit file path
    test_doc = tmp_path / "test.docm"
    test_doc.touch()

    # Test successful path resolution
    doc_path, vba_dir = get_document_paths(str(test_doc), None)
    assert doc_path == test_doc.resolve()
    assert vba_dir == test_doc.parent.resolve()

    # Test with custom VBA directory
    vba_path = tmp_path / "vba_files"
    doc_path, vba_dir = get_document_paths(str(test_doc), None, str(vba_path))
    assert vba_dir == vba_path.resolve()
    assert vba_dir.exists()

    # Test with nonexistent file
    with pytest.raises(DocumentNotFoundError, match="Document not found"):
        get_document_paths("nonexistent.docm", None)

    # Test with no paths provided
    with pytest.raises(DocumentNotFoundError, match="No valid document path"):
        get_document_paths(None, None)


def test_is_office_app_installed_validation():
    """Test input validation of is_office_app_installed."""
    # Test invalid app names
    with pytest.raises(ValueError, match="Unsupported application"):
        is_office_app_installed("invalid_app")

    with pytest.raises(ValueError):
        is_office_app_installed("")

    # Test case insensitivity
    result = is_office_app_installed("EXCEL")
    assert isinstance(result, bool)  # Just verify return type, not actual value


def test_is_office_app_installed_mock():
    """Test Office detection with mocked COM objects."""
    with patch("win32com.client.GetActiveObject") as mock_active, patch("win32com.client.Dispatch") as mock_dispatch:
        # Test successful detection scenarios
        mock_active.return_value = Mock(Name="Excel")
        assert is_office_app_installed("excel") is True

        # Test fallback to Dispatch when no running instance
        mock_active.side_effect = COM_ERROR
        mock_app = Mock(Name="Excel")
        mock_dispatch.return_value = mock_app
        assert is_office_app_installed("word") is True

        # Test detection of non-installed apps
        mock_active.side_effect = COM_ERROR
        mock_dispatch.side_effect = COM_ERROR
        assert is_office_app_installed("powerpoint") is False


def test_vba_file_change_handler():
    """Test VBA file change handling."""
    with tempfile.TemporaryDirectory() as tmpdir:
        handler = VBAFileChangeHandler(doc_path=str(Path(tmpdir) / "test.docm"), vba_dir=tmpdir)

        # Test initialization
        assert handler.encoding == "cp1252"  # Default encoding
        assert handler.vba_dir == Path(tmpdir).resolve()


def test_detect_vba_encoding_edge_cases():
    """Test encoding detection with various content types."""
    with tempfile.TemporaryDirectory() as tmpdir:
        test_file = Path(tmpdir) / "test.bas"

        # Test empty file
        test_file.write_text("")
        with pytest.raises(EncodingError):
            detect_vba_encoding(str(test_file))

        # Test binary content
        test_file.write_bytes(b"\x00\x01\x02\x03")
        encoding, confidence = detect_vba_encoding(str(test_file))
        assert encoding is not None

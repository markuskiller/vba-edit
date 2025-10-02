"""Tests for path utilities."""

import sys
import tempfile
from pathlib import Path

import pytest

from vba_edit.exceptions import DocumentNotFoundError, PathError
from vba_edit.path_utils import (
    create_relative_path,
    resolve_path,
    validate_document_path,
)


def test_resolve_path():
    """Test basic path resolution."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)

        # Test default to cwd
        assert resolve_path(None) == Path.cwd()

        # Test relative path resolution
        assert resolve_path("test.txt", tmp_path) == (tmp_path / "test.txt").resolve()

        # Test absolute path
        assert resolve_path(tmp_path / "test.txt") == (tmp_path / "test.txt").resolve()

        # Test path with spaces
        assert resolve_path("test file.txt", tmp_path) == (tmp_path / "test file.txt").resolve()

        # Test with invalid path characters (Windows specific)
        if sys.platform.startswith("win"):
            with pytest.raises(PathError):
                resolve_path('test<>:"/\\|?*.txt')
        else:
            with pytest.raises(PathError):
                resolve_path("test\0.txt")  # Null character invalid on all platforms


def test_create_relative_path():
    """Test relative path creation."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)

        # Create directory structure
        vba_dir = tmp_path / "vba_files"
        vba_dir.mkdir()
        file_path = vba_dir / "module.bas"

        # Test basic relative path
        rel_path = create_relative_path(file_path, tmp_path)
        assert rel_path == Path("vba_files/module.bas")

        # Test with same directory
        rel_path = create_relative_path(file_path, vba_dir)
        assert rel_path == Path("module.bas")

        # Test with parent directory
        rel_path = create_relative_path(tmp_path, vba_dir)
        assert rel_path == Path("..")

        # Test with non-existent paths
        rel_path = create_relative_path(tmp_path / "nonexistent.txt", tmp_path / "other")
        assert rel_path == Path("../nonexistent.txt")

        # Test with invalid path characters (Windows specific)
        if sys.platform.startswith("win"):
            with pytest.raises(PathError):
                create_relative_path(Path('test<>:"/\\|?*.txt'), tmp_path)
        else:
            # Test with null character (invalid on all platforms)
            with pytest.raises(PathError):
                create_relative_path(Path("test\0.txt"), tmp_path)


def test_validate_document_path():
    """Test document path validation."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)

        # Test with non-existent file
        with pytest.raises(DocumentNotFoundError):
            validate_document_path(str(tmp_path / "nonexistent.docm"))

        # Test with no path
        with pytest.raises(DocumentNotFoundError):
            validate_document_path(None)

        # Test with empty string
        with pytest.raises(DocumentNotFoundError):
            validate_document_path("")

        # Test with existing file
        test_doc = tmp_path / "test.docm"
        test_doc.touch()
        result = validate_document_path(str(test_doc))
        assert result == test_doc.resolve()

        # Test with must_exist=False
        nonexistent = tmp_path / "future.docm"
        result = validate_document_path(str(nonexistent), must_exist=False)
        assert result == nonexistent.resolve()


if __name__ == "__main__":
    pytest.main(["-v", "-k", "test_path"])

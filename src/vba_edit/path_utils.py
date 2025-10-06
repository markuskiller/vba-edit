"""Utilities for handling paths across the application."""

import logging
import os
import sys
from pathlib import Path
from typing import Optional, Tuple, Union

from .exceptions import DocumentNotFoundError, PathError

logger = logging.getLogger(__name__)


def resolve_path(path: Optional[Union[str, Path]], base_path: Optional[Union[str, Path]] = None) -> Path:
    """Resolve a path, handling both absolute and relative paths.

    Args:
        path: The path to resolve. Can be None, in which case the current working directory is used
        base_path: Optional base path for resolving relative paths. If None, uses current working directory

    Returns:
        Resolved Path object

    Raises:
        PathError: If path resolution fails
    """
    try:
        if path is None:
            return Path.cwd()

        # Convert string to Path immediately to handle path separators correctly
        if isinstance(path, str):
            # Try to handle Windows-specific invalid characters in filename only
            if sys.platform.startswith("win"):
                # Get the filename part
                filename = os.path.basename(path)
                invalid_chars = '<>:"|?*'
                if any(c in filename for c in invalid_chars):
                    raise PathError(f"Filename contains invalid characters: {filename}")
            path = Path(path)

        # If path is absolute, just resolve it
        if path.is_absolute():
            return path.resolve()

        # Handle relative paths
        base = Path(base_path).resolve() if base_path else Path.cwd()
        return (base / path).resolve()

    except Exception as e:
        if isinstance(e, PathError):
            raise
        raise PathError(f"Failed to resolve path '{path}': {str(e)}") from e


def create_relative_path(path: Path, base: Path) -> Path:
    """Create a relative path from one path to another.

    Args:
        path: Path to make relative
        base: Base path to make relative to

    Returns:
        Relative Path object

    Raises:
        PathError: If relative path cannot be created
    """
    try:
        # Check for invalid characters in filenames only
        if sys.platform.startswith("win"):
            invalid_chars = '<>:"|?*'
            for p in (path, base):
                filename = p.name
                if any(c in filename for c in invalid_chars):
                    raise PathError(f"Filename contains invalid characters: {filename}")

        return Path(os.path.relpath(str(path), str(base)))
    except Exception as e:
        if isinstance(e, PathError):
            raise
        raise PathError(f"Failed to create relative path from {path} to {base}: {str(e)}") from e


def validate_document_path(doc_path: Optional[str], must_exist: bool = True) -> Path:
    """Validate and resolve a document path.

    Args:
        doc_path: Path to validate
        must_exist: If True, verify the path exists

    Returns:
        Resolved Path object

    Raises:
        DocumentNotFoundError: If path validation fails or file doesn't exist when required
    """
    try:
        if not doc_path:
            raise DocumentNotFoundError("No document path provided")

        path = resolve_path(doc_path)

        if must_exist and not path.exists():
            raise DocumentNotFoundError(f"Document not found: {path}")

        return path

    except PathError as e:
        raise DocumentNotFoundError(f"Invalid document path: {str(e)}")
    except DocumentNotFoundError:
        raise
    except Exception as e:
        raise DocumentNotFoundError(f"Failed to validate document path: {str(e)}")


def get_document_paths(
    doc_path: Optional[str], active_doc_path: Optional[str], vba_dir: Optional[str] = None
) -> Tuple[Path, Path]:
    """Get validated document and VBA directory paths.

    This is the main entry point for path resolution in the application.

    Args:
        doc_path: Explicit document path from CLI/API
        active_doc_path: Path from active Office document
        vba_dir: Optional VBA directory override

    Returns:
        Tuple of (document_path, vba_directory_path)

    Raises:
        DocumentNotFoundError: If no valid document path can be found
        PathError: If VBA directory path is invalid
    """
    logger.debug(f"get_document_paths: doc_path={doc_path}, active_doc_path={active_doc_path}, vba_dir={vba_dir}")

    # Try explicit path first
    try:
        if doc_path:
            doc_path_resolved = validate_document_path(doc_path)
        elif active_doc_path:
            doc_path_resolved = validate_document_path(active_doc_path)
        else:
            raise DocumentNotFoundError("No valid document path provided and no active document found")
    except DocumentNotFoundError as e:
        logger.debug(f"get_document_paths: document validation failed: {e}")
        if active_doc_path and not doc_path:
            # If using active document and it failed, add more context
            raise DocumentNotFoundError(f"Active document not accessible: {e}") from e
        raise

    # Resolve VBA directory
    if vba_dir:
        # Resolve vba_dir relative to cwd, not doc_path parent
        # This prevents path doubling when vba_dir contains relative paths
        vba_path = resolve_path(vba_dir)
    else:
        vba_path = doc_path_resolved.parent

    logger.debug(f"get_document_paths: resolved doc_path={doc_path_resolved}, vba_path(pre-mkdir)={vba_path}")

    # Ensure VBA directory exists
    try:
        vba_path.mkdir(parents=True, exist_ok=True)
        logger.debug(f"get_document_paths: ensured VBA directory exists: {vba_path}")
    except Exception as e:
        logger.exception("get_document_paths: failed to create/access VBA directory")
        raise PathError(f"Failed to create/access VBA directory: {e}")

    return doc_path_resolved, vba_path

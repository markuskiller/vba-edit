import ctypes
from pathlib import Path
from typing import Optional, Tuple

import chardet
import win32com.client


def get_active_office_document(app_type: str) -> str:
    """Get the path of the currently active Office document.

    Args:
        app_type (str): The Office application type ('word', 'excel', 'access', 'powerpoint')

    Returns:
        str: Full path to the active document

    Raises:
        ValueError: If invalid application type is specified
        RuntimeError: If Office application is not running or no document is active
    """
    app_type = app_type.lower()
    app_mapping = {
        "word": ("Word.Application", "Documents", "ActiveDocument"),
        "excel": ("Excel.Application", "Workbooks", "ActiveWorkbook"),
        "access": ("Access.Application", "CurrentProject", "FullName"),
        "powerpoint": ("PowerPoint.Application", "Presentations", "ActivePresentation"),
    }

    if app_type not in app_mapping:
        raise ValueError(f"Invalid application type. Must be one of: {', '.join(app_mapping.keys())}")

    app_class, collection_name, active_doc_property = app_mapping[app_type]

    try:
        app = win32com.client.GetObject(Class=app_class)

        # Special handling for Access since it uses a different pattern
        if app_type == "access":
            active_doc = getattr(app, collection_name)
            if not active_doc:
                raise RuntimeError("No Access database is currently open")
            return getattr(active_doc, active_doc_property)

        # Handle Word, Excel, and PowerPoint
        collection = getattr(app, collection_name)
        if not collection.Count:
            raise RuntimeError(f"No {app_type.capitalize()} document is currently open")

        active_doc = getattr(app, active_doc_property)
        if not active_doc:
            raise RuntimeError(f"Could not get active {app_type.capitalize()} document")

        return active_doc.FullName

    except Exception as e:
        raise RuntimeError(f"Could not connect to {app_type.capitalize()} or get active document: {e}")


def get_active_word_document() -> str:
    """Get the path of the currently active Word document.

    Returns:
        str: Full path to the active document

    Raises:
        RuntimeError: If Word is not running or no document is active
    """
    try:
        word = win32com.client.GetObject(Class="Word.Application")

        if not word.Documents.Count:
            raise RuntimeError("No Word document is currently open")

        active_doc = word.ActiveDocument
        if not active_doc:
            raise RuntimeError("Could not get active Word document")

        return active_doc.FullName
    except Exception as e:
        raise RuntimeError(f"Could not connect to Word or get active document: {e}")


def get_document_path(file_path: Optional[str] = None, app_type: str = "word") -> str:
    """Get the document path from either the provided file path or active Office document.

    Args:
        file_path: Optional path to the Office document
        app_type: Type of Office application ('word', 'excel', 'access', 'powerpoint'). Defaults to 'word'.

    Returns:
        str: Path to the document

    Raises:
        RuntimeError: If no valid document path can be determined
        ValueError: If invalid application type is specified
    """
    doc_path = file_path or get_active_office_document(app_type)

    if not Path(doc_path).exists():
        raise RuntimeError(f"Document not found: {doc_path}")

    # Convert to absolute path
    doc_path = Path(doc_path).resolve()

    return str(doc_path)


def detect_vba_encoding(file_path: str) -> Tuple[str, float]:
    """
    Detect the encoding of a VBA file using chardet.

    Args:
        file_path: Path to the file to analyze

    Returns:
        Tuple containing the detected encoding and confidence score
    """
    with open(file_path, "rb") as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        return result["encoding"], result["confidence"]


def get_windows_ansi_codepage() -> Optional[str]:
    """Get the Windows ANSI codepage as a Python encoding string.

    Returns:
        str: Python encoding name (e.g., 'cp1252') or None if not on Windows
              or if codepage couldn't be determined
    """
    try:
        # GetACP() returns the current Windows ANSI code page identifier
        codepage = ctypes.windll.kernel32.GetACP()
        return f"cp{codepage}"
    except (AttributeError, OSError):
        return None
import ctypes
import logging
import logging.handlers
import os
import sys
from functools import wraps
from pathlib import Path
from typing import Callable, Dict, Optional, Tuple

import chardet
import pywintypes
import win32com.client

from vba_edit.exceptions import (
    ApplicationError,
    DocumentClosedError,
    DocumentNotFoundError,
    EncodingError,
    OfficeError,
    PathError,
    RPCError,
    VBAAccessError,
    check_rpc_error,
)
from vba_edit.path_utils import resolve_path

# Configure module logger
logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False, logfile: Optional[str] = None) -> None:
    """Configure root logger.

    Args:
        verbose: Enable verbose (DEBUG) logging if True
        logfile: Path to log file. If None, only console logging is enabled.
    """
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    # Remove any existing handlers
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter("%(message)s"))
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    root_logger.addHandler(console_handler)

    # File handler with rotation (only if logfile is specified)
    if logfile:
        try:
            logfile_path = Path(logfile).resolve()
            file_handler = logging.handlers.RotatingFileHandler(
                logfile_path,
                maxBytes=1024 * 1024,  # 1MB
                backupCount=3,
                encoding="utf-8",
            )
            file_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
            file_handler.setLevel(logging.DEBUG)
            root_logger.addHandler(file_handler)
        except Exception as e:
            logger.warning(f"Could not set up file logging: {e}")


# COM error code for DISP_E_EXCEPTION
DISP_E_EXCEPTION = -2147352567  # 0x80020005


def is_vba_access_error(error: Exception) -> bool:
    """Check if an error is related to VBA project access being disabled.

    Args:
        error: The exception to check

    Returns:
        bool: True if error indicates VBA access is disabled
    """
    # Handle both real COM errors and our mock
    if not (hasattr(error, "args") and len(error.args) >= 3):
        return False

    # Check for DISP_E_EXCEPTION
    if error.args[0] != -2147352567:  # 0x80020005
        return False

    # Verify it's a VBA access error by checking the app specific error code
    try:
        if len(error.args) >= 3 and isinstance(error.args[2], tuple):
            scode = str(error.args[2][5])
            return any(
                c in scode
                for c in [
                    "-2146822220",  # "trust access to WORD VBA project object model"
                    "-2146827284",  # "trust access to EXCEL VBA project object model"
                    "-2147188160",  # "trust access to POWERPOINT VBA project object model"
                ]
            )
        # Catch ACCESS VBA error or any other Office app error
        elif isinstance(error.args[2], tuple) and str(error.args[2][3]).lower().endswith(".chm"):
            return True
    except (IndexError, AttributeError):
        pass

    return False


def get_vba_error_details(error: Exception) -> Dict[str, str]:
    """Extract detailed information from a VBA-related COM error.

    Args:
        error: The exception to analyze

    Returns:
        Dict containing error details with these keys:
        - hresult: COM error code
        - message: Error message
        - source: Source application
        - description: Detailed error description
        - scode: Specific error code
    """
    details = {
        "hresult": "",
        "message": "",
        "source": "",
        "description": "",
        "office_help_file": "",
        "office_help_context": "",
        "scode": "",
    }

    if not isinstance(error, pywintypes.com_error):
        logger.debug("Not a COM error")
        logger.debug(str(error))
        details["message"] = str(error)
        return details

    try:
        if len(error.args) >= 3 and isinstance(error.args[2], tuple):
            details["hresult"] = str(error.args[0])
            details["message"] = str(error.args[1])
            error_details = error.args[2]

            if len(error_details) >= 6:
                details["source"] = str(error_details[1])
                details["description"] = str(error_details[2])
                details["office_help_file"] = str(error_details[3])
                details["office_help_context"] = str(error_details[4])
                details["scode"] = str(error_details[5])

    except Exception:
        details["message"] = str(error)
        pass

    logger.debug(f"Full error details: {error}")
    logger.debug(f"COM error hresult: {details['hresult']}")
    logger.debug(f"COM error message: {details['message']}")
    logger.debug(f"VBA error source: {details['source']}")
    logger.debug(f"VBA error description: {details['description']}")
    logger.debug(f"VBA error help file: {details['office_help_file']}")
    logger.debug(f"VBA error help context: {details['office_help_context']}")
    logger.debug(f"VBA error scode: {details['scode']}")

    return details


def error_handler(func: Callable) -> Callable:
    """Decorator for consistent error handling across functions.

    Args:
        func: Function to wrap with error handling

    Returns:
        Wrapped function with error handling
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except OfficeError:
            logger.debug(f"Known error in {func.__name__}", exc_info=True)
            raise
        except Exception as e:
            logger.error(f"Unexpected error in {func.__name__}: {str(e)}")
            logger.debug("Detailed error information:", exc_info=True)
            raise OfficeError(f"Operation failed: {str(e)}") from e

    return wrapper


class VBAFileChangeHandler:
    """Handler for VBA file changes."""

    def __init__(self, doc_path: str, vba_dir: str, encoding: Optional[str] = "cp1252"):
        """Initialize the VBA file change handler.

        Args:
            doc_path: Path to the Word document
            vba_dir: Directory containing VBA files
            encoding: Character encoding for VBA files
        """
        self.doc_path = Path(doc_path).resolve()
        self.vba_dir = Path(vba_dir).resolve()
        self.encoding = encoding
        self.word = None
        self.doc = None
        self.logger = logging.getLogger(__name__)

        self.logger.debug(f"Initialized VBAFileChangeHandler with document: {self.doc_path}")
        self.logger.debug(f"VBA directory: {self.vba_dir}")
        self.logger.debug(f"Using encoding: {self.encoding}")

    @error_handler
    def import_changed_file(self, file_path: Path) -> None:
        """Import a single VBA file that has changed.

        Args:
            file_path: Path to the changed VBA file

        Raises:
            VBAAccessError: If VBA project access is denied
            DocumentClosedError: If document is closed
            RPCError: If Word application is not available
            OfficeError: For other VBA-related errors
        """
        self.logger.info(f"Processing changes in {file_path.name}")
        temp_file = None

        try:
            if self.word is None:
                self.logger.debug("Initializing Word application")
                try:
                    self.word = win32com.client.Dispatch("Word.Application")
                    self.word.Visible = True
                    self.doc = self.word.Documents.Open(str(self.doc_path))
                except Exception as e:
                    if check_rpc_error(e):
                        raise RPCError(
                            "\nWord application is not available. Please ensure Word is running and "
                            "try again with 'word-vba import' to import your changes."
                        ) from e
                    raise

            try:
                vba_project = self.doc.VBProject
            except Exception as e:
                if check_rpc_error(e):
                    raise DocumentClosedError(
                        "\nThe Word document has been closed. The edit session will be terminated.\n"
                        "IMPORTANT: Any changes made after closing the document must be imported using\n"
                        "'word-vba import' before starting a new edit session, otherwise they will be lost."
                    ) from e
                raise VBAAccessError(
                    "Cannot access VBA project. Please ensure 'Trust access to the VBA "
                    "project object model' is enabled in Word Trust Center Settings."
                ) from e

            # Read content with UTF-8 encoding (as exported)
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()

            components = vba_project.VBComponents
            component_name = file_path.stem

            if component_name == "ThisDocument":
                self.logger.debug("Processing ThisDocument module")
                doc_component = components("ThisDocument")

                # Skip header section for ThisDocument
                content_lines = content.splitlines()
                if len(content_lines) > 9:
                    actual_code = "\n".join(content_lines[9:])
                else:
                    actual_code = ""

                # Convert content to specified encoding
                content_bytes = actual_code.encode(self.encoding)
                temp_file = file_path.with_suffix(".temp")

                with open(temp_file, "wb") as f:
                    f.write(content_bytes)

                # Read back with proper encoding
                with open(temp_file, "r", encoding=self.encoding) as f:
                    new_code = f.read()

                # Update existing ThisDocument module
                self.logger.debug("Updating ThisDocument module")
                doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)
                if new_code.strip():
                    doc_component.CodeModule.AddFromString(new_code)

            else:
                self.logger.debug(f"Processing regular component: {component_name}")
                # Handle regular components
                content_bytes = content.encode(self.encoding)
                temp_file = file_path.with_suffix(".temp")

                with open(temp_file, "wb") as f:
                    f.write(content_bytes)

                # Remove existing component if it exists
                try:
                    existing = components(component_name)
                    self.logger.debug(f"Removing existing component: {component_name}")
                    components.Remove(existing)
                except Exception:
                    self.logger.debug(f"No existing component to remove: {component_name}")

                # Import the component
                self.logger.debug(f"Importing component: {component_name}")
                components.Import(str(temp_file))

            self.logger.info(f"Successfully imported: {file_path.name}")
            self.doc.Save()

        except DocumentClosedError as e:
            self.logger.error(str(e))
            # Don't raise here - let the watch loop continue
        except RPCError as e:
            self.logger.error(str(e))
            # Don't raise here - let the watch loop continue
        except Exception as e:
            self.logger.error(f"Failed to process {file_path.name}: {str(e)}")
            self.logger.debug("Detailed error information:", exc_info=True)
            raise
        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except Exception as e:
                    self.logger.warning(f"Failed to remove temporary file {temp_file}: {e}")


@error_handler
def get_active_office_document(app_type: str) -> str:
    """Get the path of the currently active Office document.

    Args:
        app_type: The Office application type ('word', 'excel', 'access', 'powerpoint')

    Returns:
        Full path to the active document

    Raises:
        ApplicationError: If Office application is not running or no document is active
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

    logger.debug(f"Getting active {app_type} document")
    app_class, collection_name, active_doc_property = app_mapping[app_type]

    try:
        app = win32com.client.GetObject(Class=app_class)

        # Special handling for Access
        if app_type == "access":
            active_doc = getattr(app, collection_name)
            if not active_doc:
                raise ApplicationError("No Access database is currently open")
            return getattr(active_doc, active_doc_property)

        elif app_type == "powerpoint":
            if not app.Windows.Count:
                raise ApplicationError("No PowerPoint presentation is currently open")

            # Get underlying presentation even if in slideshow mode
            if app.SlideShowWindows.Count > 0:
                active_doc = app.SlideShowWindows(1).View.Presentation
            else:
                active_window = app.ActiveWindow
                if not active_window:
                    raise ApplicationError("No active PowerPoint window found")
                active_doc = active_window.Presentation

            if not active_doc:
                raise ApplicationError("Could not get active PowerPoint presentation")

            return active_doc.FullName

        # Handle Word, Excel, and PowerPoint
        collection = getattr(app, collection_name)
        if not collection.Count:
            raise ApplicationError(f"No {app_type.capitalize()} document is currently open")

        active_doc = getattr(app, active_doc_property)
        if not active_doc:
            raise ApplicationError(f"Could not get active {app_type.capitalize()} document")

        doc_path = active_doc.FullName
        logger.debug(f"Found active document: {doc_path}")
        return doc_path

    except ApplicationError:
        raise
    except Exception as e:
        logger.info(f"No active {app_type.title()} document found. Use --file to specify a document path")
        logger.info(f"or open a macro-enabled document in MS {app_type.title()}.")
        raise ApplicationError(f"Could not connect to {app_type.capitalize()} or get active document: {e}")


@error_handler
def get_document_path(file_path: Optional[str] = None, app_type: str = "word") -> str:
    """Get the document path from either the provided file path or active Office document.

    Args:
        file_path: Optional path to the Office document
        app_type: Type of Office application ('word', 'excel', 'access', 'powerpoint')

    Returns:
        Path to the document

    Raises:
        DocumentNotFoundError: If document cannot be found
        ApplicationError: If Office application is not running or no document is active
    """
    logger.debug(f"Getting document path (file_path={file_path}, app_type={app_type})")

    if file_path:
        path = resolve_path(file_path)
        if not path.exists():
            raise DocumentNotFoundError(f"Document not found: {path}")
        logger.debug(f"Using provided document path: {path}")
        return str(path)

    try:
        doc_path = resolve_path(get_active_office_document(app_type))
        if not doc_path.exists():
            raise DocumentNotFoundError(f"Active document not found: {doc_path}")
        logger.debug(f"Using active document path: {doc_path}")
        return str(doc_path)
    except (ApplicationError, PathError) as e:
        raise DocumentNotFoundError(f"Could not determine document path: {e}")


def is_office_app_installed(app_name: str) -> bool:
    """Check if a specific Microsoft Office application is installed and accessible.

    Args:
        app_name: Office application name ('excel', 'word', 'access', 'powerpoint')

    Returns:
        bool: True if application is installed and accessible, False otherwise

    Examples:
        >>> is_office_app_installed('excel')
        True
        >>> is_office_app_installed('word')
        False
    """
    app_progids = {
        "excel": "Excel.Application",
        "word": "Word.Application",
        "access": "Access.Application",
        "powerpoint": "PowerPoint.Application",
    }

    app_name = app_name.lower()
    if app_name not in app_progids:
        raise ValueError(f"Unsupported application: {app_name}. Must be one of: {', '.join(app_progids.keys())}")

    app = None
    try:
        # First try to get an active instance
        try:
            win32com.client.GetActiveObject(app_progids[app_name])
            return True
        except pywintypes.com_error:
            # No active instance, try to create new one
            app = win32com.client.Dispatch(app_progids[app_name])
            # Test actual access to confirm it's working
            app.Name  # This will raise com_error if app isn't really available
            return True

    except pywintypes.com_error:
        return False
    finally:
        if app is not None:
            try:
                app.Quit()
            except pywintypes.com_error:
                pass


@error_handler
def detect_vba_encoding(file_path: str) -> Tuple[str, float]:
    """Detect the encoding of a VBA file using chardet.

    Args:
        file_path: Path to the file to analyze

    Returns:
        Tuple containing the detected encoding and confidence score

    Raises:
        EncodingError: If encoding detection fails
    """
    logger.debug(f"Detecting encoding for file: {file_path}")
    try:
        with open(file_path, "rb") as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)

            if not result["encoding"]:
                raise EncodingError(f"Could not detect encoding for file: {file_path}")

            logger.debug(f"Detected encoding: {result['encoding']} (confidence: {result['confidence']})")
            return result["encoding"], result["confidence"]
    except Exception as e:
        raise EncodingError(f"Failed to detect encoding: {e}")


@error_handler
def get_windows_ansi_codepage() -> Optional[str]:
    """Get the Windows ANSI codepage as a Python encoding string.

    Returns:
        Python encoding name (e.g., 'cp1252') or None if not on Windows
        or if codepage couldn't be determined
    """
    logger.debug("Getting Windows ANSI codepage")
    try:
        codepage = ctypes.windll.kernel32.GetACP()
        encoding = f"cp{codepage}"
        logger.debug(f"Windows ANSI codepage: {encoding}")
        return encoding
    except (AttributeError, OSError) as e:
        logger.debug(f"Could not get Windows ANSI codepage: {e}")
        return None


class OfficeApp:
    """Base class for Office application testing."""

    def __init__(self, app_name: str, prog_id: str):
        self.app_name = app_name
        self.prog_id = prog_id
        self.app = None
        self.doc = None
        self.test_file = None

    def start(self) -> None:
        """Start the Office application."""
        try:
            self.app = win32com.client.Dispatch(self.prog_id)
            if self.app_name != "access":
                self.app.Visible = True

        except Exception as e:
            logger.warning(f"Failed to start {self.app_name}: {e}")
            raise

    def cleanup(self) -> None:
        """Clean up resources."""
        try:
            if hasattr(self, "cleanup_doc"):
                self.cleanup_doc()
            if self.app and hasattr(self.app, "Quit"):
                self.app.Quit()
            if self.test_file and os.path.exists(self.test_file):
                os.unlink(self.test_file)
        except Exception as e:
            logger.error(f"Cleanup error for {self.app_name}: {e}")

    def get_vba_error(self) -> Optional[Tuple[type, str, tuple]]:
        """Get VBA access error. Should be implemented by subclasses."""
        raise NotImplementedError


class CheckWordApp(OfficeApp):
    def __init__(self):
        super().__init__("word", "Word.Application")

    def cleanup_doc(self):
        if self.doc:
            self.doc.Close(SaveChanges=False)

    def get_vba_error(self) -> Optional[Tuple[type, str, tuple]]:
        try:
            self.test_file = str(Path.cwd() / "test.docm")
            self.doc = self.app.Documents.Add()
            self.doc.SaveAs(self.test_file, FileFormat=13)  # wdFormatDocumentMacroEnabled
            _ = self.doc.VBProject
            # Test if we can actually access VBA project methods
            _ = self.doc.VBProject.VBComponents.Count
            return None
        except Exception as e:
            return (type(e), str(e), getattr(e, "args", None))


class CheckExcelApp(OfficeApp):
    def __init__(self):
        super().__init__("excel", "Excel.Application")

    def cleanup_doc(self):
        if self.doc:
            self.doc.Close(SaveChanges=False)

    def get_vba_error(self) -> Optional[Tuple[type, str, tuple]]:
        try:
            self.test_file = str(Path.cwd() / "test.xlsm")
            self.doc = self.app.Workbooks.Add()
            self.doc.SaveAs(self.test_file, FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
            _ = self.doc.VBProject
            # Test if we can actually access VBA project methods
            _ = self.doc.VBProject.VBComponents.Count
            return None
        except Exception as e:
            return (type(e), str(e), getattr(e, "args", None))


class CheckAccessApp(OfficeApp):
    def __init__(self):
        super().__init__("access", "Access.Application")

    def cleanup_doc(self):
        if self.doc:
            self.doc.Close()
        self.app.CloseCurrentDatabase()

    def get_vba_error(self) -> Optional[Tuple[type, str, tuple]]:
        try:
            self.test_file = str(Path.cwd() / "test.accdb")
            if os.path.exists(self.test_file):
                os.unlink(self.test_file)

            # Create empty database
            self.app.NewCurrentDatabase(self.test_file)
            self.doc = self.app.CurrentDb()
            _ = self.app.VBE.ActiveVBProject
            # Test if we can actually access VBA project methods
            _ = self.app.VBE.ActiveVBProject.VBComponents.Count
            return None
        except Exception as e:
            return (type(e), str(e), getattr(e, "args", None))


class CheckPowerPointApp(OfficeApp):
    def __init__(self):
        super().__init__("powerpoint", "PowerPoint.Application")

    def cleanup_doc(self):
        if self.doc:
            self.doc.Close()

    def get_vba_error(self) -> Optional[Tuple[type, str, tuple]]:
        try:
            self.test_file = str(Path.cwd() / "test.pptm")
            self.doc = self.app.Presentations.Add()
            self.doc.SaveAs(self.test_file)  # PowerPoint defaults to macro-enabled format
            _ = self.doc.VBProject
            # Test if we can actually access VBA project methods
            _ = self.doc.VBProject.VBComponents.Count
            return None
        except Exception as e:
            return (type(e), str(e), getattr(e, "args", None))


def check_office_app(app: OfficeApp) -> None:
    """Check VBA access for an Office application."""
    logger.info(f"Result of MS {app.app_name.title()} VBA project model Trust Access check:")
    try:
        app.start()
        error = app.get_vba_error()
        if error:
            error_type, error_msg, error_args = error
            logger.info(f"Error type: {error_type}")
            logger.info(f"Error message: {error_msg}")
            if error_args:
                logger.info("Error args:")
                for i, arg in enumerate(error_args):
                    logger.info(f"  {i}: {arg}")
                if isinstance(error_args, tuple) and len(error_args) > 2 and isinstance(error_args[2], tuple):
                    inner = error_args[2]
                    logger.debug("  Inner error details:")
                    for i, detail in enumerate(inner):
                        logger.debug(f"    {i}: {detail}")
            if app.app_name == "access":
                logger.warning(
                    "--> VBA project model is not accessible (make sure that database is in trusted location"
                )
                logger.warning("    and you have permissions to access it)")
            else:
                logger.warning(
                    f"--> VBA project model access seems to be disabled (enable Trust Access in the MS {app.app_name.title()} Trust Center)"
                )

        else:
            logger.info("--> VBA project model access is enabled (no further action needed)")
    except Exception as e:
        logger.warning(f"Failed to check {app.app_name}: {e}")
    finally:
        app.cleanup()


def check_vba_trust_access(app_name: Optional[str] = None) -> None:
    if not app_name:
        app_name = "Office applications"
    logger.info(f"Checking VBA Trust Access errors for MS {app_name.title()} ...")
    logger.debug(
        f"If Trust Access is disabled, this command can be used to extract the MS {app_name.title()} error messages for debugging purposes."
    )

    app_classes = {
        "word": CheckWordApp,
        "excel": CheckExcelApp,
        "access": CheckAccessApp,
        "powerpoint": CheckPowerPointApp,
    }

    if app_name and app_name != "Office applications":
        if app_name.lower() in app_classes:
            apps = [app_classes[app_name.lower()]()]
        else:
            logger.warning(f"Unsupported application name: {app_name}")
            return
    else:
        apps = [cls() for cls in app_classes.values()]

    for app in apps:
        check_office_app(app)
        logger.info("\n")


if __name__ == "__main__":
    app_name = sys.argv[1] if len(sys.argv) > 1 else None
    setup_logging(verbose=True)
    setup_logging(logfile=os.path.join(Path.cwd(), "vba_trust_access.log"))
    check_vba_trust_access(app_name)

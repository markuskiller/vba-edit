"""Tests for Office VBA handling."""

import tempfile
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import Mock, PropertyMock, patch

import pytest
import pythoncom

from vba_edit.exceptions import DocumentClosedError, DocumentNotFoundError, RPCError
from vba_edit.office_vba import (
    AccessVBAHandler,
    ExcelVBAHandler,
    VBAComponentHandler,
    VBAModuleType,
    WordVBAHandler,
)


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_vba_files(temp_dir):
    """Create sample VBA files for testing."""
    # Create standard module
    standard_module = temp_dir / "TestModule.bas"
    standard_module.write_text('Attribute VB_Name = "TestModule"\nSub Test()\n    Debug.Print "Hello"\nEnd Sub')

    # Create class module
    class_module = temp_dir / "TestClass.cls"
    class_module.write_text(
        "VERSION 1.0 CLASS\n"
        "BEGIN\n"
        "  MultiUse = -1  'True\n"
        "END\n"
        'Attribute VB_Name = "TestClass"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Attribute VB_Creatable = False\n"
        "Attribute VB_PredeclaredId = False\n"
        "Attribute VB_Exposed = False\n"
        "Public Sub TestMethod()\n"
        '    Debug.Print "Class Method"\n'
        "End Sub"
    )

    # Create document module
    doc_module = temp_dir / "ThisDocument.cls"
    doc_module.write_text(
        "VERSION 1.0 CLASS\n"
        "BEGIN\n"
        "  MultiUse = -1  'True\n"
        "END\n"
        'Attribute VB_Name = "ThisDocument"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Attribute VB_Creatable = False\n"
        "Attribute VB_PredeclaredId = True\n"
        "Attribute VB_Exposed = True\n"
        "Private Sub Document_Open()\n"
        '    Debug.Print "Document Opened"\n'
        "End Sub"
    )

    return temp_dir


class MockCOMError(Exception):
    """Mock COM error for testing without causing Windows fatal exceptions."""

    def __init__(self, hresult, text, details, helpfile=None):
        self.args = (hresult, text, details, helpfile)


@contextmanager
def com_initialized():
    """Context manager for COM initialization/cleanup."""
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


class BaseOfficeMock:
    """Base class for Office mock fixtures."""

    def __init__(self, handler_class, temp_dir, mock_document, file_extension):
        self.handler_class = handler_class
        self.temp_dir = temp_dir
        self.mock_document = mock_document
        self.file_extension = file_extension
        self.handler = None
        self.mock_app = None

    def setup(self):
        """Setup the mock handler and app."""
        doc_path = self.temp_dir / f"test{self.file_extension}"
        doc_path.touch()

        self.mock_app = Mock()
        self._configure_mock_app()

        with patch("win32com.client.Dispatch") as mock_dispatch:
            mock_dispatch.return_value = self.mock_app
            self.handler = self.handler_class(doc_path=str(doc_path), vba_dir=str(self.temp_dir))
            self.handler.app = self.mock_app
            self.handler.doc = self.mock_document

    def cleanup(self):
        """Cleanup mock objects and references."""
        if hasattr(self, "handler"):
            if hasattr(self.handler, "doc"):
                self.handler.doc = None
            if hasattr(self.handler, "app"):
                self.handler.app = None
            self.handler = None
        self.mock_app = None

    def _configure_mock_app(self):
        """Configure app-specific mock behavior. Override in subclasses."""
        raise NotImplementedError


@pytest.fixture
def mock_word_handler(temp_dir, mock_document):
    """Create a WordVBAHandler with mocked COM objects."""

    class WordMock(BaseOfficeMock):
        def _configure_mock_app(self):
            self.mock_app.Documents.Open.return_value = self.mock_document

    with com_initialized():
        mock = WordMock(WordVBAHandler, temp_dir, mock_document, ".docm")
        mock.setup()
        yield mock.handler
        mock.cleanup()


@pytest.fixture
def mock_document():
    """Create a mock document with VBA project."""
    mock_doc = Mock()
    mock_vbproj = Mock()
    mock_doc.VBProject = mock_vbproj
    mock_components = Mock()
    mock_vbproj.VBComponents = mock_components
    return mock_doc


@pytest.fixture
def mock_excel_handler(temp_dir, mock_document):
    """Create an ExcelVBAHandler with mocked COM objects."""

    class ExcelMock(BaseOfficeMock):
        def _configure_mock_app(self):
            self.mock_app.Workbooks.Open.return_value = self.mock_document

    with com_initialized():
        mock = ExcelMock(ExcelVBAHandler, temp_dir, mock_document, ".xlsm")
        mock.setup()
        yield mock.handler
        mock.cleanup()


@pytest.fixture
def mock_access_handler(temp_dir, mock_document):
    """Create an AccessVBAHandler with mocked COM objects."""

    class AccessMock(BaseOfficeMock):
        def _configure_mock_app(self):
            self.mock_app.CurrentDb.return_value = self.mock_document
            # Access-specific configuration
            self.mock_app.VBE = Mock()
            self.mock_app.VBE.ActiveVBProject = self.mock_document.VBProject

    with com_initialized():
        mock = AccessMock(AccessVBAHandler, temp_dir, mock_document, ".accdb")
        mock.setup()
        yield mock.handler
        mock.cleanup()


def create_mock_component():
    """Create a fresh mock component with code module."""
    mock_component = Mock()
    mock_code_module = Mock()
    mock_code_module.CountOfLines = 0
    mock_component.CodeModule = mock_code_module
    return mock_component, mock_code_module


def test_path_handling(temp_dir):
    """Test path handling in VBA handlers."""
    # Create test document
    doc_path = temp_dir / "test.docm"
    doc_path.touch()
    vba_dir = temp_dir / "vba"

    # Test normal initialization
    handler = WordVBAHandler(doc_path=str(doc_path), vba_dir=str(vba_dir))
    assert handler.doc_path == doc_path.resolve()
    assert handler.vba_dir == vba_dir.resolve()
    assert vba_dir.exists()

    # Test with nonexistent document
    nonexistent = temp_dir / "nonexistent.docm"
    with pytest.raises(DocumentNotFoundError) as exc_info:
        WordVBAHandler(doc_path=str(nonexistent), vba_dir=str(vba_dir))
    assert "not found" in str(exc_info.value).lower()


def test_vba_error_handling(mock_word_handler):
    """Test VBA-specific error conditions."""
    # # Create a mock COM error that simulates VBA access denied
    # mock_error = MockCOMError(
    #     -2147352567,  # DISP_E_EXCEPTION
    #     "Exception occurred",
    #     (0, "Microsoft Word", "VBA Project access is not trusted", "wdmain11.chm", 25548, -2146822220),
    #     None,
    # )

    # with patch.object(mock_word_handler.doc, "VBProject", new_callable=PropertyMock) as mock_project:
    #     # Use our mock error instead of pywintypes.com_error
    #     mock_project.side_effect = mock_error
    #     with pytest.raises(VBAAccessError) as exc_info:
    #         mock_word_handler.get_vba_project()
    #     assert "Trust access to the VBA project" in str(exc_info.value)

    # Test RPC server error
    with patch.object(mock_word_handler.doc, "Name", new_callable=PropertyMock) as mock_name:
        mock_name.side_effect = Exception("RPC server is unavailable")
        with pytest.raises(RPCError) as exc_info:
            mock_word_handler.is_document_open()
        assert "lost connection" in str(exc_info.value).lower()

    # # Test general VBA error
    # with patch.object(mock_word_handler.doc, "VBProject", new_callable=PropertyMock) as mock_project:
    #     mock_project.side_effect = Exception("Some unexpected VBA error")
    #     with pytest.raises(VBAError) as exc_info:
    #         mock_word_handler.get_vba_project()
    #     assert "wdmain11.chm" in str(exc_info.value).lower()


def test_component_handler():
    """Test VBA component handler functionality."""
    handler = VBAComponentHandler()

    # Test module type identification
    assert handler.get_module_type(Path("test.bas")) == VBAModuleType.STANDARD
    assert handler.get_module_type(Path("test.cls")) == VBAModuleType.CLASS
    assert handler.get_module_type(Path("test.frm")) == VBAModuleType.FORM
    assert handler.get_module_type(Path("ThisDocument.cls")) == VBAModuleType.DOCUMENT
    assert handler.get_module_type(Path("ThisWorkbook.cls")) == VBAModuleType.DOCUMENT
    assert handler.get_module_type(Path("Sheet1.cls")) == VBAModuleType.DOCUMENT

    # Test invalid extension
    with pytest.raises(ValueError):
        handler.get_module_type(Path("test.invalid"))


def test_component_header_handling():
    """Test VBA component header handling."""
    handler = VBAComponentHandler()

    # Test header splitting
    content = 'Attribute VB_Name = "TestModule"\nOption Explicit\nSub Test()\nEnd Sub'
    header, code = handler.split_vba_content(content)
    assert 'Attribute VB_Name = "TestModule"' in header
    assert "Option Explicit" in code
    assert "Sub Test()" in code

    # Test minimal header creation
    header = handler.create_minimal_header("TestModule", VBAModuleType.STANDARD)
    assert 'Attribute VB_Name = "TestModule"' in header

    class_header = handler.create_minimal_header("TestClass", VBAModuleType.CLASS)
    assert "VERSION 1.0 CLASS" in class_header
    assert "MultiUse = -1" in class_header


def test_word_handler_functionality(mock_word_handler, sample_vba_files):
    """Test Word VBA handler specific functionality."""
    handler = mock_word_handler

    # Test basic properties
    assert handler.app_name == "Word"
    assert handler.app_progid == "Word.Application"
    assert handler.get_document_module_name() == "ThisDocument"

    # Test document status checking
    type(handler.doc).Name = PropertyMock(return_value="test.docm")
    type(handler.doc).FullName = PropertyMock(return_value=str(handler.doc_path))
    assert handler.is_document_open()

    # Test document module update using local mocks
    mock_component, mock_code_module = create_mock_component()
    components = Mock()
    components.return_value = mock_component

    handler._update_document_module("ThisDocument", "' Test Code", components)
    mock_code_module.AddFromString.assert_called_once_with("' Test Code")


def test_excel_handler_functionality(mock_excel_handler, sample_vba_files):
    """Test Excel VBA handler specific functionality."""
    handler = mock_excel_handler

    # Test basic properties
    assert handler.app_name == "Excel"
    assert handler.app_progid == "Excel.Application"
    assert handler.get_document_module_name() == "ThisWorkbook"

    # Test document status checking
    type(handler.doc).Name = PropertyMock(return_value="test.xlsm")
    type(handler.doc).FullName = PropertyMock(return_value=str(handler.doc_path))
    assert handler.is_document_open()

    # Test workbook module update using local mocks
    mock_component, mock_code_module = create_mock_component()
    components = Mock()
    components.return_value = mock_component

    handler._update_document_module("ThisWorkbook", "' Test Code", components)
    mock_code_module.AddFromString.assert_called_once_with("' Test Code")


def test_access_handler_functionality(mock_access_handler, sample_vba_files):
    """Test Access VBA handler specific functionality."""
    handler = mock_access_handler

    # Test basic properties
    assert handler.app_name == "Access"
    assert handler.app_progid == "Access.Application"
    assert handler.get_document_module_name() == ""

    # Test database status checking
    handler.doc.Name = str(handler.doc_path)
    assert handler.is_document_open()

    # Test module update using local mocks
    mock_component, mock_code_module = create_mock_component()
    components = Mock()
    components.return_value = mock_component

    handler._update_document_module("TestModule", "' Test Code", components)
    mock_code_module.AddFromString.assert_called_once_with("' Test Code")


@pytest.mark.skip(reason="File watching too difficult to mock properly, sucessfully tested in live interaction")
def test_watch_changes_handling(mock_word_handler, temp_dir):
    """Test file watching functionality."""
    handler = mock_word_handler
    test_module = temp_dir / "TestModule.bas"
    test_module.write_text("' Test Code")

    start_time = 0

    def mock_time():
        nonlocal start_time
        start_time += 31
        return start_time

    with patch("time.time", side_effect=mock_time), patch("time.sleep"):
        handler.is_document_open = Mock(side_effect=[True, False])

        # Call the actual method
        handler.watch_changes()  # Changed from edit_vba()

        # Verify it ran
        assert handler.is_document_open.call_count >= 1


def test_watchfiles_integration():
    """Test that watchfiles is properly integrated and can be imported."""
    watchfiles = pytest.importorskip("watchfiles", reason="watchfiles not available")

    # Verify the Change enum has expected values
    assert hasattr(watchfiles.Change, "added")
    assert hasattr(watchfiles.Change, "modified")
    assert hasattr(watchfiles.Change, "deleted")


@pytest.mark.skip(reason="File watching too difficult to mock properly, sucessfully tested in live interaction")
def test_watchfiles_change_detection(mock_word_handler, temp_dir):
    """Test watchfiles change detection with mocked file changes."""
    handler = mock_word_handler
    test_module = temp_dir / "TestModule.bas"
    test_module.write_text('Attribute VB_Name = "TestModule"\nSub Test()\nEnd Sub')

    from watchfiles import Change

    mock_changes = [(Change.modified, str(test_module))]

    with patch("watchfiles.watch") as mock_watch:
        # Return iterator that yields once then stops
        mock_watch.return_value = iter([mock_changes])  # Added iter()

        handler.is_document_open = Mock(side_effect=[False])  # Exit immediately

        with patch.object(handler, "import_single_file"):
            handler.watch_changes()

        assert mock_watch.called


if __name__ == "__main__":
    pytest.main(["-v", __file__])

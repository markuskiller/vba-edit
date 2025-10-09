"""Test helpers for CLI interface testing."""

import atexit
import os
import subprocess
import time
from pathlib import Path
from typing import List, Optional

import pytest
import pythoncom
import win32com.client

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import OFFICE_MACRO_EXTENSIONS, SUPPORTED_APPS

# Global application instances
_app_instances = {}
_initialized = False


def check_office_apps_are_safe_to_use():
    """Check if Office applications (Excel, Word, PowerPoint) are running.

    Even apps running with 0 open documents are considered unsafe because:
    1. Tests will create documents in these apps
    2. Those documents might not get cleaned up properly
    3. Force-killing the app at test end loses any test-created documents

    Returns:
        tuple: (is_safe: bool, message: str) - whether it's safe to proceed and a message
    """
    # Check if override is set
    if os.environ.get("PYTEST_ALLOW_OFFICE_FORCE_QUIT") == "1":
        return True, "Safety check overridden by PYTEST_ALLOW_OFFICE_FORCE_QUIT=1"

    running_apps = []
    app_documents = {}

    # Check Excel
    try:
        app = win32com.client.GetObject(Class="Excel.Application")
        count = app.Workbooks.Count
        running_apps.append("Excel")
        app_documents["Excel"] = count
    except pythoncom.com_error:
        pass  # Not running - safe
    except Exception:
        pass  # Can't check - proceed

    # Check Word
    try:
        app = win32com.client.GetObject(Class="Word.Application")
        count = app.Documents.Count
        running_apps.append("Word")
        app_documents["Word"] = count
    except pythoncom.com_error:
        pass  # Not running - safe
    except Exception:
        pass  # Can't check - proceed

    # Check PowerPoint
    try:
        app = win32com.client.GetObject(Class="PowerPoint.Application")
        count = app.Presentations.Count
        running_apps.append("PowerPoint")
        app_documents["PowerPoint"] = count
    except pythoncom.com_error:
        pass  # Not running - safe
    except Exception:
        pass  # Can't check - proceed

    # If any Office apps are running, it's unsafe (even with 0 documents)
    if running_apps:
        message = f"\n{'=' * 70}\n⚠️  WARNING: Office applications are running!\n{'=' * 70}\n"
        for app_name in running_apps:
            doc_count = app_documents.get(app_name, 0)
            if doc_count > 0:
                message += f"  • {app_name} is running with {doc_count} open document(s)\n"
            else:
                message += f"  • {app_name} is running (no documents currently open)\n"
        message += (
            f"\n⚠️  Even apps with 0 documents are unsafe because:\n"
            f"   - Tests will create documents in these running apps\n"
            f"   - Those test documents might not get cleaned up\n"
            f"   - Force-quit at test end will lose any test-created work\n\n"
            f"These integration tests will FORCE QUIT Office apps at the end,\n"
            f"which could cause DATA LOSS if you have unsaved work.\n\n"
            f"To run these tests safely:\n"
            f"  1. Save all your Office documents\n"
            f"  2. Close ALL Office windows completely (Excel, Word, PowerPoint)\n"
            f"  3. Make sure no Office apps are in the taskbar\n"
            f"  4. Run the tests again\n\n"
            f"To override this check:\n"
            f"  Set environment variable: PYTEST_ALLOW_OFFICE_FORCE_QUIT=1\n"
            f"{'=' * 70}\n"
        )
        return False, message

    return True, "All Office applications are closed (safe to run tests)"


def check_excel_is_safe_to_use():
    """Legacy function for Excel-specific checks. Calls the comprehensive check.

    Returns:
        tuple: (is_safe: bool, message: str) - whether it's safe to proceed and a message
    """
    return check_office_apps_are_safe_to_use()


def get_or_create_app(app_name: str):
    """Get existing session application instance or create new one."""
    global _app_instances, _initialized

    if not _initialized:
        # Register cleanup on exit
        atexit.register(cleanup_all_apps)
        _initialized = True

    app_type = app_name.lower()

    # Check if we have a cached instance
    if app_type in _app_instances:
        print(f"Reusing existing {app_type} instance for test session...")
        return _app_instances[app_type]

    # Create new instance
    print(f"Creating {app_type} instance for test session...")

    # Application configurations
    app_configs = {
        "excel": "Excel.Application",
        "word": "Word.Application",
        "powerpoint": "PowerPoint.Application",
        "access": "Access.Application",
    }

    if app_type not in app_configs:
        raise ValueError(f"Unsupported application type: {app_type}")

    try:
        # Initialize COM
        pythoncom.CoInitialize()

        # Create application instance
        app = win32com.client.Dispatch(app_configs[app_type])

        # Wait for application to be ready
        _wait_for_app_ready(app, app_type)

        # Configure application
        _configure_app(app, app_type)

        _app_instances[app_type] = app
        print(f"{app_type.title()} instance ready for testing")

    except Exception as e:
        print(f"Failed to create {app_type} instance: {e}")
        raise

    return _app_instances[app_type]


def _wait_for_app_ready(app, app_type, timeout=15.0):
    """Wait for application to be ready for operations."""
    start_time = time.time()

    while time.time() - start_time < timeout:
        try:
            # Test basic property access to ensure app is ready
            _ = app.Name
            time.sleep(0.2)  # Small additional delay
            return
        except Exception:
            time.sleep(0.1)

    raise RuntimeError(f"{app_type} application not ready within {timeout} seconds")


def _configure_app(app, app_type):
    """Configure application properties for testing."""
    config_attempts = 3

    # Common properties to set
    properties = [
        ("DisplayAlerts", False),
        ("Visible", True),  # Keep visible for debugging
    ]

    # App-specific properties
    if app_type == "excel":
        properties.extend(
            [
                ("ScreenUpdating", True),
                ("EnableEvents", True),
            ]
        )
    elif app_type == "word":
        properties.extend(
            [
                ("ShowAnimation", False),
            ]
        )
    elif app_type == "access":
        # Access-specific configuration
        # Note: DoCmd.SetWarnings only works after a database is open,
        # so it's called in ReferenceDocuments.__enter__ instead

        # Try to set AutomationSecurity to disable macro warnings
        try:
            # msoAutomationSecurityLow = 1
            app.AutomationSecurity = 1
            print("Set Access AutomationSecurity to Low for testing")
        except Exception as e:
            print(f"Warning: Could not set Access AutomationSecurity: {e}")

    # Set properties with retry logic
    for prop_name, prop_value in properties:
        for attempt in range(config_attempts):
            try:
                setattr(app, prop_name, prop_value)
                break
            except Exception as e:
                if attempt == config_attempts - 1:
                    print(f"Warning: Could not set {app_type}.{prop_name}: {e}")
                else:
                    time.sleep(0.1)


def cleanup_all_apps():
    """Clean up all application instances."""
    global _app_instances

    for app_type, app in _app_instances.items():
        try:
            print(f"Closing {app_type} at end of test session...")

            # Close all documents first
            _close_all_documents(app, app_type)

            # Quit the application
            app.Quit()
            print(f"Successfully closed {app_type}")

        except Exception as e:
            print(f"Warning: Could not quit {app_type}: {e}")

    _app_instances.clear()


def _close_all_documents(app, app_type):
    """Close all open documents for an application."""
    try:
        if app_type == "excel":
            while app.Workbooks.Count > 0:
                app.Workbooks(1).Close(SaveChanges=False)
        elif app_type == "word":
            while app.Documents.Count > 0:
                app.Documents(1).Close(SaveChanges=False)
        elif app_type == "powerpoint":
            while app.Presentations.Count > 0:
                app.Presentations(1).Close()
        elif app_type == "access":
            # Access handles this differently
            try:
                app.CloseCurrentDatabase()
            except Exception:
                pass
    except Exception as e:
        print(f"Warning: Could not close all documents for {app_type}: {e}")


def force_cleanup_apps():
    """Force cleanup - useful for manual cleanup during development."""
    cleanup_all_apps()


class ReferenceDocuments:
    """Context manager for handling Office reference documents for testing purposes."""

    def __init__(self, path: Path, app_type: str):
        self.path = path
        self.app_type = app_type.lower()
        self.app = None
        self.doc = None

    def __enter__(self):
        """Open the document and create a basic VBA project."""
        try:
            app_configs = {
                "word": {
                    "doc_method": lambda app: app.Documents.Add(),
                    "save_format": 13,  # wdFormatDocumentMacroEnabled
                },
                "excel": {
                    "doc_method": lambda app: app.Workbooks.Add(),
                    "save_format": 52,  # xlOpenXMLWorkbookMacroEnabled
                },
                "powerpoint": {
                    "doc_method": lambda app: app.Presentations.Add(WithWindow=True),
                    "save_format": 25,  # ppSaveAsOpenXMLPresentationMacroEnabled
                },
                "access": {
                    "doc_method": lambda app, path: app.NewCurrentDatabase(str(path)),
                    "save_format": None,  # Access doesn't use SaveAs format parameter
                },
            }

            if self.app_type not in app_configs:
                raise ValueError(f"Unsupported application type: {self.app_type}")

            config = app_configs[self.app_type]

            # Use the session-wide application instance
            self.app = get_or_create_app(self.app_type)

            # Access handles document creation differently (must provide path upfront)
            if self.app_type == "access":
                print(f"Creating new {self.app_type} document...")
                # Close any existing database first
                try:
                    self.app.CloseCurrentDatabase()
                except Exception:
                    pass  # Ignore if no database is open
                config["doc_method"](self.app, self.path)
                # Access doesn't return a document object, we work with CurrentDb
                self.doc = None

                # Now that database is open, disable warnings
                try:
                    self.app.DoCmd.SetWarnings(False)
                    print("Disabled Access warnings for open database")
                except Exception as e:
                    print(f"Warning: Could not disable Access database warnings: {e}")
            else:
                # Create new document
                print(f"Creating new {self.app_type} document...")
                self.doc = config["doc_method"](self.app)

            try:
                # Get VBA project (Access uses different property)
                if self.app_type == "access":
                    vba_project = self.app.VBE.ActiveVBProject
                else:
                    vba_project = self.doc.VBProject

                # Add standard module with simple test code
                module = vba_project.VBComponents.Add(1)  # 1 = standard module
                module.Name = "TestModule"
                code = 'Sub Test()\n    Debug.Print "Test"\nEnd Sub'
                module.CodeModule.AddFromString(code)

                # Add class module with Rubberduck folder annotation
                class_module = vba_project.VBComponents.Add(2)  # 2 = class module
                class_module.Name = "TestClass"
                class_code = (
                    '\'@Folder("Business.Domain")\n'
                    "Option Explicit\n\n"
                    "Private m_name As String\n\n"
                    "Public Property Get Name() As String\n"
                    "    Name = m_name\n"
                    "End Property\n\n"
                    "Public Property Let Name(ByVal value As String)\n"
                    "    m_name = value\n"
                    "End Property\n\n"
                    "Public Sub Initialize()\n"
                    '    Debug.Print "TestClass initialized"\n'
                    "End Sub"
                )
                class_module.CodeModule.AddFromString(class_code)

            except Exception as ve:
                raise VBAError(
                    "Cannot access VBA project. Please ensure 'Trust access to the "
                    "VBA project object model' is enabled in Trust Center Settings."
                ) from ve

            # Save document (Access already created with path)
            if self.app_type == "access":
                # Access database already saved during NewCurrentDatabase
                pass
            else:
                self.doc.SaveAs(str(self.path), config["save_format"])

            print(f"Created and saved {self.app_type} document: {self.path}")
            return self.path

        except Exception as e:
            self._cleanup()
            raise VBAError(f"Failed to create test document: {e}") from e

    def _cleanup(self):
        """Clean up resources."""
        if hasattr(self, "doc") and self.doc:
            try:
                self.doc.Close(False)
            except Exception:
                pass
            self.doc = None
        elif self.app_type == "access" and hasattr(self, "app") and self.app:
            # Access doesn't have a doc object, close the database directly
            try:
                self.app.CloseCurrentDatabase()
            except Exception:
                pass

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Close the document but keep the application open."""
        self._cleanup()


@pytest.fixture
def temp_office_doc(tmp_path, vba_app, request):
    """Fixture providing a temporary Office document for testing."""
    extension = OFFICE_MACRO_EXTENSIONS[vba_app]

    # Use test node name to create unique filename for each test
    test_name = request.node.name.replace("[", "_").replace("]", "").replace("::", "_")
    doc_path = tmp_path / f"test_doc_{test_name}{extension}"

    # Create document, then close it before yielding
    with ReferenceDocuments(doc_path, vba_app) as _:
        pass  # Document is created and saved, will be closed on __exit__

    # Now yield the path to the closed document
    yield doc_path


class CLITester:
    """Helper class for testing CLI interfaces."""

    def __init__(self, command: str):
        """Initialize with base command.

        Args:
            command: CLI command (e.g., 'word-vba', 'excel-vba')
        """
        self.command = command
        self.app_name = command.replace("-vba", "")

    def run(self, args: List[str], input_text: Optional[str] = None) -> subprocess.CompletedProcess:
        """Run CLI command with given arguments.

        Args:
            args: List of command arguments
            input_text: Optional input to provide to command

        Returns:
            CompletedProcess instance with command results
        """
        cmd = [self.command] + args
        return subprocess.run(cmd, input=input_text.encode() if input_text else None, capture_output=True, text=True)

    def assert_success(self, args: List[str], expected_output: Optional[str] = None) -> None:
        """Assert command succeeds and optionally check output.

        Args:
            args: Command arguments
            expected_output: Optional string to check in output
        """
        result = self.run(args)
        full_output = result.stdout + result.stderr
        # Consider success if either return code is 0 or it's an empty VBA project
        success = result.returncode == 0 or "No VBA components found" in full_output
        assert success, f"Command failed with output: {full_output}"
        if expected_output:
            assert expected_output in full_output, f"Expected '{expected_output}' in output"

    def assert_error(self, args: List[str], expected_error: str) -> None:
        """Assert command fails with expected error message.

        Args:
            args: Command arguments
            expected_error: Error message to check for
        """
        result = self.run(args)
        assert result.returncode != 0, "Command should have failed"
        full_output = result.stdout + result.stderr
        assert expected_error in full_output, f"Expected error '{expected_error}' not found in output"


def get_installed_apps(selected_apps=None) -> List[str]:
    """Get list of supported apps that are installed."""
    if selected_apps is None:
        selected_apps = ["excel", "word", "access"]

    return [app for app in selected_apps if app in SUPPORTED_APPS and _check_app_available(app)]


def _check_app_available(app_name: str) -> bool:
    """Check if an Office app is available without using COM.

    Args:
        app_name: Name of Office application to check

    Returns:
        True if app is available, False otherwise
    """
    try:
        cmd = [f"{app_name}-vba", "--help"]
        result = subprocess.run(cmd, capture_output=True, text=True)
        return result.returncode == 0
    except Exception:
        return False

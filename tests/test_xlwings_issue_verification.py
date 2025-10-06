"""Unit tests verifying that vba-edit solves specific xlwings header/attribute issues.

These tests validate the `_split_header_and_code` method which is the core functionality
for parsing VBA module headers and code. This is where xlwings has known bugs.

References:
- xlwings Issue #2148: Attribute sync errors (keyboard shortcuts, procedure attributes)
  https://github.com/xlwings/xlwings/issues/2148
  Problem: Procedure-level attributes like keyboard shortcuts don't sync properly

- xlwings Issue #2088: UserForm header parsing uses hardcoded line count
  https://github.com/xlwings/xlwings/issues/2088
  Problem: Attribute VB_Exposed appears in code section causing syntax errors
"""

from pathlib import Path
from unittest.mock import patch

import pytest

from vba_edit.office_vba import OfficeVBAHandler, VBAModuleType


class TestableHandler(OfficeVBAHandler):
    """Concrete implementation for testing."""

    @property
    def app_name(self) -> str:
        return "TestApp"

    @property
    def document_type(self) -> str:
        return "TestDoc"

    @property
    def app_progid(self) -> str:
        return "Test.Application"

    def get_application(self, **kwargs):
        pass

    def get_document(self, **kwargs):
        pass

    def is_document_open(self) -> bool:
        return True

    def _open_document_impl(self, doc_path):
        pass

    def _update_document_module(self, component, code, name):
        pass

    def get_document_module_name(self, component_name: str) -> str:
        return component_name


@pytest.fixture
def handler():
    """Create handler for testing."""
    with patch("vba_edit.office_vba.get_document_paths") as mock:
        mock.return_value = (Path("test.xlsm"), Path("test_vba"))
        return TestableHandler(doc_path="test.xlsm", vba_dir="test_vba")


class TestXlwingsIssue2148_KeyboardShortcuts:
    """Tests for xlwings Issue #2148 - Keyboard shortcut attributes.

    xlwings problem: Procedure-level attributes (e.g., keyboard shortcuts) don't sync.
    vba-edit solution: Preserve all procedure-level attributes in code section.
    """

    def test_keyboard_shortcut_ctrl_l(self, handler):
        """Test Ctrl+L keyboard shortcut preservation (exact scenario from issue #2148)."""
        vba_content = """Attribute VB_Name = "Module1"

Sub MyFunction()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "l\\n14"
    MsgBox "Executed with Ctrl+L"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # Module-level attribute in header
        assert "Attribute VB_Name" in header

        # Procedure-level attribute MUST be in code (this is what xlwings loses)
        assert 'VB_ProcData.VB_Invoke_Func = "l\\n14"' in code, (
            "Keyboard shortcut attribute must be preserved in code section (xlwings Issue #2148)"
        )

    def test_multiple_keyboard_shortcuts(self, handler):
        """Test multiple procedures with different keyboard shortcuts."""
        vba_content = """Attribute VB_Name = "Module1"

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "a\\n14"
    MsgBox "Ctrl+A"
End Sub

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "b\\n14"
    MsgBox "Ctrl+B"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # Both keyboard shortcuts must be preserved
        assert "a\\n14" in code, "Ctrl+A shortcut must be preserved"
        assert "b\\n14" in code, "Ctrl+B shortcut must be preserved"

    def test_procedure_with_description_attribute(self, handler):
        """Test VB_Description attribute preservation."""
        vba_content = """Attribute VB_Name = "Module1"

Sub MyMacro()
Attribute MyMacro.VB_Description = "My custom macro"
Attribute MyMacro.VB_ProcData.VB_Invoke_Func = "m\\n14"
    MsgBox "Test"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        assert 'VB_Description = "My custom macro"' in code
        assert "VB_ProcData.VB_Invoke_Func" in code

    def test_class_module_vb_name_exported(self, handler):
        """Test that class modules include VB_Name attribute (missing in xlwings)."""
        vba_content = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub DoSomething()
    MsgBox "Class method"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # CRITICAL: VB_Name must be in header for class modules
        # This is what xlwings is missing according to issue #2148
        assert "Attribute VB_Name" in header, "Class modules MUST export Attribute VB_Name (xlwings Issue #2148 bug)"
        assert "MyClass" in header


class TestXlwingsIssue2088_UserFormHeaders:
    """Tests for xlwings Issue #2088 - UserForm header parsing.

    xlwings problem: Uses hardcoded line count, causing attributes to leak into code.
    vba-edit solution: Dynamic parsing separates header from code correctly.
    """

    def test_userform_minimal_header(self, handler):
        """Test UserForm with minimal properties (baseline test)."""
        vba_content = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    MsgBox "Init"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # All attributes MUST be in header
        assert "Attribute VB_Name" in header
        assert "Attribute VB_Exposed = False" in header

        # NO attributes should be in code
        assert "Attribute" not in code, "No attributes should leak into code section"

        # Code should start with actual VBA code
        assert code.strip().startswith("Private Sub")

    def test_userform_with_showmodal_CRITICAL(self, handler):
        """CRITICAL TEST: UserForm with ShowModal property (exact xlwings Issue #2088 scenario).

        This is the EXACT scenario that breaks in xlwings - ShowModal adds extra lines,
        xlwings' hardcoded line count cuts at wrong position, causing "Attribute VB_Exposed"
        to appear in code section as syntax error.
        """
        vba_content = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9720.001
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   14190
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_Predeclared Id = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    MsgBox "Click"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # ShowModal property must be in header
        assert "ShowModal" in header

        # CRITICAL ASSERTION: This is what xlwings gets WRONG
        assert "Attribute VB_Exposed = False" not in code, (
            "BUG FIX: Attribute VB_Exposed must NOT be in code (xlwings Issue #2088 bug)"
        )

        assert "Attribute VB_Exposed = False" in header, "Attribute VB_Exposed must be in header"

        # Code section must be clean - no attributes at all
        assert "Attribute" not in code, (
            "Code section must not contain any Attribute lines (xlwings hardcoded line count bug)"
        )

    def test_userform_with_whatsthishelp(self, handler):
        """Test UserForm with WhatsThisHelp (also triggers xlwings bug #2088)."""
        vba_content = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Test"
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    Me.Caption = "Active"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # Both extra properties should be in header
        assert "ShowModal" in header
        assert "WhatsThisHelp" in header

        # All 5 attributes must be in header
        assert header.count("Attribute") == 5

        # NO attributes in code
        assert "Attribute" not in code

    def test_userform_many_properties_stress_test(self, handler):
        """Stress test with many UserForm properties to ensure dynamic parsing works."""
        vba_content = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ComplexForm 
   BackColor       =   &H8000000F&
   Caption         =   "Complex UserForm"
   ClientHeight    =   6000
   ClientWidth     =   10000
   DrawBuffer      =   100000
   Font            =   "Tahoma"
   HelpContextID   =   1000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "ComplexForm.frx":0000
   Picture         =   "ComplexForm.frx":0100
   ScrollBars      =   3  'fmScrollBarsBoth
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "CustomTag"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Zoom            =   100
End
Attribute VB_Name = "ComplexForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Me.BackColor = vbWhite
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # Verify all properties are in header (spot check some)
        assert "BackColor" in header
        assert "DrawBuffer" in header
        assert "ScrollBars" in header
        assert "WhatsThisButton" in header
        assert "Zoom" in header

        # All 5 Attribute lines must be in header
        assert header.count("Attribute") == 5

        # Code section must be clean
        assert "Attribute" not in code
        assert "VERSION" not in code
        assert "Begin {" not in code


class TestHeaderModes:
    """Test both header storage modes handle attributes correctly."""

    def test_save_headers_mode(self, handler):
        """Test --save-headers mode (separate .header files)."""
        handler.save_headers = True

        vba_content = """Attribute VB_Name = "Module1"

Sub TestMacro()
Attribute TestMacro.VB_ProcData.VB_Invoke_Func = "t\\n14"
    MsgBox "Test"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # Module attribute in header
        assert "Attribute VB_Name" in header

        # Procedure attribute in code
        assert "VB_ProcData" in code

    def test_in_file_headers_mode(self, handler):
        """Test --in-file-headers mode (embedded headers)."""
        handler.in_file_headers = True

        vba_content = """Attribute VB_Name = "Module1"

Sub TestMacro()
Attribute TestMacro.VB_ProcData.VB_Invoke_Func = "t\\n14"
    MsgBox "Test"
End Sub
"""

        # In in-file mode, all content should be preserved
        header, code = handler.component_handler.split_vba_content(vba_content)

        full_content = header + "\n" + code
        assert "Attribute VB_Name" in full_content
        assert "VB_ProcData" in full_content


class TestRealWorldScenarios:
    """Exact scenarios from xlwings GitHub issues."""

    def test_issue_2148_exact_reported_scenario(self, handler):
        """Exact code from xlwings Issue #2148 report."""
        # User digitaisit reported this exact scenario
        vba_content = """Attribute VB_Name = "Module1"

Sub MyFunction()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "l\\n14"
    ' This macro should be callable with Ctrl+L
    MsgBox "Executed with Ctrl+L"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # The critical test: keyboard shortcut must be preserved
        full_content = header + "\n" + code
        assert 'VB_ProcData.VB_Invoke_Func = "l\\n14"' in full_content, (
            "VERIFICATION: vba-edit correctly preserves keyboard shortcuts (xlwings Issue #2148)"
        )

    def test_issue_2088_exact_reported_scenario(self, handler):
        """Exact code from xlwings Issue #2088 report."""
        # User awheelr reported this exact scenario causing syntax errors
        vba_content = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9720.001
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   14190
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    MsgBox "Test"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # The critical test: Attribute VB_Exposed must NOT be in code (xlwings puts it there)
        assert "Attribute VB_Exposed = False" not in code, (
            "VERIFICATION: vba-edit correctly separates attributes from code (xlwings Issue #2088)"
        )

        assert "Attribute VB_Exposed = False" in header, "Attribute VB_Exposed must be in header section"

        # Verify first line of code is actual code, not attribute
        code_lines = [line for line in code.split("\n") if line.strip()]
        if code_lines:
            first_line = code_lines[0]
            assert not first_line.startswith("Attribute"), f"First code line must not be Attribute, got: {first_line}"


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])

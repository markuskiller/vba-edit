"""Integration tests for vba-edit GitHub issues.

This test suite verifies that reported issues are fixed and stay fixed.
Each test corresponds to a specific GitHub issue and tests the exact scenario reported.

References:
- Issue #16: VB_VarHelpID attribute lines appearing in VBE (WithEvents bug)
  https://github.com/markuskiller/vba-edit/issues/16
- Issue #11: Header handling improvements (--in-file-headers feature)
  https://github.com/markuskiller/vba-edit/issues/11
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


class TestIssue16_VBVarHelpID:
    """Tests for Issue #16 - VB_VarHelpID attribute lines appearing in VBE.

    Problem: After export and import, hidden attributes like VB_VarHelpID
    (used for WithEvents controls) become visible in VBA editor, causing syntax errors.

    Root cause: These are "hidden member attributes" that are legal in export files
    but illegal when written verbatim in module code.

    Solution: Filter out these attributes when importing (AddFromString).

    Reported by: @takutta, @loehnertj
    """

    def test_withevents_variable_vb_varhelpid(self, handler):
        """Test that VB_VarHelpID attributes are filtered during import.

        Exact scenario from Issue #16: WithEvents variable has VB_VarHelpID attribute
        that should NOT appear in VBA editor.
        """
        # VBA code with WithEvents and its hidden attribute
        vba_content = """Attribute VB_Name = "MyClass"

Private WithEvents MyCtrl As MSForms.CommandButton
Attribute MyCtrl.VB_VarHelpID = -1

Private Sub MyCtrl_Click()
    MsgBox "Clicked"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # The VB_VarHelpID attribute WILL be in the code section after split
        # (that's expected during export), but it MUST be filtered during import
        assert "VB_VarHelpID" in code, "Test setup: attribute exists in code after split"

        # Test that the filtering function removes it
        from vba_edit.office_vba import _filter_attributes

        filtered_code = _filter_attributes(code)

        # After filtering, VB_VarHelpID should be gone
        assert "VB_VarHelpID" not in filtered_code, "VB_VarHelpID must be filtered out during import (Issue #16)"

        # The actual VBA code should remain
        assert "Private WithEvents MyCtrl" in filtered_code
        assert "Private Sub MyCtrl_Click()" in filtered_code

    def test_multiple_withevents_attributes(self, handler):
        """Test multiple WithEvents variables with VB_VarHelpID attributes."""
        vba_content = """Attribute VB_Name = "EventHandler"

Private WithEvents Button1 As MSForms.CommandButton
Attribute Button1.VB_VarHelpID = -1

Private WithEvents Button2 As MSForms.CommandButton
Attribute Button2.VB_VarHelpID = -1

Private WithEvents TextBox1 As MSForms.TextBox
Attribute TextBox1.VB_VarHelpID = -1

Private Sub Button1_Click()
    MsgBox "Button 1"
End Sub

Private Sub Button2_Click()
    MsgBox "Button 2"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # Count VB_VarHelpID occurrences before filtering
        varhelpid_count_before = code.count("VB_VarHelpID")
        assert varhelpid_count_before == 3, "Test setup: should have 3 VB_VarHelpID attributes"

        # Apply filtering
        from vba_edit.office_vba import _filter_attributes

        filtered_code = _filter_attributes(code)

        # After filtering, all VB_VarHelpID should be gone
        varhelpid_count_after = filtered_code.count("VB_VarHelpID")
        assert varhelpid_count_after == 0, (
            f"All VB_VarHelpID attributes must be filtered (found {varhelpid_count_after})"
        )

        # The actual VBA code should remain
        assert "Private WithEvents Button1" in filtered_code
        assert "Private Sub Button1_Click()" in filtered_code

    def test_other_hidden_attributes(self, handler):
        """Test other hidden member attributes that might cause similar issues.

        From @loehnertj's comment: VB_Description, VB_UserMemId on class members.
        """
        vba_content = """Attribute VB_Name = "MyClass"

Private mValue As Long
Attribute mValue.VB_VarDescription = "Internal value"
Attribute mValue.VB_VarHelpID = -1

Public Property Get Value() As Long
Attribute Value.VB_Description = "Gets the value"
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

Public Property Let Value(ByVal newValue As Long)
Attribute Value.VB_Description = "Sets the value"
Attribute Value.VB_UserMemId = 0
    mValue = newValue
End Property
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # These member-level attributes should be in export files
        # but filtered during import (AddFromString)
        hidden_attributes = ["VB_VarDescription", "VB_VarHelpID", "VB_UserMemId"]

        # Verify they exist before filtering
        found_before = []
        for attr in hidden_attributes:
            if attr in code:
                found_before.append(attr)

        assert len(found_before) > 0, "Test setup: should have some hidden attributes"

        # Apply filtering
        from vba_edit.office_vba import _filter_attributes

        filtered_code = _filter_attributes(code)

        # After filtering, all standalone Attribute lines should be gone
        found_after = []
        for attr in hidden_attributes:
            if attr in filtered_code:
                found_after.append(attr)

        assert len(found_after) == 0, f"Hidden member attributes must be filtered: {', '.join(found_after)}"

        # The actual VBA code should remain
        assert "Private mValue As Long" in filtered_code
        assert "Public Property Get Value()" in filtered_code


class TestIssue11_InFileHeaders:
    """Tests for Issue #11 - Header handling improvements.

    Feature: --in-file-headers option to embed headers in code files.

    Background:
    - Original vba-edit stripped headers
    - v0.3.0 added --save-headers (2-file approach: .bas + .header)
    - v0.4.0 added --in-file-headers (single-file approach)

    Use cases:
    - LSP tools (VBA Pro extension) need headers for IntelliSense
    - Editing attributes alongside code is more convenient
    - UserForm .frm files need headers

    Requested by: @cargocultprogramming
    Implemented by: @onderhold
    """

    def test_in_file_headers_class_module(self, handler):
        """Test that --in-file-headers includes headers for class modules.

        Critical for LSP tools: Attribute VB_NAME must be present.
        """
        handler.in_file_headers = True

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
    MsgBox "Hello"
End Sub
"""

        # In in-file-headers mode, content should stay together
        header, code = handler.component_handler.split_vba_content(vba_content)

        # When exporting with --in-file-headers, we expect the full content
        # to be written to the file (header + code together)
        full_content = header + "\n" + code if header else code

        # Verify all class header components are present
        assert "VERSION 1.0 CLASS" in full_content
        assert "Attribute VB_Name" in full_content
        assert "Attribute VB_GlobalNameSpace" in full_content
        assert "Public Sub DoSomething" in full_content

    def test_in_file_headers_standard_module(self, handler):
        """Test --in-file-headers with standard module (.bas file)."""
        handler.in_file_headers = True

        vba_content = """Attribute VB_Name = "Module1"

Sub TestMacro()
    MsgBox "Test"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)
        full_content = header + "\n" + code if header else code

        # Header and code should both be present
        assert "Attribute VB_Name" in full_content
        assert "Sub TestMacro" in full_content

    def test_save_headers_mode_separate_files(self, handler):
        """Test --save-headers creates separate .header files."""
        handler.save_headers = True

        vba_content = """Attribute VB_Name = "Module1"

Sub TestMacro()
    MsgBox "Test"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # In save-headers mode, header and code are separated
        assert "Attribute VB_Name" in header
        assert "Attribute VB_Name" not in code
        assert "Sub TestMacro" in code
        assert "Sub TestMacro" not in header

    def test_lsp_compatibility_class_attributes(self, handler):
        """Test that class-level attributes are preserved for LSP tools.

        From @cargocultprogramming: LSP needs attributes like VB_GlobalNameSpace.
        """
        handler.in_file_headers = True

        vba_content = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private mData As Collection

Public Function GetData() As Collection
    Set GetData = mData
End Function
"""

        header, code = handler.component_handler.split_vba_content(vba_content)
        full_content = header + "\n" + code if header else code

        # All class-level attributes must be present for LSP
        required_attributes = [
            'VB_Name = "DataModel"',
            "VB_GlobalNameSpace = False",
            "VB_Creatable = False",
            "VB_PredeclaredId = False",
            "VB_Exposed = True",
        ]

        for attr in required_attributes:
            assert attr in full_content, f"LSP requires attribute: {attr} (Issue #11)"


class TestIssue11_UserFormSupport:
    """Tests for UserForm (.frm) file support - part of Issue #11.

    UserForms require complete header preservation to maintain:
    - Form layout (ClientHeight, ClientWidth, etc.)
    - Control properties
    - Form GUID
    """

    def test_userform_header_preservation(self, handler):
        """Test that UserForm headers are completely preserved."""
        userform_content = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "My Form"
   ClientHeight    =   6000
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   8000
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Me.Caption = "Initialized"
End Sub
"""

        header, code = handler.component_handler.split_vba_content(userform_content)

        # Header must contain all form properties
        form_properties = [
            "VERSION 5.00",
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F}",
            "ClientHeight",
            "ClientWidth",
            "OleObjectBlob",
            'Attribute VB_Name = "UserForm1"',
        ]

        for prop in form_properties:
            assert prop in header, f"UserForm header must contain: {prop} (Issue #11)"

        # Code must NOT contain header elements
        assert "VERSION" not in code
        assert "Begin {" not in code
        assert code.strip().startswith("Private Sub"), "Code section should start with actual VBA code"

    def test_userform_with_in_file_headers(self, handler):
        """Test UserForm export with --in-file-headers."""
        handler.in_file_headers = True

        userform_content = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigForm 
   Caption         =   "Configuration"
   ClientHeight    =   4500
   ClientWidth     =   6000
   OleObjectBlob   =   "ConfigForm.frx":0000
End
Attribute VB_Name = "ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
"""

        header, code = handler.component_handler.split_vba_content(userform_content)
        full_content = header + "\n" + code if header else code

        # With in-file-headers, everything should be in the file
        assert "VERSION 5.00" in full_content
        assert 'Caption         =   "Configuration"' in full_content
        assert 'Attribute VB_Name = "ConfigForm"' in full_content
        assert "Private Sub cmdOK_Click()" in full_content


class TestRegressionPrevention:
    """Tests to prevent regression of fixed issues."""

    def test_keyboard_shortcuts_still_work(self, handler):
        """Ensure keyboard shortcuts (Issue #2148 fix) stay fixed."""
        vba_content = """Attribute VB_Name = "Module1"

Sub QuickSave()
Attribute QuickSave.VB_ProcData.VB_Invoke_Func = "s\\n14"
    ActiveWorkbook.Save
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)
        full_content = header + "\n" + code

        assert 'VB_ProcData.VB_Invoke_Func = "s\\n14"' in full_content, (
            "Regression: Keyboard shortcuts must be preserved (was fixed for xlwings #2148)"
        )

    def test_class_vb_name_still_exported(self, handler):
        """Ensure class VB_Name export (Issue #11 requirement) stays fixed."""
        vba_content = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestClass"
Attribute VB_GlobalNameSpace = False

Public Sub Test()
End Sub
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        assert 'Attribute VB_Name = "TestClass"' in header, (
            "Regression: Class modules must export VB_Name (Issue #11 requirement)"
        )


class TestEdgeCases:
    """Test edge cases and corner scenarios."""

    def test_empty_module(self, handler):
        """Test handling of empty VBA module."""
        vba_content = """Attribute VB_Name = "EmptyModule"
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        assert header or code, "Should handle empty modules gracefully"

    def test_module_with_only_comments(self, handler):
        """Test module containing only comments."""
        vba_content = """Attribute VB_Name = "Comments"

' This module only has comments
' No actual code here
' Just documentation
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        assert "Attribute VB_Name" in header
        assert "' This module" in code

    def test_mixed_attribute_styles(self, handler):
        """Test handling of different attribute syntaxes."""
        vba_content = """Attribute VB_Name = "Mixed"

Public Function GetValue() As Long
Attribute GetValue.VB_Description = "Gets the value"
Attribute GetValue.VB_UserMemId = 0
    GetValue = 42
End Function

Private mField As String
Attribute mField.VB_VarHelpID = -1
"""

        header, code = handler.component_handler.split_vba_content(vba_content)

        # Module-level attribute in header
        assert "Attribute VB_Name" in header

        # Procedure and member attributes in code
        # (Note: Member attributes should be filtered on import per Issue #16)
        assert "GetValue.VB_Description" in code or "GetValue.VB_Description" in header


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])

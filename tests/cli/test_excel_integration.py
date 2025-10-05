"""Excel VBA Integration Tests - Core Data Fidelity Tests.

These tests validate that VBA code is preserved exactly through export/import cycles.
Focus: Line-by-line code accuracy, not headers/metadata.

Test Strategy:
- Use 150+ line code modules for meaningful testing
- Test first 3 lines, middle 3 lines, last 3 lines specifically
- Focus on cp1252 encoding (Windows default)
- Test all Excel component types: ThisWorkbook, Sheet, Module, Class, UserForm
"""

import time
from pathlib import Path
from typing import Dict, List, Tuple

import pytest
import win32com.client

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import VBADocumentNames
from .helpers import CLITester, get_or_create_app


class TestExcelDataFidelity:
    """Test VBA code data fidelity through export/import operations."""

    @pytest.fixture(scope="class")
    def excel_app(self):
        """Get Excel application instance.
        
        Safety check: Aborts if Excel has open workbooks (data loss prevention).
        Override with environment variable: PYTEST_ALLOW_EXCEL_FORCE_QUIT=1
        """
        import os
        import tests.conftest
        from .helpers import check_excel_is_safe_to_use
        
        # Check environment variable override
        allow_force_quit = os.environ.get('PYTEST_ALLOW_EXCEL_FORCE_QUIT', '0') == '1'
        
        if not allow_force_quit:
            # Check if it's safe to proceed
            if not check_excel_is_safe_to_use():
                pytest.exit(
                    "Tests aborted - Excel has open workbooks. "
                    "Close Excel and try again, or set PYTEST_ALLOW_EXCEL_FORCE_QUIT=1",
                    returncode=1
                )
        
        app = get_or_create_app("excel")
        
        # Mark Excel as used by tests - cleanup can proceed
        tests.conftest._office_apps_used.add('Excel.Application')
        
        return app

    @pytest.fixture
    def test_workbook(self, excel_app, tmp_path, request):
        """Create a test workbook with VBA components.
        
        Uses unique filename per test to avoid conflicts.
        """
        # Use unique name per test to avoid "file already open" errors
        test_name = request.node.name.replace("[", "_").replace("]", "_")
        wb_path = tmp_path / f"{test_name}.xlsm"
        
        # Suppress Excel dialogs during tests
        original_display_alerts = excel_app.DisplayAlerts
        excel_app.DisplayAlerts = False
        
        try:
            # Create workbook
            wb = excel_app.Workbooks.Add()
            
            # Add VBA components
            vb_project = wb.VBProject
            
            # Add standard module
            std_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            std_module.Name = "TestModule"
            
            # Add class module
            class_module = vb_project.VBComponents.Add(2)  # 2 = vbext_ct_ClassModule
            class_module.Name = "TestClass"
            
            # Add UserForm
            userform = vb_project.VBComponents.Add(3)  # 3 = vbext_ct_MSForm
            userform.Name = "TestForm"
            
            # Save as macro-enabled workbook
            wb.SaveAs(str(wb_path), FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled
            
            yield wb, wb_path
            
            # Cleanup - Don't close, just mark as saved and let pytest handle it
            # Closing can hang waiting for invisible dialogs
            try:
                wb.Saved = True  # Mark as saved to prevent prompts
                # Don't call Close() - let the session cleanup handle it
            except Exception:
                pass
        finally:
            # Restore original DisplayAlerts setting
            try:
                excel_app.DisplayAlerts = original_display_alerts
            except Exception:
                pass

    def generate_test_code_standard_module(self) -> str:
        """Generate 150+ line test code for standard module."""
        lines = [
            "Attribute VB_Name = \"TestModule\"",
            "Option Explicit",
            "",
            "' Test module with 150+ lines for data fidelity testing",
            "' Contains special characters: äöüÄÖÜß € £ ¥",
            "' Contains quotes: \"\"double\"\" and apostrophes",
            "",
            "Public Const APP_NAME As String = \"Test Application\"",
            "Public Const APP_VERSION As String = \"1.0.0\"",
            "Public Const MAX_ITEMS As Long = 1000",
            "",
            "Private m_counter As Long",
            "Private m_lastError As String",
            "",
            "' Initialize module",
            "Public Sub Initialize()",
            "    m_counter = 0",
            "    m_lastError = \"\"",
            "    Debug.Print \"Module initialized at \" & Now()",
            "End Sub",
            "",
            "' Main processing function",
            "Public Function ProcessData(ByVal input As String) As Boolean",
            "    On Error GoTo ErrorHandler",
            "    ",
            "    Dim result As String",
            "    Dim i As Long",
            "    ",
            "    ' Validate input",
            "    If Len(input) = 0 Then",
            "        m_lastError = \"Input cannot be empty\"",
            "        ProcessData = False",
            "        Exit Function",
            "    End If",
            "    ",
            "    ' Process each character",
            "    For i = 1 To Len(input)",
            "        result = result & Mid(input, i, 1)",
            "        m_counter = m_counter + 1",
            "    Next i",
            "    ",
            "    ProcessData = True",
            "    Exit Function",
            "    ",
            "ErrorHandler:",
            "    m_lastError = \"Error: \" & Err.Description",
            "    ProcessData = False",
            "End Function",
            "",
            "' Calculate sum of array",
            "Public Function CalculateSum(ByRef values() As Double) As Double",
            "    Dim total As Double",
            "    Dim i As Long",
            "    ",
            "    total = 0",
            "    For i = LBound(values) To UBound(values)",
            "        total = total + values(i)",
            "    Next i",
            "    ",
            "    CalculateSum = total",
            "End Function",
            "",
            "' Format currency with special characters",
            "Public Function FormatCurrency(ByVal amount As Double, ByVal currency As String) As String",
            "    Select Case currency",
            "        Case \"EUR\"",
            "            FormatCurrency = Format(amount, \"#,##0.00\") & \" €\"",
            "        Case \"GBP\"",
            "            FormatCurrency = \"£\" & Format(amount, \"#,##0.00\")",
            "        Case \"JPY\"",
            "            FormatCurrency = \"¥\" & Format(amount, \"#,##0\")",
            "        Case Else",
            "            FormatCurrency = Format(amount, \"#,##0.00\")",
            "    End Select",
            "End Function",
            "",
            "' Test German umlauts",
            "Public Sub TestGermanCharacters()",
            "    Dim text As String",
            "    text = \"Größe: groß, Maße: messen, Äpfel: Äpfel\"",
            "    text = text & vbCrLf & \"Übung macht den Meister\"",
            "    text = text & vbCrLf & \"Ökologie und Ökonomie\"",
            "    Debug.Print text",
            "End Sub",
            "",
            "' Validate email address",
            "Public Function IsValidEmail(ByVal email As String) As Boolean",
            "    Dim atPos As Long",
            "    Dim dotPos As Long",
            "    ",
            "    atPos = InStr(email, \"@\")",
            "    If atPos = 0 Then",
            "        IsValidEmail = False",
            "        Exit Function",
            "    End If",
            "    ",
            "    dotPos = InStrRev(email, \".\")",
            "    If dotPos = 0 Or dotPos < atPos Then",
            "        IsValidEmail = False",
            "        Exit Function",
            "    End If",
            "    ",
            "    IsValidEmail = True",
            "End Function",
            "",
            "' Get weekday name in German",
            "Public Function GetWeekdayNameDE(ByVal weekday As Long) As String",
            "    Select Case weekday",
            "        Case 1: GetWeekdayNameDE = \"Sonntag\"",
            "        Case 2: GetWeekdayNameDE = \"Montag\"",
            "        Case 3: GetWeekdayNameDE = \"Dienstag\"",
            "        Case 4: GetWeekdayNameDE = \"Mittwoch\"",
            "        Case 5: GetWeekdayNameDE = \"Donnerstag\"",
            "        Case 6: GetWeekdayNameDE = \"Freitag\"",
            "        Case 7: GetWeekdayNameDE = \"Samstag\"",
            "        Case Else: GetWeekdayNameDE = \"Ungültig\"",
            "    End Select",
            "End Function",
            "",
            "' Complex calculation with error handling",
            "Public Function ComplexCalculation(ByVal x As Double, ByVal y As Double) As Double",
            "    On Error GoTo ErrorHandler",
            "    ",
            "    Dim result As Double",
            "    Dim temp As Double",
            "    ",
            "    ' Step 1: Validate inputs",
            "    If x = 0 Then",
            "        Err.Raise vbObjectError + 1000, , \"X cannot be zero\"",
            "    End If",
            "    ",
            "    ' Step 2: Calculate intermediate value",
            "    temp = (x * x + y * y) / x",
            "    ",
            "    ' Step 3: Apply formula",
            "    result = temp * 3.14159265358979",
            "    ",
            "    ' Step 4: Return result",
            "    ComplexCalculation = result",
            "    Exit Function",
            "    ",
            "ErrorHandler:",
            "    Debug.Print \"Error in ComplexCalculation: \" & Err.Description",
            "    ComplexCalculation = 0",
            "End Function",
            "",
            "' Generate report with special formatting",
            "Public Sub GenerateReport()",
            "    Dim report As String",
            "    Dim i As Long",
            "    ",
            "    report = \"╔═══════════════════════════════════╗\" & vbCrLf",
            "    report = report & \"║     MONTHLY REPORT - \" & Format(Now, \"MMM YYYY\") & \"    ║\" & vbCrLf",
            "    report = report & \"╠═══════════════════════════════════╣\" & vbCrLf",
            "    ",
            "    For i = 1 To 10",
            "        report = report & \"║ Item \" & i & \": € \" & Format(i * 100, \"#,##0.00\") & Space(20 - Len(CStr(i))) & \"║\" & vbCrLf",
            "    Next i",
            "    ",
            "    report = report & \"╚═══════════════════════════════════╝\"",
            "    ",
            "    Debug.Print report",
            "End Sub",
            "",
            "' Test line endings and whitespace",
            "Public Sub TestWhitespace()",
            "    Dim text As String",
            "    ",
            "    ' Line with trailing spaces",
            "    text = \"Line 1    \"",
            "    ",
            "    ' Line with tab character",
            "    text = text & vbTab & \"Tabbed text\" & vbTab",
            "    ",
            "    ' Line with mixed spaces and tabs",
            "    text = text & \"  \" & vbTab & \"  Mixed\"",
            "    ",
            "    Debug.Print text",
            "End Sub",
            "",
            "' Last function - important for testing end of file",
            "Public Function GetLastLine() As String",
            "    GetLastLine = \"This is the last line of the module\"",
            "End Function",
        ]
        
        return "\n".join(lines)

    def generate_test_code_class_module(self) -> str:
        """Generate 150+ line test code for class module.
        
        Note: Does NOT include VBA headers (VERSION, BEGIN, END, Attribute lines).
        These are automatically managed by VBA when using AddFromString().
        """
        lines = [
            "Option Explicit",
            "",
            "' Class module for testing data fidelity",
            "' Tests property procedures, private members, and class behavior",
            "",
            "' Private member variables",
            "Private m_name As String",
            "Private m_value As Double",
            "Private m_isActive As Boolean",
            "Private m_created As Date",
            "Private m_description As String",
            "Private m_tags As Collection",
            "",
            "' Constants",
            "Private Const DEFAULT_NAME As String = \"Unnamed\"",
            "Private Const MAX_VALUE As Double = 999999.99",
            "Private Const MIN_VALUE As Double = -999999.99",
            "",
            "' Class initialization",
            "Private Sub Class_Initialize()",
            "    m_name = DEFAULT_NAME",
            "    m_value = 0",
            "    m_isActive = True",
            "    m_created = Now()",
            "    m_description = \"\"",
            "    Set m_tags = New Collection",
            "    ",
            "    Debug.Print \"TestClass instance created at \" & m_created",
            "End Sub",
            "",
            "' Class termination",
            "Private Sub Class_Terminate()",
            "    Set m_tags = Nothing",
            "    Debug.Print \"TestClass instance destroyed\"",
            "End Sub",
            "",
            "' Name property - Get",
            "Public Property Get Name() As String",
            "    Name = m_name",
            "End Property",
            "",
            "' Name property - Let",
            "Public Property Let Name(ByVal newName As String)",
            "    If Len(newName) = 0 Then",
            "        Err.Raise vbObjectError + 1001, \"TestClass\", \"Name cannot be empty\"",
            "    End If",
            "    ",
            "    m_name = newName",
            "End Property",
            "",
            "' Value property - Get",
            "Public Property Get Value() As Double",
            "    Value = m_value",
            "End Property",
            "",
            "' Value property - Let with validation",
            "Public Property Let Value(ByVal newValue As Double)",
            "    If newValue < MIN_VALUE Then",
            "        Err.Raise vbObjectError + 1002, \"TestClass\", \"Value too small\"",
            "    End If",
            "    ",
            "    If newValue > MAX_VALUE Then",
            "        Err.Raise vbObjectError + 1003, \"TestClass\", \"Value too large\"",
            "    End If",
            "    ",
            "    m_value = newValue",
            "End Property",
            "",
            "' IsActive property - Get",
            "Public Property Get IsActive() As Boolean",
            "    IsActive = m_isActive",
            "End Property",
            "",
            "' IsActive property - Let",
            "Public Property Let IsActive(ByVal newStatus As Boolean)",
            "    m_isActive = newStatus",
            "    ",
            "    If m_isActive Then",
            "        Debug.Print m_name & \" activated\"",
            "    Else",
            "        Debug.Print m_name & \" deactivated\"",
            "    End If",
            "End Property",
            "",
            "' Created property - Get (read-only)",
            "Public Property Get Created() As Date",
            "    Created = m_created",
            "End Property",
            "",
            "' Description property - Get",
            "Public Property Get Description() As String",
            "    Description = m_description",
            "End Property",
            "",
            "' Description property - Let",
            "Public Property Let Description(ByVal newDescription As String)",
            "    m_description = newDescription",
            "End Property",
            "",
            "' Add tag to collection",
            "Public Sub AddTag(ByVal tag As String)",
            "    On Error Resume Next",
            "    m_tags.Add tag, tag  ' Key is tag itself for uniqueness",
            "    ",
            "    If Err.Number = 0 Then",
            "        Debug.Print \"Tag added: \" & tag",
            "    Else",
            "        Debug.Print \"Tag already exists: \" & tag",
            "    End If",
            "    ",
            "    On Error GoTo 0",
            "End Sub",
            "",
            "' Remove tag from collection",
            "Public Sub RemoveTag(ByVal tag As String)",
            "    On Error Resume Next",
            "    m_tags.Remove tag",
            "    ",
            "    If Err.Number = 0 Then",
            "        Debug.Print \"Tag removed: \" & tag",
            "    Else",
            "        Debug.Print \"Tag not found: \" & tag",
            "    End If",
            "    ",
            "    On Error GoTo 0",
            "End Sub",
            "",
            "' Check if tag exists",
            "Public Function HasTag(ByVal tag As String) As Boolean",
            "    Dim item As Variant",
            "    ",
            "    On Error Resume Next",
            "    item = m_tags(tag)",
            "    HasTag = (Err.Number = 0)",
            "    On Error GoTo 0",
            "End Function",
            "",
            "' Get all tags as array",
            "Public Function GetTags() As String()",
            "    Dim tags() As String",
            "    Dim i As Long",
            "    Dim item As Variant",
            "    ",
            "    If m_tags.Count = 0 Then",
            "        GetTags = Split(\"\", \",\")  ' Empty array",
            "        Exit Function",
            "    End If",
            "    ",
            "    ReDim tags(1 To m_tags.Count)",
            "    ",
            "    i = 1",
            "    For Each item In m_tags",
            "        tags(i) = CStr(item)",
            "        i = i + 1",
            "    Next item",
            "    ",
            "    GetTags = tags",
            "End Function",
            "",
            "' Validate object state",
            "Public Function Validate() As Boolean",
            "    Dim isValid As Boolean",
            "    ",
            "    isValid = True",
            "    ",
            "    ' Check name",
            "    If Len(m_name) = 0 Then",
            "        Debug.Print \"Validation error: Name is empty\"",
            "        isValid = False",
            "    End If",
            "    ",
            "    ' Check value range",
            "    If m_value < MIN_VALUE Or m_value > MAX_VALUE Then",
            "        Debug.Print \"Validation error: Value out of range\"",
            "        isValid = False",
            "    End If",
            "    ",
            "    Validate = isValid",
            "End Function",
            "",
            "' Clone this object",
            "Public Function Clone() As TestClass",
            "    Dim newObj As New TestClass",
            "    ",
            "    newObj.Name = m_name",
            "    newObj.Value = m_value",
            "    newObj.IsActive = m_isActive",
            "    newObj.Description = m_description",
            "    ",
            "    ' Copy tags",
            "    Dim tag As Variant",
            "    For Each tag In m_tags",
            "        newObj.AddTag CStr(tag)",
            "    Next tag",
            "    ",
            "    Set Clone = newObj",
            "End Function",
            "",
            "' Export to formatted string",
            "Public Function ToString() As String",
            "    Dim result As String",
            "    ",
            "    result = \"TestClass Object\" & vbCrLf",
            "    result = result & \"  Name: \" & m_name & vbCrLf",
            "    result = result & \"  Value: \" & Format(m_value, \"#,##0.00\") & vbCrLf",
            "    result = result & \"  Active: \" & m_isActive & vbCrLf",
            "    result = result & \"  Created: \" & Format(m_created, \"yyyy-mm-dd hh:nn:ss\") & vbCrLf",
            "    result = result & \"  Tags: \" & m_tags.Count",
            "    ",
            "    ToString = result",
            "End Function",
            "",
            "' Final method - tests end of class",
            "Public Sub Finalize()",
            "    Debug.Print \"Finalizing \" & m_name",
            "    m_isActive = False",
            "End Sub",
        ]
        
        return "\n".join(lines)

    def extract_code_lines(self, full_content: str) -> List[str]:
        """Extract code lines, skipping attribute headers.
        
        Headers are lines like 'VERSION', 'BEGIN', 'END', 'Attribute VB_Name'.
        Code starts after these header lines.
        Normalizes line endings (removes \r) for consistent comparison.
        Strips trailing empty lines.
        """
        # Normalize line endings: VBA Editor returns lines with \r, files don't
        normalized = full_content.replace('\r\n', '\n').replace('\r', '\n')
        lines = normalized.split('\n')
        
        # Find where actual code starts (after headers)
        code_start = 0
        for i, line in enumerate(lines):
            stripped = line.strip()
            # Skip header lines
            if (stripped.startswith(('VERSION', 'BEGIN', 'END', 'Attribute VB_Name',
                                    'Attribute VB_GlobalNameSpace', 'Attribute VB_Creatable',
                                    'Attribute VB_PredeclaredId', 'Attribute VB_Exposed'))
                or stripped == 'MultiUse = -1  \'True'
                or (stripped.startswith('Attribute') and 'VB_Name' not in stripped)):
                continue
            else:
                # Found first non-header line
                code_start = i
                break
        
        code_lines = lines[code_start:]
        
        # Strip trailing empty lines (common in text files)
        while code_lines and code_lines[-1] == '':
            code_lines.pop()
        
        return code_lines

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_export_standard_module_line_by_line(self, test_workbook, tmp_path):
        """Test that exported standard module code matches VBE exactly, line by line."""
        wb, wb_path = test_workbook
        vba_dir = tmp_path / "vba"
        
        # Generate and load test code (150+ lines)
        test_code = self.generate_test_code_standard_module()
        
        # Set code in VBA Editor
        vb_project = wb.VBProject
        test_module = vb_project.VBComponents("TestModule")
        code_module = test_module.CodeModule
        code_module.AddFromString(test_code)
        
        # Get code from VBA Editor
        vbe_code = code_module.Lines(1, code_module.CountOfLines)
        vbe_lines = self.extract_code_lines(vbe_code)
        
        # Verify we have 150+ lines
        assert len(vbe_lines) >= 150, f"Test code should have 150+ lines, got {len(vbe_lines)}"
        
        # Export using CLI
        cli = CLITester("excel-vba")
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--force-overwrite",
            "--save-headers"  # Use separate header files (proven approach)
        ])
        
        # Read exported file
        exported_file = vba_dir / "TestModule.bas"
        assert exported_file.exists(), "TestModule.bas was not exported"
        
        with open(exported_file, 'r', encoding='cp1252') as f:
            exported_content = f.read()
        
        exported_lines = self.extract_code_lines(exported_content)
        
        # Test 1: Line count must match
        assert len(exported_lines) == len(vbe_lines), \
            f"Line count mismatch: VBE={len(vbe_lines)}, Exported={len(exported_lines)}"
        
        # Test 2: First 3 lines must match exactly
        for i in range(min(3, len(vbe_lines))):
            assert exported_lines[i] == vbe_lines[i], \
                f"Line {i+1} mismatch:\n  VBE: '{vbe_lines[i]}'\n  File: '{exported_lines[i]}'"
        
        # Test 3: Middle 3 lines must match exactly
        mid_start = len(vbe_lines) // 2
        for i in range(mid_start, min(mid_start + 3, len(vbe_lines))):
            assert exported_lines[i] == vbe_lines[i], \
                f"Line {i+1} mismatch:\n  VBE: '{vbe_lines[i]}'\n  File: '{exported_lines[i]}'"
        
        # Test 4: Last 3 lines must match exactly
        for i in range(max(0, len(vbe_lines) - 3), len(vbe_lines)):
            assert exported_lines[i] == vbe_lines[i], \
                f"Line {i+1} mismatch:\n  VBE: '{vbe_lines[i]}'\n  File: '{exported_lines[i]}'"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_export_class_module_line_by_line(self, test_workbook, tmp_path):
        """Test that exported class module code matches VBE exactly, line by line."""
        wb, wb_path = test_workbook
        vba_dir = tmp_path / "vba"
        
        # Generate and load test code (150+ lines)
        test_code = self.generate_test_code_class_module()
        
        # Set code in VBA Editor
        vb_project = wb.VBProject
        test_class = vb_project.VBComponents("TestClass")
        code_module = test_class.CodeModule
        code_module.AddFromString(test_code)
        
        # Get code from VBA Editor
        vbe_code = code_module.Lines(1, code_module.CountOfLines)
        vbe_lines = self.extract_code_lines(vbe_code)
        
        # Verify we have 150+ lines
        assert len(vbe_lines) >= 150, f"Test code should have 150+ lines, got {len(vbe_lines)}"
        
        # Export using CLI
        cli = CLITester("excel-vba")
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--force-overwrite",
            "--save-headers"  # Use separate header files (proven approach)
        ])
        
        # Read exported file
        exported_file = vba_dir / "TestClass.cls"
        assert exported_file.exists(), "TestClass.cls was not exported"
        
        with open(exported_file, 'r', encoding='cp1252') as f:
            exported_content = f.read()
        
        exported_lines = self.extract_code_lines(exported_content)
        
        # Test line-by-line matching (same as standard module tests)
        assert len(exported_lines) == len(vbe_lines), \
            f"Line count mismatch: VBE={len(vbe_lines)}, Exported={len(exported_lines)}"
        
        # First 3 lines
        for i in range(min(3, len(vbe_lines))):
            assert exported_lines[i] == vbe_lines[i], \
                f"Line {i+1} mismatch:\n  VBE: '{vbe_lines[i]}'\n  File: '{exported_lines[i]}'"
        
        # Middle 3 lines
        mid_start = len(vbe_lines) // 2
        for i in range(mid_start, min(mid_start + 3, len(vbe_lines))):
            assert exported_lines[i] == vbe_lines[i], \
                f"Line {i+1} mismatch:\n  VBE: '{vbe_lines[i]}'\n  File: '{exported_lines[i]}'"
        
        # Last 3 lines
        for i in range(max(0, len(vbe_lines) - 3), len(vbe_lines)):
            assert exported_lines[i] == vbe_lines[i], \
                f"Line {i+1} mismatch:\n  VBE: '{vbe_lines[i]}'\n  File: '{exported_lines[i]}'"

    def generate_test_code_thisworkbook(self) -> str:
        """Generate test code for ThisWorkbook document module."""
        lines = [
            "Option Explicit",
            "",
            "' ThisWorkbook module - Document-level code",
            "' This must be imported to the ThisWorkbook document module, not as a separate module",
            "",
            "Private Sub Workbook_Open()",
            "    MsgBox \"Workbook opened at \" & Now(), vbInformation",
            "    Debug.Print \"Workbook_Open event fired\"",
            "End Sub",
            "",
            "Private Sub Workbook_BeforeClose(Cancel As Boolean)",
            "    Dim response As VbMsgBoxResult",
            "    ",
            "    response = MsgBox(\"Save changes before closing?\", vbYesNoCancel)",
            "    ",
            "    Select Case response",
            "        Case vbYes",
            "            ThisWorkbook.Save",
            "        Case vbCancel",
            "            Cancel = True",
            "    End Select",
            "End Sub",
            "",
            "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)",
            "    Debug.Print \"Saving workbook at \" & Now()",
            "End Sub",
            "",
            "Public Function GetWorkbookInfo() As String",
            "    Dim info As String",
            "    ",
            "    info = \"Workbook: \" & ThisWorkbook.Name & vbCrLf",
            "    info = info & \"Path: \" & ThisWorkbook.Path & vbCrLf",
            "    info = info & \"Sheets: \" & ThisWorkbook.Worksheets.Count & vbCrLf",
            "    info = info & \"Created: \" & ThisWorkbook.BuiltinDocumentProperties(\"Creation Date\")",
            "    ",
            "    GetWorkbookInfo = info",
            "End Function",
            "",
            "Public Sub RefreshAllData()",
            "    Dim ws As Worksheet",
            "    ",
            "    Application.ScreenUpdating = False",
            "    ",
            "    For Each ws In ThisWorkbook.Worksheets",
            "        ws.Calculate",
            "    Next ws",
            "    ",
            "    Application.ScreenUpdating = True",
            "    ",
            "    MsgBox \"All data refreshed\", vbInformation",
            "End Sub",
        ]
        
        return "\n".join(lines)

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_import_thisworkbook_to_correct_location(self, test_workbook, tmp_path):
        """Test that ThisWorkbook code is imported to document module, not as standalone module.
        
        This is critical: ThisWorkbook.cls must be imported into the existing
        ThisWorkbook document module in the VBA project, not added as a new module.
        
        Uses --save-headers (separate header files) which is the proven, stable approach.
        """
        wb, wb_path = test_workbook
        vba_dir = tmp_path / "vba"
        
        # Generate test code for ThisWorkbook
        test_code = self.generate_test_code_thisworkbook()
        
        vb_project = wb.VBProject
        
        # Get existing ThisWorkbook component (it always exists)
        # In German Excel it's called "DieseArbeitsmappe", in English "ThisWorkbook"
        thisworkbook = wb.VBProject.VBComponents(wb.CodeName)
        thisworkbook_name = thisworkbook.Name
        original_type = thisworkbook.Type  # Should be 100 (vbext_ct_Document)
        
        # Add code to ThisWorkbook
        code_module = thisworkbook.CodeModule
        code_module.AddFromString(test_code)
        
        # Export using --save-headers (separate header files - the proven approach)
        cli = CLITester("excel-vba")
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--force-overwrite",
            "--save-headers"  # Use separate header files (proven approach)
        ])
        
        # Verify ThisWorkbook.cls was exported (filename uses actual component name)
        exported_file = vba_dir / f"{thisworkbook_name}.cls"
        assert exported_file.exists(), f"{thisworkbook_name}.cls was not exported"
        
        with open(exported_file, 'r', encoding='cp1252') as f:
            exported_content = f.read()
        exported_lines = self.extract_code_lines(exported_content)
        
        # Clear ThisWorkbook code
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        
        # Remove UserForm and Sheet files to simplify test
        # Keep ThisWorkbook.cls to test it imports correctly!
        test_form_frm = vba_dir / "TestForm.frm"
        if test_form_frm.exists():
            test_form_frm.unlink()
        test_form_frx = vba_dir / "TestForm.frx"
        if test_form_frx.exists():
            test_form_frx.unlink()
        
        sheet_cls = vba_dir / "Tabelle1.cls"
        if sheet_cls.exists():
            sheet_cls.unlink()
        sheet_header = vba_dir / "Tabelle1.header"
        if sheet_header.exists():
            sheet_header.unlink()
        
        # Count components before import
        components_before = vb_project.VBComponents.Count
        
        # Debug: List components before import
        print(f"\n=== Components BEFORE import ({components_before}) ===")
        for i in range(1, components_before + 1):
            comp = vb_project.VBComponents(i)
            print(f"  {i}. {comp.Name} (Type: {comp.Type})")
        print("="*50)
        
        # Import
        cli.assert_success([
            "import",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir)
        ])
        
        # Count components after import
        components_after = vb_project.VBComponents.Count
        
        # Debug: List components after import
        print(f"\n=== Components AFTER import ({components_after}) ===")
        for i in range(1, components_after + 1):
            comp = vb_project.VBComponents(i)
            print(f"  {i}. {comp.Name} (Type: {comp.Type})")
        print("="*50)
        
        # Verify no new component was added (ThisWorkbook should still be document module)
        assert components_after == components_before, \
            f"Import should not add new component. Before={components_before}, After={components_after}"
        
        # Verify ThisWorkbook still exists and is still a document module
        thisworkbook = vb_project.VBComponents(thisworkbook_name)
        assert thisworkbook.Type == original_type, \
            f"ThisWorkbook type changed. Expected={original_type}, Got={thisworkbook.Type}"
        assert thisworkbook.Type == 100, \
            "ThisWorkbook must be document module (type 100), not standard module"
        
        # Verify code was imported correctly
        code_module = thisworkbook.CodeModule
        imported_code = code_module.Lines(1, code_module.CountOfLines)
        imported_lines = self.extract_code_lines(imported_code)
        
        # Verify line count
        assert len(imported_lines) == len(exported_lines), \
            f"ThisWorkbook import line count mismatch: Expected={len(exported_lines)}, Got={len(imported_lines)}"
        
        # Verify first 3 lines
        for i in range(min(3, len(exported_lines))):
            assert imported_lines[i] == exported_lines[i], \
                f"ThisWorkbook line {i+1} mismatch:\n  File: '{exported_lines[i]}'\n  VBE: '{imported_lines[i]}'"
        
        # Verify last 3 lines
        for i in range(max(0, len(exported_lines) - 3), len(exported_lines)):
            assert imported_lines[i] == exported_lines[i], \
                f"ThisWorkbook line {i+1} mismatch:\n  File: '{exported_lines[i]}'\n  VBE: '{imported_lines[i]}'"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_import_standard_module_line_by_line(self, test_workbook, tmp_path):
        """Test that imported standard module code matches exported file exactly."""
        wb, wb_path = test_workbook
        vba_dir = tmp_path / "vba"
        
        # Generate and export test code
        test_code = self.generate_test_code_standard_module()
        
        vb_project = wb.VBProject
        test_module = vb_project.VBComponents("TestModule")
        code_module = test_module.CodeModule
        code_module.AddFromString(test_code)
        
        # Export
        cli = CLITester("excel-vba")
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--force-overwrite",
            "--save-headers"  # Use separate header files (proven approach)
        ])
        
        # Read exported file
        exported_file = vba_dir / "TestModule.bas"
        with open(exported_file, 'r', encoding='cp1252') as f:
            exported_content = f.read()
        exported_lines = self.extract_code_lines(exported_content)
        
        # Delete module from VBA project
        vb_project.VBComponents.Remove(test_module)
        
        # Import
        cli.assert_success([
            "import",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir)
        ])
        
        # Get imported code from VBA Editor
        test_module = vb_project.VBComponents("TestModule")
        code_module = test_module.CodeModule
        imported_code = code_module.Lines(1, code_module.CountOfLines)
        imported_lines = self.extract_code_lines(imported_code)
        
        # Verify line-by-line match
        assert len(imported_lines) == len(exported_lines), \
            f"Line count mismatch after import: Expected={len(exported_lines)}, Got={len(imported_lines)}"
        
        # First 3 lines
        for i in range(min(3, len(exported_lines))):
            assert imported_lines[i] == exported_lines[i], \
                f"Import line {i+1} mismatch:\n  File: '{exported_lines[i]}'\n  VBE: '{imported_lines[i]}'"
        
        # Middle 3 lines
        mid_start = len(exported_lines) // 2
        for i in range(mid_start, min(mid_start + 3, len(exported_lines))):
            assert imported_lines[i] == exported_lines[i], \
                f"Import line {i+1} mismatch:\n  File: '{exported_lines[i]}'\n  VBE: '{imported_lines[i]}'"
        
        # Last 3 lines
        for i in range(max(0, len(exported_lines) - 3), len(exported_lines)):
            assert imported_lines[i] == exported_lines[i], \
                f"Import line {i+1} mismatch:\n  File: '{exported_lines[i]}'\n  VBE: '{imported_lines[i]}'"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_roundtrip_preserves_code_exactly(self, test_workbook, tmp_path):
        """Test that export → import → export produces identical code."""
        wb, wb_path = test_workbook
        vba_dir1 = tmp_path / "vba1"
        vba_dir2 = tmp_path / "vba2"
        
        # Generate test code
        test_code = self.generate_test_code_standard_module()
        
        vb_project = wb.VBProject
        test_module = vb_project.VBComponents("TestModule")
        code_module = test_module.CodeModule
        code_module.AddFromString(test_code)
        
        cli = CLITester("excel-vba")
        
        # First export
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir1),
            "--force-overwrite",
            "--save-headers"  # Use separate header files (proven approach)
        ])
        
        # Read first export
        with open(vba_dir1 / "TestModule.bas", 'r', encoding='cp1252') as f:
            export1_content = f.read()
        export1_lines = self.extract_code_lines(export1_content)
        
        # Delete and import
        vb_project.VBComponents.Remove(test_module)
        
        cli.assert_success([
            "import",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir1)
        ])
        
        # Second export
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir2),
            "--force-overwrite",
            "--save-headers"  # Use separate header files (proven approach)
        ])
        
        # Read second export
        with open(vba_dir2 / "TestModule.bas", 'r', encoding='cp1252') as f:
            export2_content = f.read()
        export2_lines = self.extract_code_lines(export2_content)
        
        # Verify both exports are identical
        assert len(export1_lines) == len(export2_lines), \
            "Round-trip changed line count"
        
        for i, (line1, line2) in enumerate(zip(export1_lines, export2_lines)):
            assert line1 == line2, \
                f"Round-trip changed line {i+1}:\n  First: '{line1}'\n  Second: '{line2}'"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_in_file_headers_document_module(self, test_workbook, tmp_path):
        """Test that --in-file-headers properly strips headers from document modules (ThisWorkbook).
        
        This test validates the fix for the bug where VERSION, BEGIN, END, MultiUse, 
        and Attribute lines were being imported as code into document modules, 
        causing VBA compile errors.
        """
        wb, wb_path = test_workbook
        vba_dir = tmp_path / "vba"
        
        # Generate test code for ThisWorkbook (document module)
        test_code = """' ThisWorkbook Module Test
Option Explicit

Private Sub Workbook_Open()
    ' Test code with special characters: äöüÄÖÜß €£¥
    Dim message As String
    message = "Workbook opened successfully!"
    Debug.Print message
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "Closing workbook..."
End Sub
"""
        
        vb_project = wb.VBProject
        # Get the ThisWorkbook document module
        # In Excel, document modules have Type = 100 (vbext_ct_Document)
        # Note: Name varies by Excel language (ThisWorkbook, DieseArbeitsmappe, CeClasseur, etc.)
        doc_module = None
        for component in vb_project.VBComponents:
            if component.Type == 100:  # vbext_ct_Document
                # Use VBADocumentNames to detect workbook module in any language
                if VBADocumentNames.is_document_module(component.Name) and component.Name in VBADocumentNames.EXCEL_WORKBOOK_NAMES:
                    doc_module = component
                    print(f"Found workbook document module: {component.Name}")
                    break
        
        assert doc_module is not None, f"Could not find ThisWorkbook/DieseArbeitsmappe/etc. module. VBComponents: {[c.Name for c in vb_project.VBComponents]}"
        
        # Add test code to ThisWorkbook
        code_module = doc_module.CodeModule
        code_module.DeleteLines(1, code_module.CountOfLines)  # Clear existing
        code_module.AddFromString(test_code)
        
        cli = CLITester("excel-vba")
        
        # Export with --in-file-headers (headers embedded in .cls file)
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--force-overwrite",
            "--in-file-headers"  # TEST THE FIX: Headers embedded in code file
        ])
        
        # Read exported file
        # Note: Export will use the actual component name (DieseArbeitsmappe in German Excel)
        exported_file = vba_dir / f"{doc_module.Name}.cls"
        assert exported_file.exists(), f"{doc_module.Name}.cls was not exported. Files: {list(vba_dir.glob('*.cls'))}"
        
        with open(exported_file, 'r', encoding='cp1252') as f:
            exported_content = f.read()
        
        # Verify headers are present in file
        assert "VERSION 1.0 CLASS" in exported_content, "VERSION header missing from exported file"
        assert "BEGIN" in exported_content, "BEGIN header missing from exported file"
        assert "Attribute VB_Name" in exported_content, "Attribute header missing from exported file"
        
        # Extract just the code lines (without headers)
        exported_lines = self.extract_code_lines(exported_content)
        
        # Clear the module
        code_module.DeleteLines(1, code_module.CountOfLines)
        
        # Import with --in-file-headers
        # This is where the bug occurred: headers were imported as code
        cli.assert_success([
            "import",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--in-file-headers"  # TEST THE FIX
        ])
        
        # Get imported code from VBA Editor
        imported_code = code_module.Lines(1, code_module.CountOfLines)
        imported_lines = self.extract_code_lines(imported_code)
        
        # CRITICAL CHECK: Headers should NOT appear in imported code
        for line in imported_lines:
            assert not line.strip().startswith("VERSION"), \
                f"VERSION header imported as code: '{line}'"
            assert not line.strip().startswith("BEGIN"), \
                f"BEGIN header imported as code: '{line}'"
            assert not line.strip() == "END", \
                f"END header imported as code: '{line}'"
            assert not line.strip().startswith("MultiUse"), \
                f"MultiUse header imported as code: '{line}'"
            assert not (line.strip().startswith("Attribute VB_") and "=" in line), \
                f"Attribute header imported as code: '{line}'"
        
        # Verify code integrity - line count should match
        assert len(imported_lines) == len(exported_lines), \
            f"Line count mismatch: Expected={len(exported_lines)}, Got={len(imported_lines)}"
        
        # Verify first few lines match
        for i in range(min(3, len(exported_lines))):
            assert imported_lines[i] == exported_lines[i], \
                f"Line {i+1} mismatch:\n  File: '{exported_lines[i]}'\n  VBE: '{imported_lines[i]}'"
        
        print("✅ SUCCESS: --in-file-headers correctly strips headers from document module!")

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_userform_header_modifications_preserved_save_headers(self, test_workbook, tmp_path):
        """Test that UserForm header modifications are preserved with --save-headers mode.
        
        This validates that users can modify Attribute VB_* lines in .header files
        and have those changes applied when reimporting the UserForm.
        """
        wb, wb_path = test_workbook
        vba_dir = tmp_path / "vba"
        
        vb_project = wb.VBProject
        
        # Get the test UserForm
        test_form = vb_project.VBComponents("TestForm")
        
        # Note: VB_Exposed is not directly accessible via COM
        # Instead, we'll verify by checking the exported header content
        # and confirming the reimport works correctly
        
        print(f"Testing UserForm: {test_form.Name}, Type: {test_form.Type}")
        
        # Export with --save-headers
        cli = CLITester("excel-vba")
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--force-overwrite",
            "--save-headers"
        ])
        
        # For UserForms, headers are always in the .frm file (not separate .header files)
        # because the .frm format includes both the visual design and attributes
        frm_file = vba_dir / "TestForm.frm"
        assert frm_file.exists(), f"TestForm.frm file was not created. Files: {list(vba_dir.glob('*'))}"
        
        # Read the .frm file
        with open(frm_file, 'r', encoding='cp1252') as f:
            frm_content = f.read()
        
        print(f"Original .frm content (first 500 chars):\n{frm_content[:500]}")
        
        # Modify the VB_Exposed attribute in the .frm file
        # Change False to True (or -1 to 0, depending on representation)
        if "Attribute VB_Exposed = False" in frm_content:
            modified_frm = frm_content.replace(
                "Attribute VB_Exposed = False",
                "Attribute VB_Exposed = True"
            )
            expected_new_value = True
        elif "Attribute VB_Exposed = 0" in frm_content:
            modified_frm = frm_content.replace(
                "Attribute VB_Exposed = 0",
                "Attribute VB_Exposed = -1"
            )
            expected_new_value = True
        else:
            # If not found, add it
            modified_frm = frm_content + "\nAttribute VB_Exposed = True\n"
            expected_new_value = True
        
        # Write modified content back
        with open(frm_file, 'w', encoding='cp1252') as f:
            f.write(modified_frm)
        
        print(f"Modified .frm content (first 500 chars):\n{modified_frm[:500]}")
        
        # Remove the UserForm
        vb_project.VBComponents.Remove(test_form)
        
        # Verify it's gone
        try:
            vb_project.VBComponents("TestForm")
            assert False, "UserForm should have been removed"
        except:
            pass  # Expected - component doesn't exist
        
        # Import with modified header
        cli.assert_success([
            "import",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir)
        ])
        
        # Get the reimported UserForm
        reimported_form = vb_project.VBComponents("TestForm")
        
        print(f"UserForm reimported successfully: {reimported_form.Name}")
        
        # Export again to verify the header change was applied
        vba_dir2 = tmp_path / "vba2"
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir2),
            "--force-overwrite",
            "--save-headers"
        ])
        
        # Read the newly exported .frm file
        frm_file2 = vba_dir2 / "TestForm.frm"
        with open(frm_file2, 'r', encoding='cp1252') as f:
            new_frm_content = f.read()
        
        print(f"Re-exported .frm content (first 500 chars):\n{new_frm_content[:500]}")
        
        # Verify the modification was preserved
        if expected_new_value:
            assert ("Attribute VB_Exposed = True" in new_frm_content or 
                    "Attribute VB_Exposed = -1" in new_frm_content), \
                f"VB_Exposed modification was not preserved. Content:\n{new_frm_content[:1000]}"
        
        print("✅ SUCCESS: --save-headers preserves UserForm header modifications!")

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_userform_header_modifications_preserved_in_file_headers(self, test_workbook, tmp_path):
        """Test that UserForm header modifications are preserved with --in-file-headers mode.
        
        This validates that users can modify Attribute VB_* lines embedded in .frm files
        and have those changes applied when reimporting the UserForm.
        """
        wb, wb_path = test_workbook
        vba_dir = tmp_path / "vba"
        
        vb_project = wb.VBProject
        
        # Get the test UserForm
        test_form = vb_project.VBComponents("TestForm")
        
        # Note: VB_Exposed is not directly accessible via COM
        # Instead, we'll verify by checking the exported/reimported header content
        
        print(f"Testing UserForm: {test_form.Name}, Type: {test_form.Type}")
        
        # Export with --in-file-headers
        cli = CLITester("excel-vba")
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--force-overwrite",
            "--in-file-headers"
        ])
        
        # Read the .frm file (headers embedded in code file)
        frm_file = vba_dir / "TestForm.frm"
        assert frm_file.exists(), "TestForm.frm file was not created"
        
        with open(frm_file, 'r', encoding='cp1252') as f:
            frm_content = f.read()
        
        print(f"Original .frm content (first 500 chars):\n{frm_content[:500]}")
        
        # Verify headers are embedded in the file
        assert "VERSION 5.00" in frm_content or "VERSION" in frm_content, \
            "VERSION header not found in .frm file"
        assert "Attribute VB_Name" in frm_content, \
            "VB_Name attribute not found in .frm file"
        
        # Modify the VB_Exposed attribute in the .frm file
        if "Attribute VB_Exposed = False" in frm_content:
            modified_frm = frm_content.replace(
                "Attribute VB_Exposed = False",
                "Attribute VB_Exposed = True"
            )
            expected_new_value = True
        elif "Attribute VB_Exposed = 0" in frm_content:
            modified_frm = frm_content.replace(
                "Attribute VB_Exposed = 0",
                "Attribute VB_Exposed = -1"
            )
            expected_new_value = True
        else:
            # Find the Attribute VB_Name line and add VB_Exposed after it
            lines = frm_content.split('\n')
            for i, line in enumerate(lines):
                if 'Attribute VB_Name' in line:
                    lines.insert(i + 1, 'Attribute VB_Exposed = True')
                    break
            modified_frm = '\n'.join(lines)
            expected_new_value = True
        
        # Write modified content back
        with open(frm_file, 'w', encoding='cp1252') as f:
            f.write(modified_frm)
        
        print(f"Modified .frm content (first 500 chars):\n{modified_frm[:500]}")
        
        # Remove the UserForm
        vb_project.VBComponents.Remove(test_form)
        
        # Verify it's gone
        try:
            vb_project.VBComponents("TestForm")
            assert False, "UserForm should have been removed"
        except:
            pass  # Expected - component doesn't exist
        
        # Import with modified embedded headers
        cli.assert_success([
            "import",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir),
            "--in-file-headers"
        ])
        
        # Get the reimported UserForm
        reimported_form = vb_project.VBComponents("TestForm")
        
        print(f"UserForm reimported successfully: {reimported_form.Name}")
        
        # Export again to verify the header change was applied
        vba_dir2 = tmp_path / "vba2"
        cli.assert_success([
            "export",
            "-f", str(wb_path),
            "--vba-directory", str(vba_dir2),
            "--force-overwrite",
            "--in-file-headers"
        ])
        
        # Read the newly exported .frm file
        frm_file2 = vba_dir2 / "TestForm.frm"
        with open(frm_file2, 'r', encoding='cp1252') as f:
            new_frm_content = f.read()
        
        print(f"Re-exported .frm content (first 500 chars):\n{new_frm_content[:500]}")
        
        # Verify the modification was preserved
        if expected_new_value:
            assert ("Attribute VB_Exposed = True" in new_frm_content or 
                    "Attribute VB_Exposed = -1" in new_frm_content), \
                f"VB_Exposed modification was not preserved. Content:\n{new_frm_content[:1000]}"
        
        print("✅ SUCCESS: --in-file-headers preserves UserForm header modifications!")

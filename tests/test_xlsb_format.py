"""Test Excel Binary Workbook (.xlsb) format support.

This test validates that vba-edit correctly handles xlsb files, which
previously caused Rich markup errors during export/import operations.

Related: Issue #46, Fixed in v0.4.2
"""

import pytest
from pathlib import Path

from tests.cli.helpers import CLITester, get_or_create_app


class TestXLSBFormat:
    """Test that xlsb format works correctly for all operations."""

    @pytest.fixture(scope="class")
    def excel_app(self):
        """Get Excel application instance."""
        import os
        from tests.cli.helpers import check_excel_is_safe_to_use

        # Check environment variable override
        allow_force_quit = os.environ.get("PYTEST_ALLOW_EXCEL_FORCE_QUIT", "0") == "1"

        if not allow_force_quit:
            # Check if it's safe to proceed
            is_safe, message = check_excel_is_safe_to_use()
            if not is_safe:
                pytest.skip(
                    "Excel integration tests skipped - Excel has open workbooks. "
                    "Close Excel and try again, or set PYTEST_ALLOW_EXCEL_FORCE_QUIT=1"
                )

        app = get_or_create_app("excel")
        return app

    @pytest.fixture
    def xlsb_workbook(self, excel_app, tmp_path, request):
        """Create a test .xlsb workbook with VBA code."""
        test_name = request.node.name.replace("[", "_").replace("]", "_")
        wb_path = tmp_path / f"{test_name}.xlsb"

        # Suppress Excel dialogs during tests
        original_display_alerts = excel_app.DisplayAlerts
        excel_app.DisplayAlerts = False

        try:
            # Create workbook
            wb = excel_app.Workbooks.Add()

            # Add VBA components
            vb_project = wb.VBProject

            # Add standard module with test code
            std_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            std_module.Name = "XLSBTestModule"

            test_code = """Option Explicit

' Test module for xlsb format validation
' Contains special characters to test encoding: äöüÄÖÜß €£¥

Public Const TEST_CONSTANT As String = "XLSB Format Test"

Public Function GetMessage() As String
    GetMessage = "Excel Binary Workbook (.xlsb) format is working!"
End Function

Public Sub TestSpecialCharacters()
    Dim text As String
    text = "Größe, Maße, Äpfel, Übung"
    text = text & " - € £ ¥ - "
    text = text & "Quotes: ""double"" and 'single'"
    Debug.Print text
End Sub

Public Function CalculateSum(ParamArray values() As Variant) As Double
    Dim total As Double
    Dim i As Long
    
    total = 0
    For i = LBound(values) To UBound(values)
        total = total + CDbl(values(i))
    Next i
    
    CalculateSum = total
End Function
"""
            std_module.CodeModule.AddFromString(test_code)

            # Add class module
            class_module = vb_project.VBComponents.Add(2)  # 2 = vbext_ct_ClassModule
            class_module.Name = "XLSBTestClass"

            class_code = """Option Explicit

Private m_value As Double

Public Property Get Value() As Double
    Value = m_value
End Property

Public Property Let Value(ByVal newValue As Double)
    m_value = newValue
End Property

Public Function ToString() As String
    ToString = "XLSBTestClass: " & CStr(m_value)
End Function
"""
            class_module.CodeModule.AddFromString(class_code)

            # Save as Excel Binary Workbook (.xlsb)
            # FileFormat=50 is xlExcel12 (Excel Binary Workbook)
            wb.SaveAs(str(wb_path), FileFormat=50)

            yield wb, wb_path

            # Cleanup
            try:
                wb.Saved = True
            except Exception:
                pass
        finally:
            try:
                excel_app.DisplayAlerts = original_display_alerts
            except Exception:
                pass

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_xlsb_export(self, xlsb_workbook, tmp_path):
        """Test that export command works correctly with xlsb files.

        This was the primary failure case in Issue #46:
        Rich markup error when exporting xlsb files.
        """
        wb, wb_path = xlsb_workbook
        vba_dir = tmp_path / "vba_export"

        cli = CLITester("excel-vba")

        # Export should succeed without Rich markup errors
        cli.assert_success(
            [
                "export",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir),
                "--force-overwrite",
                "--save-headers",
            ]
        )

        # Verify files were created
        module_file = vba_dir / "XLSBTestModule.bas"
        class_file = vba_dir / "XLSBTestClass.cls"

        assert module_file.exists(), "XLSBTestModule.bas was not exported"
        assert class_file.exists(), "XLSBTestClass.cls was not exported"

        # Verify content is readable and contains expected code
        with open(module_file, "r", encoding="cp1252") as f:
            module_content = f.read()
        
        assert "XLSB Format Test" in module_content, "Module content missing"
        assert "GetMessage" in module_content, "Function missing from export"
        assert "CalculateSum" in module_content, "Function missing from export"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_xlsb_import(self, xlsb_workbook, tmp_path):
        """Test that import command works correctly with xlsb files.

        Validates that xlsb files can be imported without Rich markup errors.
        """
        wb, wb_path = xlsb_workbook
        vba_dir = tmp_path / "vba_import"

        cli = CLITester("excel-vba")

        # First export
        cli.assert_success(
            [
                "export",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir),
                "--force-overwrite",
                "--save-headers",
                "--keep-open",
            ]
        )

        # Remove the module from VBA project
        vb_project = wb.VBProject
        test_module = vb_project.VBComponents("XLSBTestModule")
        vb_project.VBComponents.Remove(test_module)

        # Import should succeed without Rich markup errors
        cli.assert_success(
            [
                "import",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir),
            ]
        )

        # Verify module was re-imported
        imported_module = vb_project.VBComponents("XLSBTestModule")
        assert imported_module is not None, "Module was not imported"

        # Verify code is present
        code_module = imported_module.CodeModule
        imported_code = code_module.Lines(1, code_module.CountOfLines)
        
        assert "XLSB Format Test" in imported_code, "Imported code missing content"
        assert "GetMessage" in imported_code, "Function missing from import"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_xlsb_roundtrip_preserves_code(self, xlsb_workbook, tmp_path):
        """Test that export → import → export cycle preserves code exactly.

        Validates that xlsb files maintain code integrity through full cycle.
        """
        wb, wb_path = xlsb_workbook
        vba_dir1 = tmp_path / "vba1"
        vba_dir2 = tmp_path / "vba2"

        cli = CLITester("excel-vba")

        # First export
        cli.assert_success(
            [
                "export",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir1),
                "--force-overwrite",
                "--save-headers",
                "--keep-open",
            ]
        )

        # Read first export
        with open(vba_dir1 / "XLSBTestModule.bas", "r", encoding="cp1252") as f:
            export1_content = f.read()

        # Remove module and import
        vb_project = wb.VBProject
        test_module = vb_project.VBComponents("XLSBTestModule")
        vb_project.VBComponents.Remove(test_module)

        cli.assert_success(
            [
                "import",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir1),
            ]
        )

        # Second export
        cli.assert_success(
            [
                "export",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir2),
                "--force-overwrite",
                "--save-headers",
            ]
        )

        # Read second export
        with open(vba_dir2 / "XLSBTestModule.bas", "r", encoding="cp1252") as f:
            export2_content = f.read()

        # Normalize line endings for comparison
        export1_normalized = export1_content.replace("\r\n", "\n").replace("\r", "\n")
        export2_normalized = export2_content.replace("\r\n", "\n").replace("\r", "\n")

        # Content should be identical after roundtrip
        assert export1_normalized == export2_normalized, (
            "Round-trip through xlsb file changed code content"
        )

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_xlsb_with_colorization(self, xlsb_workbook, tmp_path):
        """Test that xlsb export works with colorized output enabled.

        Validates that the Rich library colorization doesn't interfere
        with xlsb processing (the original Issue #46 bug).
        """
        wb, wb_path = xlsb_workbook
        vba_dir = tmp_path / "vba_color"

        cli = CLITester("excel-vba")

        # Export WITH colors (default behavior in TTY)
        # This should NOT cause Rich markup errors
        result = cli.run(
            [
                "export",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir),
                "--force-overwrite",
                "--save-headers",
            ]
        )

        # Should succeed with returncode 0
        assert result.returncode == 0, (
            f"Export with colors failed: {result.stderr}"
        )

        # Verify files were created
        assert (vba_dir / "XLSBTestModule.bas").exists(), "Export failed"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_xlsb_with_no_color_flag(self, xlsb_workbook, tmp_path):
        """Test that xlsb export works with --no-color flag.

        Validates that disabling colors also works correctly.
        """
        wb, wb_path = xlsb_workbook
        vba_dir = tmp_path / "vba_nocolor"

        cli = CLITester("excel-vba")

        # Export WITHOUT colors
        cli.assert_success(
            [
                "export",
                "-f",
                str(wb_path),
                "--vba-directory",
                str(vba_dir),
                "--force-overwrite",
                "--save-headers",
                "--no-color",  # Explicit no-color flag
            ]
        )

        # Verify files were created
        assert (vba_dir / "XLSBTestModule.bas").exists(), "Export with --no-color failed"

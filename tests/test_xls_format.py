"""Test Excel 97-2003 (.xls) format support.

This test validates that vba-edit correctly handles legacy xls files,
which use a different binary format than modern Excel files.

Tests cover basic VBA operations on xls format to ensure backward compatibility.
"""

import pytest
from pathlib import Path

from tests.cli.helpers import CLITester, get_or_create_app


class TestXLSFormat:
    """Test that xls format (Excel 97-2003) works correctly for all operations."""

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
    def xls_workbook(self, excel_app, tmp_path, request):
        """Create a test .xls workbook with VBA code."""
        test_name = request.node.name.replace("[", "_").replace("]", "_")
        wb_path = tmp_path / f"{test_name}.xls"

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
            std_module.Name = "XLSTestModule"

            test_code = """Option Explicit

' Test module for xls format validation (Excel 97-2003)

Public Const TEST_CONSTANT As String = "XLS Format Test"

Public Function GetMessage() As String
    GetMessage = "Excel 97-2003 (.xls) format is working!"
End Function

Public Function AddNumbers(x As Double, y As Double) As Double
    AddNumbers = x + y
End Function
"""
            std_module.CodeModule.AddFromString(test_code)

            # Add class module
            class_module = vb_project.VBComponents.Add(2)  # 2 = vbext_ct_ClassModule
            class_module.Name = "XLSTestClass"

            class_code = """Option Explicit

Private m_name As String

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(ByVal newName As String)
    m_name = newName
End Property
"""
            class_module.CodeModule.AddFromString(class_code)

            # Save as Excel 97-2003 Workbook (.xls)
            # FileFormat=56 is xlExcel8 (Excel 97-2003)
            wb.SaveAs(str(wb_path), FileFormat=56)

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
    def test_xls_export(self, xls_workbook, tmp_path):
        """Test that export command works correctly with xls files."""
        wb, wb_path = xls_workbook
        vba_dir = tmp_path / "vba_export"

        cli = CLITester("excel-vba")

        # Export should succeed
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
        module_file = vba_dir / "XLSTestModule.bas"
        class_file = vba_dir / "XLSTestClass.cls"

        assert module_file.exists(), "XLSTestModule.bas was not exported"
        assert class_file.exists(), "XLSTestClass.cls was not exported"

        # Verify content is readable and contains expected code
        with open(module_file, "r", encoding="cp1252") as f:
            module_content = f.read()

        assert "XLS Format Test" in module_content, "Module content missing"
        assert "GetMessage" in module_content, "Function missing from export"
        assert "AddNumbers" in module_content, "Function missing from export"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_xls_import(self, xls_workbook, tmp_path):
        """Test that import command works correctly with xls files."""
        wb, wb_path = xls_workbook
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
        test_module = vb_project.VBComponents("XLSTestModule")
        vb_project.VBComponents.Remove(test_module)

        # Import should succeed
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
        imported_module = vb_project.VBComponents("XLSTestModule")
        assert imported_module is not None, "Module was not imported"

        # Verify code is present
        code_module = imported_module.CodeModule
        imported_code = code_module.Lines(1, code_module.CountOfLines)

        assert "XLS Format Test" in imported_code, "Imported code missing content"
        assert "GetMessage" in imported_code, "Function missing from import"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    @pytest.mark.excel
    def test_xls_roundtrip_preserves_code(self, xls_workbook, tmp_path):
        """Test that export → import → export cycle preserves code exactly."""
        wb, wb_path = xls_workbook
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
        with open(vba_dir1 / "XLSTestModule.bas", "r", encoding="cp1252") as f:
            export1_content = f.read()

        # Remove module and import
        vb_project = wb.VBProject
        test_module = vb_project.VBComponents("XLSTestModule")
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
        with open(vba_dir2 / "XLSTestModule.bas", "r", encoding="cp1252") as f:
            export2_content = f.read()

        # Normalize line endings for comparison
        export1_normalized = export1_content.replace("\r\n", "\n").replace("\r", "\n")
        export2_normalized = export2_content.replace("\r\n", "\n").replace("\r", "\n")

        # Content should be identical after roundtrip
        assert export1_normalized == export2_normalized, "Round-trip through xls file changed code content"

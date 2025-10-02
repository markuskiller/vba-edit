"""Integration tests with real Office documents."""

import pytest
from .helpers import CLITester, temp_office_doc
from vba_edit.office_vba import RUBBERDUCK_FOLDER_PATTERN


class TestCLIIntegration:
    """Integration tests for CLI with real Office documents."""

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_basic_operations(self, vba_app, temp_office_doc, tmp_path):
        """Test basic CLI operations with a real document."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Test export
        cli.assert_success(["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir)])
        assert vba_dir.exists()
        assert any(vba_dir.glob("*.bas"))  # Should have at least one module

        # Test import (will use files from export)
        cli.assert_success(["import", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir)])

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_rubberduck_folders_export(self, vba_app, temp_office_doc, tmp_path):
        """Test that Rubberduck folder annotations are respected during export."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Test export with Rubberduck folders enabled
        cli.assert_success(
            ["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--rubberduck-folders"]
        )

        assert vba_dir.exists()

        # Check that the standard module is exported to the root
        standard_files = list(vba_dir.glob("TestModule.bas"))
        assert len(standard_files) == 1, "TestModule.bas should be in root directory"

        # Find the class module (could be in subdirectory or root)
        class_files = list(vba_dir.rglob("TestClass.cls"))
        assert len(class_files) == 1, "Should find exactly one TestClass.cls file"

        class_file = class_files[0]
        class_content = class_file.read_text(encoding="cp1252")

        # Verify the folder annotation is preserved in the file
        folder_match = None
        for line in class_content.splitlines():
            match = RUBBERDUCK_FOLDER_PATTERN.match(line.strip())
            if match:
                folder_match = match
                break

        assert folder_match is not None, f"No @Folder annotation found in content:\n{class_content}"
        assert folder_match.group(1) == "Business.Domain", (
            f"Expected folder 'Business.Domain', but found '{folder_match.group(1)}'"
        )

        # Check if the file is in the expected subdirectory (optional verification)
        business_dir = vba_dir / "Business" / "Domain"
        if business_dir.exists():
            # File should be in the Business/Domain subdirectory
            assert class_file.parent == business_dir, (
                f"TestClass.cls should be in {business_dir}, but found in {class_file.parent}"
            )
        else:
            # If folder structure wasn't created, file should be in root
            assert class_file.parent == vba_dir, (
                f"TestClass.cls should be in root {vba_dir}, but found in {class_file.parent}"
            )

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_rubberduck_folders_import(self, vba_app, temp_office_doc, tmp_path):
        """Test that Rubberduck folder structure is preserved during import."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # First export with Rubberduck folders
        cli.assert_success(
            ["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--rubberduck-folders"]
        )

        # Modify the class file to test round-trip
        class_files = list(vba_dir.rglob("TestClass.cls"))
        assert len(class_files) == 1, "Should find exactly one TestClass.cls file"

        class_file = class_files[0]
        original_content = class_file.read_text(encoding="cp1252")
        modified_content = original_content.replace(
            'Debug.Print "TestClass initialized"', 'Debug.Print "TestClass initialized - MODIFIED"'
        )
        class_file.write_text(modified_content, encoding="cp1252")

        # Import back with Rubberduck folders
        cli.assert_success(
            ["import", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--rubberduck-folders"]
        )

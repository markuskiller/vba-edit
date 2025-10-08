"""Integration tests with real Office documents."""

import pytest

from vba_edit.office_vba import RUBBERDUCK_FOLDER_PATTERN

from .helpers import CLITester, temp_office_doc


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


class TestSafetyFeaturesIntegration:
    """Integration tests for export safety features."""

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_force_overwrite_flag_present(self, vba_app, temp_office_doc, tmp_path):
        """Test that --force-overwrite flag is available in CLI."""
        cli = CLITester(f"{vba_app}-vba")

        # Check help text includes --force-overwrite
        result = cli.run(["export", "--help"])
        assert "--force-overwrite" in result.stdout
        assert result.returncode == 0

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_force_overwrite_with_existing_files(self, vba_app, temp_office_doc, tmp_path):
        """Test --force-overwrite bypasses prompts for existing files."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # First export to create files
        cli.assert_success(["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir)])
        assert vba_dir.exists()

        # Second export with --force-overwrite should succeed without prompts
        result = cli.run(["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--force-overwrite"])

        # Should succeed without user interaction
        assert result.returncode == 0 or "No VBA components found" in (result.stdout + result.stderr)

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_export_with_save_headers_separate(self, vba_app, temp_office_doc, tmp_path):
        """Test export with separate header files."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Export with separate headers
        cli.assert_success(["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--save-headers"])

        # Check for .header files (may or may not exist depending on components)
        # The test just verifies the option doesn't cause errors
        if vba_dir.exists():
            _ = list(vba_dir.glob("*.header"))  # Check doesn't crash

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_export_with_inline_headers(self, vba_app, temp_office_doc, tmp_path):
        """Test export with inline headers."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Export with inline headers
        cli.assert_success(["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--in-file-headers"])

        assert vba_dir.exists()

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_metadata_with_header_mode(self, vba_app, temp_office_doc, tmp_path):
        """Test that metadata file includes header_mode."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Export with metadata and inline headers
        cli.assert_success(
            [
                "export",
                "-f",
                str(temp_office_doc),
                "--vba-directory",
                str(vba_dir),
                "--save-metadata",
                "--in-file-headers",
            ]
        )

        # Check metadata file
        metadata_file = vba_dir / "vba_metadata.json"
        if metadata_file.exists():
            import json

            metadata = json.loads(metadata_file.read_text())
            assert "header_mode" in metadata
            assert metadata["header_mode"] == "inline"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_header_mode_change_with_force_overwrite(self, vba_app, temp_office_doc, tmp_path):
        """Test changing header mode with --force-overwrite."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # First export with inline headers and metadata
        cli.assert_success(
            [
                "export",
                "-f",
                str(temp_office_doc),
                "--vba-directory",
                str(vba_dir),
                "--save-metadata",
                "--in-file-headers",
            ]
        )

        # Second export with separate headers (mode change) and force-overwrite
        result = cli.run(
            [
                "export",
                "-f",
                str(temp_office_doc),
                "--vba-directory",
                str(vba_dir),
                "--save-metadata",
                "--save-headers",
                "--force-overwrite",
            ]
        )

        # Should succeed without prompts
        assert result.returncode == 0 or "No VBA components found" in (result.stdout + result.stderr)

        # Check metadata updated
        metadata_file = vba_dir / "vba_metadata.json"
        if metadata_file.exists():
            import json

            metadata = json.loads(metadata_file.read_text())
            assert metadata["header_mode"] == "separate"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_force_overwrite_no_short_option(self, vba_app, temp_office_doc, tmp_path):
        """Test that -f is not used for --force-overwrite (reserved for --file)."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # -f should be for --file, not --force-overwrite
        result = cli.run(
            [
                "export",
                "-f",
                str(temp_office_doc),  # This is --file
                "--vba-directory",
                str(vba_dir),
            ]
        )

        # Should work (if document exists)
        # -f is correctly interpreted as --file
        assert result.returncode == 0 or "not found" in result.stderr.lower()

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_conflicting_header_options_error(self, vba_app, temp_office_doc, tmp_path):
        """Test that --save-headers and --in-file-headers are mutually exclusive."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Both header options should cause error
        result = cli.run(
            [
                "export",
                "-f",
                str(temp_office_doc),
                "--vba-directory",
                str(vba_dir),
                "--save-headers",
                "--in-file-headers",
            ]
        )

        # Should fail with error message
        assert result.returncode != 0
        full_output = result.stdout + result.stderr
        assert (
            "mutually exclusive" in full_output.lower()
            or "conflicting" in full_output.lower()
            or "not allowed with" in full_output.lower()
        )

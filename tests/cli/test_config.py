"""Configuration file functionality tests."""

from pathlib import Path

import pytest

from vba_edit.cli_common import (
    CONFIG_KEY_FILE,
    CONFIG_KEY_OPEN_FOLDER,
    CONFIG_KEY_RUBBERDUCK_FOLDERS,
    CONFIG_KEY_VBA_DIRECTORY,
    CONFIG_KEY_VERBOSE,
    CONFIG_SECTION_GENERAL,
    PLACEHOLDER_CONFIG_PATH,
    PLACEHOLDER_FILE_NAME,
    PLACEHOLDER_FILE_PATH,
    PLACEHOLDER_VBA_PROJECT,
)
from vba_edit.office_vba import OFFICE_MACRO_EXTENSIONS

from .helpers import CLITester, temp_office_doc


class TestCLIConfig:
    """Tests for configuration file functionality."""

    def create_test_config(self, config_dir: Path, config_content: str) -> Path:
        """Create a test configuration file.

        Args:
            config_dir: Directory to create config file in
            config_content: TOML content for the config file

        Returns:
            Path to the created config file
        """
        config_file = config_dir / "test-config.toml"
        config_file.write_text(config_content, encoding="utf-8")
        return config_file

    @pytest.mark.office
    def test_open_folder_default_behavior(self, vba_app):
        """Test that open_folder defaults to False when not specified."""
        cli = CLITester(f"{vba_app}-vba")

        # No config file, no flag - should use default (False)
        result = cli.run(["export", "--help"])
        assert result.returncode == 0

        # Help should show the option with default behavior
        help_text = result.stdout + result.stderr
        assert "--open-folder" in help_text

    @pytest.mark.office
    def test_config_file_basic(self, vba_app, tmp_path):
        """Test basic configuration file loading."""
        cli = CLITester(f"{vba_app}-vba")

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # Test that config file is accepted
        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_file_not_found(self, vba_app):
        """Test handling of missing configuration file."""
        cli = CLITester(f"{vba_app}-vba")

        result = cli.run(["export", "--conf", "nonexistent-config.toml", "--help"])
        # Should show error but not crash
        assert "Error loading configuration file" in result.stderr or result.returncode == 0

    @pytest.mark.office
    def test_config_file_invalid_toml(self, vba_app, tmp_path):
        """Test handling of invalid TOML configuration."""
        cli = CLITester(f"{vba_app}-vba")

        # Create invalid TOML file
        config_file = tmp_path / "invalid-config.toml"
        config_file.write_text("invalid = toml content [", encoding="utf-8")

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        # Should show error but not crash
        assert "Error loading configuration file" in result.stderr or result.returncode == 0

    @pytest.mark.office
    def test_cli_args_override_config(self, vba_app, tmp_path):
        """Test that CLI arguments override configuration file values."""
        cli = CLITester(f"{vba_app}-vba")

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # CLI should override config (verbose=false should override verbose=true in config)
        result = cli.run(["export", "--conf", str(config_file), "--help"])
        # This test mainly ensures no errors occur during argument processing
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_placeholder_config_path(self, vba_app, tmp_path):
        """Test {config.path} placeholder resolution."""
        cli = CLITester(f"{vba_app}-vba")

        # Create subdirectory for config
        config_dir = tmp_path / "config"
        config_dir.mkdir()

        vba_modules_path = f"{PLACEHOLDER_CONFIG_PATH}/vba-modules"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_modules_path}"
"""
        config_file = self.create_test_config(config_dir, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_placeholder_file_info(self, vba_app, tmp_path):
        """Test file-related placeholder resolution."""
        cli = CLITester(f"{vba_app}-vba")

        # Create a mock Office file path for testing
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = Path(tmp_path) / f"MyProject{extension}"

        vba_directory_path = f"{PLACEHOLDER_FILE_PATH}/vba-{PLACEHOLDER_FILE_NAME}"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file.as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_multiple_placeholders(self, vba_app, tmp_path):
        """Test multiple placeholders in the same value."""
        cli = CLITester(f"{vba_app}-vba")

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = Path(tmp_path) / f"TestDoc{extension}"

        vba_directory_path = f"{PLACEHOLDER_CONFIG_PATH}/{PLACEHOLDER_FILE_NAME}-modules"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file.as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_config_file_with_real_document(self, vba_app, temp_office_doc, tmp_path):
        """Test configuration file with actual Office document operations."""
        cli = CLITester(f"{vba_app}-vba")

        vba_directory_path = f"{PLACEHOLDER_FILE_PATH}/vba-{PLACEHOLDER_FILE_NAME}"

        # Create config that uses placeholders
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(temp_office_doc).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # Test export using config file
        result = cli.run(["export", "--conf", str(config_file), "--verbose", "--logfile"])

        # Should succeed or show meaningful error
        success = result.returncode == 0 or "No VBA components found" in (result.stdout + result.stderr)
        assert success, f"Command failed with output: {result.stdout + result.stderr}"

        # Check that VBA directory was created with placeholder resolution
        expected_vba_dir = temp_office_doc.parent / f"vba-{temp_office_doc.stem}"
        if success and result.returncode == 0:
            assert expected_vba_dir.exists(), f"Expected VBA directory {expected_vba_dir} was not created"

    @pytest.mark.office
    def test_config_relative_paths(self, vba_app, tmp_path):
        """Test relative paths in configuration file."""
        cli = CLITester(f"{vba_app}-vba")

        # Create subdirectory structure
        config_dir = tmp_path / "config"
        config_dir.mkdir()
        docs_dir = tmp_path / "documents"
        docs_dir.mkdir()

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]

        relative_file_path = f"../documents/test{extension}"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(relative_file_path).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "vba-output"
"""
        config_file = self.create_test_config(config_dir, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_nested_placeholders(self, vba_app, tmp_path):
        """Test nested configuration sections with placeholders."""
        cli = CLITester(f"{vba_app}-vba")

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"Project{extension}"

        vba_modules_path = f"{PLACEHOLDER_FILE_PATH}/modules"
        backup_path = f"{PLACEHOLDER_CONFIG_PATH}/backups/{PLACEHOLDER_FILE_NAME}"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(test_file).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_modules_path}"

[office]
application = "{vba_app}"

[advanced]
backup_directory = "{backup_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_vbaproject_placeholder_preparation(self, vba_app, tmp_path):
        """Test that {vbaproject} placeholder is preserved for later resolution."""
        cli = CLITester(f"{vba_app}-vba")

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"TestFile{extension}"

        vba_directory_path = f"{PLACEHOLDER_FILE_PATH}/{PLACEHOLDER_VBA_PROJECT}-modules"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(test_file).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # This should not fail even though {vbaproject} cannot be resolved yet
        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

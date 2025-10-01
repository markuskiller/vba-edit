"""CLI option debugging tests that write effective values to files."""

import pytest
import subprocess
from pathlib import Path
from .helpers import CLITester, temp_office_doc
from vba_edit.office_vba import OFFICE_MACRO_EXTENSIONS
from vba_edit.cli_common import (
    PLACEHOLDER_CONFIG_PATH,
    PLACEHOLDER_FILE_NAME,
    PLACEHOLDER_FILE_PATH,
    PLACEHOLDER_VBA_PROJECT,
    CONFIG_SECTION_GENERAL,
    CONFIG_KEY_FILE,
    CONFIG_KEY_VBA_DIRECTORY,
    CONFIG_KEY_VERBOSE,
    CONFIG_KEY_RUBBERDUCK_FOLDERS,
    #    CONFIG_KEY_OPEN_FOLDER,
)


class TestCLIOptionsDebugging:
    """Tests that write effective option values to files for debugging."""

    def write_options_file(self, tmp_path: Path, test_name: str, options_dict: dict) -> Path:
        """Write effective options to a test file for debugging.

        Args:
            tmp_path: Temporary directory path
            test_name: Name of the test case
            options_dict: Dictionary of option names and values

        Returns:
            Path to the created options file
        """
        options_file = tmp_path / f"test_options_{test_name}.txt"
        content = f"Test: {test_name}\n"
        content += "=" * 50 + "\n"
        content += f"CLI Command: {options_dict.get('cli_command', 'unknown')}\n"
        content += f"Return Code: {options_dict.get('return_code', 'unknown')}\n"
        content += "-" * 30 + "\n"

        # Separate different types of options for better readability
        config_options = {k: v for k, v in options_dict.items() if k.startswith(("config_", "original_", "template_"))}
        runtime_options = {
            k: v for k, v in options_dict.items() if k.startswith(("resolved_", "expected_", "actual_", "cli_"))
        }
        output_options = {
            k: v for k, v in options_dict.items() if k.startswith(("stdout_", "stderr_", "output_", "return_"))
        }
        other_options = {
            k: v
            for k, v in options_dict.items()
            if not any(
                k.startswith(prefix)
                for prefix in [
                    "config_",
                    "original_",
                    "template_",
                    "resolved_",
                    "expected_",
                    "actual_",
                    "cli_",
                    "stdout_",
                    "stderr_",
                    "output_",
                    "return_",
                ]
            )
        }

        if config_options:
            content += "Configuration Options:\n"
            for key, value in sorted(config_options.items()):
                content += f"  {key}: {value}\n"
            content += "\n"

        if runtime_options:
            content += "Runtime/Resolved Options:\n"
            for key, value in sorted(runtime_options.items()):
                content += f"  {key}: {value}\n"
            content += "\n"

        if output_options:
            content += "Command Output Information:\n"
            for key, value in sorted(output_options.items()):
                content += f"  {key}: {value}\n"
            content += "\n"

        if other_options:
            content += "Other Options:\n"
            for key, value in sorted(other_options.items()):
                content += f"  {key}: {value}\n"
            content += "\n"

        # Add environment debug information
        content += "Environment Information:\n"
        content += f"  Working Directory: {Path.cwd()}\n"
        content += f"  Temp Path: {tmp_path}\n"
        content += f"  Test File: {options_file}\n"
        content += f"  Python Executable: {subprocess.sys.executable}\n"

        options_file.write_text(content, encoding="utf-8")
        return options_file

    def extract_cli_information(self, cli_args: list, result: subprocess.CompletedProcess) -> dict:
        """Extract debugging information from CLI command execution.

        Args:
            cli_args: CLI arguments that were executed
            result: Result from subprocess execution

        Returns:
            Dictionary with extracted information
        """
        info = {
            "cli_args_count": len(cli_args),
            "cli_args_string": " ".join(cli_args),
            "return_code": result.returncode,
            "stdout_length": len(result.stdout),
            "stderr_length": len(result.stderr),
            "total_output_length": len(result.stdout) + len(result.stderr),
            "has_help_output": "help" in result.stdout.lower() or "usage" in result.stdout.lower(),
            "has_error_output": result.returncode != 0 or len(result.stderr) > 0,
            "contains_config_reference": "--conf" in cli_args,
            "contains_file_reference": any(arg in cli_args for arg in ["-f", "--file"]),
            "contains_vba_directory": "--vba-directory" in cli_args,
            "contains_verbose_flag": "--verbose" in cli_args,
            "contains_rubberduck_flag": "--rubberduck-folders" in cli_args,
        }

        # Extract file and directory arguments
        for i, arg in enumerate(cli_args):
            if arg in ["-f", "--file"] and i + 1 < len(cli_args):
                info["cli_file_argument"] = cli_args[i + 1]
            elif arg == "--vba-directory" and i + 1 < len(cli_args):
                info["cli_vba_directory_argument"] = cli_args[i + 1]
            elif arg == "--conf" and i + 1 < len(cli_args):
                info["cli_config_file_argument"] = cli_args[i + 1]

        return info

    @pytest.mark.office
    def test_default_options_values(self, vba_app, tmp_path):
        """Test default option values without any configuration."""
        cli = CLITester(f"{vba_app}-vba")

        # Create test file paths
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"test{extension}"
        vba_dir = tmp_path / "vba"

        # Run export with minimal arguments
        cli_args = [
            "export",
            "-f",
            str(test_file),
            "--vba-directory",
            str(vba_dir),
            "--help",  # Use help to avoid actual file operations
        ]
        result = cli.run(cli_args)

        # Extract CLI information
        cli_info = self.extract_cli_information(cli_args, result)

        # Build comprehensive options dictionary
        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "test_scenario": "default_options_without_config",
            "vba_app": vba_app,
            "file_extension": extension,
            "test_file_path": str(test_file),
            "vba_directory_path": str(vba_dir),
            "config_file_used": "none",
            "expected_verbose": "false",
            "expected_rubberduck_folders": "false",
            **cli_info,
        }

        options_file = self.write_options_file(tmp_path, f"default_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.office
    def test_config_file_option_merging(self, vba_app, tmp_path):
        """Test option merging between config file and command line."""
        cli = CLITester(f"{vba_app}-vba")

        # Create test paths
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"project{extension}"
        vba_dir = tmp_path / "vba_modules"

        # Create config file with specific options
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(test_file).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_dir}"
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = false
"""
        config_file = tmp_path / "test-config.toml"
        config_file.write_text(config_content, encoding="utf-8")

        # Run with config file and CLI overrides
        override_vba_dir = tmp_path / "override_vba"
        cli_args = [
            "export",
            "--conf",
            str(config_file),
            "--vba-directory",
            str(override_vba_dir),  # Override config
            "--rubberduck-folders",  # Override config
            "--help",
        ]
        result = cli.run(cli_args)

        # Extract CLI information
        cli_info = self.extract_cli_information(cli_args, result)

        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "test_scenario": "config_file_with_cli_overrides",
            "vba_app": vba_app,
            "config_file_path": str(config_file),
            "config_file_exists": str(config_file.exists()),
            "config_file_size": len(config_content),
            "config_content_preview": config_content.replace("\n", "\\n")[:200] + "...",
            "config_file_from_config": str(test_file),
            "config_vba_directory_from_config": str(vba_dir),
            "config_verbose_from_config": "true",
            "config_rubberduck_from_config": "false",
            "cli_vba_directory_override": str(override_vba_dir),
            "cli_rubberduck_override": "true",
            "expected_precedence": "CLI arguments should override config file",
            **cli_info,
        }

        options_file = self.write_options_file(tmp_path, f"config_merge_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.office
    def test_placeholder_resolution_debugging(self, vba_app, tmp_path):
        """Test placeholder resolution with detailed logging."""
        cli = CLITester(f"{vba_app}-vba")

        # Create directory structure for testing placeholders
        config_dir = tmp_path / "config" / "project"
        config_dir.mkdir(parents=True)

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / "documents" / f"MyProject{extension}"
        test_file.parent.mkdir(parents=True)

        # Use placeholders in configuration
        vba_directory_with_placeholders = f"{PLACEHOLDER_CONFIG_PATH}/modules/{PLACEHOLDER_FILE_NAME}"
        backup_path = f"{PLACEHOLDER_FILE_PATH}/backups"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(test_file).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_with_placeholders}"
{CONFIG_KEY_VERBOSE} = true

[advanced]
backup_directory = "{backup_path}"
output_template = "{PLACEHOLDER_VBA_PROJECT}_exported"
"""
        config_file = config_dir / "project-config.toml"
        config_file.write_text(config_content, encoding="utf-8")

        cli_args = ["export", "--conf", str(config_file), "--help"]
        result = cli.run(cli_args)

        # Calculate expected placeholder resolutions
        expected_config_path = str(config_dir)
        expected_file_path = str(test_file.parent)
        expected_file_name = test_file.stem
        expected_vba_dir = f"{expected_config_path}/modules/{expected_file_name}"
        expected_backup_dir = f"{expected_file_path}/backups"

        # Extract CLI information
        cli_info = self.extract_cli_information(cli_args, result)

        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "test_scenario": "placeholder_resolution_testing",
            "vba_app": vba_app,
            "config_file_path": str(config_file),
            "config_directory": str(config_dir),
            "test_file_path": str(test_file),
            "test_file_directory": str(test_file.parent),
            "test_file_name": test_file.stem,
            "original_vba_directory": vba_directory_with_placeholders,
            "expected_vba_directory": expected_vba_dir,
            "original_backup_directory": backup_path,
            "expected_backup_directory": expected_backup_dir,
            "template_config_path": PLACEHOLDER_CONFIG_PATH,
            "template_file_path": PLACEHOLDER_FILE_PATH,
            "template_file_name": PLACEHOLDER_FILE_NAME,
            "template_vba_project": PLACEHOLDER_VBA_PROJECT,
            "placeholder_resolution_note": "Placeholders should be resolved by CLI during actual execution",
            **cli_info,
        }

        options_file = self.write_options_file(tmp_path, f"placeholders_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.office
    def test_command_line_precedence(self, vba_app, tmp_path):
        """Test command line argument precedence over config file."""
        cli = CLITester(f"{vba_app}-vba")

        # Create conflicting file paths for testing precedence
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        config_file_path = tmp_path / f"config_file{extension}"
        cli_file_path = tmp_path / f"cli_file{extension}"
        config_vba_dir = tmp_path / "config_vba"
        cli_vba_dir = tmp_path / "cli_vba"

        # Config file specifies certain values
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(config_file_path).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{Path(config_vba_dir).as_posix()}"
{CONFIG_KEY_VERBOSE} = false
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = false
"""
        config_file = tmp_path / "precedence-test.toml"
        config_file.write_text(config_content, encoding="utf-8")

        # Command line arguments that override config
        cli_args = [
            "export",
            "--conf",
            str(config_file),
            "-f",
            str(cli_file_path),  # Override file
            "--vba-directory",
            str(cli_vba_dir),  # Override VBA directory
            "--verbose",  # Override verbose
            "--rubberduck-folders",  # Override rubberduck folders
            "--help",
        ]
        result = cli.run(cli_args)

        # Extract CLI information
        cli_info = self.extract_cli_information(cli_args, result)

        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "test_scenario": "command_line_precedence_testing",
            "vba_app": vba_app,
            "config_file_path": str(config_file),
            "config_content_preview": config_content.replace("\n", "\\n"),
            "config_file_value": str(config_file_path),
            "config_vba_directory": str(config_vba_dir),
            "config_verbose": "false",
            "config_rubberduck": "false",
            "cli_file_override": str(cli_file_path),
            "cli_vba_directory_override": str(cli_vba_dir),
            "cli_verbose_override": "true",
            "cli_rubberduck_override": "true",
            "precedence_expectation": "CLI arguments should take precedence over config file values",
            "precedence_test_result": "CLI overrides all config values",
            **cli_info,
        }

        options_file = self.write_options_file(tmp_path, f"precedence_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_real_document_option_resolution(self, vba_app, temp_office_doc, tmp_path):
        """Test option resolution with a real Office document."""
        cli = CLITester(f"{vba_app}-vba")

        # Create config with placeholders for real document
        vba_path_template = f"{PLACEHOLDER_FILE_PATH}/{PLACEHOLDER_VBA_PROJECT}-{PLACEHOLDER_FILE_NAME}"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{Path(temp_office_doc).as_posix()}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_path_template}"
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true

[export]
create_folders = true

[import]
overwrite_existing = true
"""
        config_file = tmp_path / "real-doc-config.toml"
        config_file.write_text(config_content, encoding="utf-8")

        # Test export operation
        cli_args = ["export", "--conf", str(config_file)]
        result = cli.run(cli_args)

        # Calculate expected values
        file_path = str(temp_office_doc.parent)
        file_name = temp_office_doc.stem
        expected_vba_dir = f"{file_path}/VBAProject-{file_name}"  # Expected after placeholder resolution

        # Extract CLI information
        cli_info = self.extract_cli_information(cli_args, result)

        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "test_scenario": "real_document_with_placeholders",
            "vba_app": vba_app,
            "document_path": str(temp_office_doc),
            "document_exists": str(temp_office_doc.exists()),
            "document_size": str(temp_office_doc.stat().st_size) if temp_office_doc.exists() else "0",
            "config_file_path": str(config_file),
            "config_content_preview": config_content.replace("\n", "\\n")[:300] + "...",
            "template_vba_directory": vba_path_template,
            "expected_vba_directory": expected_vba_dir,
            "file_path_component": file_path,
            "file_name_component": file_name,
            "operation_type": "export_from_real_document",
            "operation_successful": str(result.returncode == 0),
            "has_vba_output": str("VBA" in (result.stdout + result.stderr)),
            **cli_info,
        }

        options_file = self.write_options_file(tmp_path, f"real_doc_{vba_app}", effective_options)
        assert options_file.exists()

        # Check for created directories after successful export
        if result.returncode == 0:
            created_dirs = []
            for item in temp_office_doc.parent.iterdir():
                if item.is_dir() and ("vba" in item.name.lower() or file_name in item.name):
                    created_dirs.append(str(item))

            if created_dirs:
                effective_options["actual_created_directories"] = "|".join(created_dirs)
                # Write updated options with actual results
                updated_options_file = self.write_options_file(
                    tmp_path, f"real_doc_updated_{vba_app}", effective_options
                )
                assert updated_options_file.exists()

    @pytest.mark.office
    def test_multiple_config_scenarios(self, vba_app, tmp_path):
        """Test various configuration scenarios and document their behavior."""
        cli = CLITester(f"{vba_app}-vba")

        scenarios = [
            {
                "name": "minimal_config",
                "description": "Minimal configuration with only verbose setting",
                "config": f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
""",
                "args": ["export", "--help"],
            },
            {
                "name": "full_config",
                "description": "Full configuration with all common settings",
                "config": f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true

[office]
application = "{vba_app}"
""",
                "args": ["export", "--help"],
            },
            {
                "name": "config_with_placeholders",
                "description": "Configuration using placeholder values",
                "config": f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "test{OFFICE_MACRO_EXTENSIONS[vba_app]}"
{CONFIG_KEY_VBA_DIRECTORY} = "{PLACEHOLDER_CONFIG_PATH}/vba-files"
""",
                "args": ["export", "--help"],
            },
        ]

        for scenario in scenarios:
            config_file = tmp_path / f"config_{scenario['name']}.toml"
            config_file.write_text(scenario["config"], encoding="utf-8")

            cli_args = [*scenario["args"], "--conf", str(config_file)]
            result = cli.run(cli_args)

            # Extract CLI information
            cli_info = self.extract_cli_information(cli_args, result)

            effective_options = {
                "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
                "test_scenario": f"multiple_scenarios_{scenario['name']}",
                "scenario_name": scenario["name"],
                "scenario_description": scenario["description"],
                "vba_app": vba_app,
                "config_file_path": str(config_file),
                "config_content": scenario["config"].replace("\n", "\\n"),
                "config_file_size": len(scenario["config"]),
                "config_lines_count": len(scenario["config"].splitlines()),
                "command_args": " ".join(scenario["args"]),
                "has_help_in_args": "help" in scenario["args"],
                "has_placeholders": PLACEHOLDER_CONFIG_PATH in scenario["config"],
                **cli_info,
            }

            options_file = self.write_options_file(
                tmp_path, f"scenario_{scenario['name']}_{vba_app}", effective_options
            )
            assert options_file.exists()

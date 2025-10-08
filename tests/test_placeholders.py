"""Tests for simplified placeholder format (v0.4.1+)."""

import pytest

from vba_edit.cli_common import (
    PLACEHOLDER_CONFIG_PATH,
    PLACEHOLDER_FILE_FULLNAME,
    PLACEHOLDER_FILE_FULLNAME_LEGACY,
    PLACEHOLDER_FILE_NAME,
    PLACEHOLDER_FILE_NAME_LEGACY,
    PLACEHOLDER_FILE_PATH,
    PLACEHOLDER_FILE_PATH_LEGACY,
    PLACEHOLDER_FILE_VBAPROJECT,
    PLACEHOLDER_VBA_PROJECT_LEGACY,
    get_placeholder_values,
    resolve_placeholders_in_value,
)


class TestSimplifiedPlaceholders:
    """Tests for new simplified placeholder format."""

    def test_new_placeholder_constants(self):
        """Test that new simplified placeholder constants are defined correctly."""
        assert PLACEHOLDER_FILE_NAME == "{file.name}"
        assert PLACEHOLDER_FILE_FULLNAME == "{file.fullname}"
        assert PLACEHOLDER_FILE_PATH == "{file.path}"
        assert PLACEHOLDER_FILE_VBAPROJECT == "{file.vbaproject}"
        assert PLACEHOLDER_CONFIG_PATH == "{config.path}"

    def test_legacy_placeholder_constants(self):
        """Test that legacy placeholder constants are still available."""
        assert PLACEHOLDER_FILE_NAME_LEGACY == "{general.file.name}"
        assert PLACEHOLDER_FILE_FULLNAME_LEGACY == "{general.file.fullname}"
        assert PLACEHOLDER_FILE_PATH_LEGACY == "{general.file.path}"
        assert PLACEHOLDER_VBA_PROJECT_LEGACY == "{vbaproject}"

    def test_new_placeholder_resolution(self):
        """Test that new simplified placeholders are resolved correctly."""
        placeholders = get_placeholder_values(
            config_file_path="C:/Projects/config.toml", file_path="C:/Projects/docs/MyDocument.docx"
        )

        # New format should be resolved
        assert placeholders[PLACEHOLDER_FILE_NAME] == "MyDocument"
        assert placeholders[PLACEHOLDER_FILE_FULLNAME] == "MyDocument.docx"
        assert placeholders[PLACEHOLDER_FILE_PATH] == "C:\\Projects\\docs"
        assert placeholders[PLACEHOLDER_CONFIG_PATH] == "C:\\Projects"

    def test_legacy_placeholder_resolution(self):
        """Test that legacy placeholders are still resolved for backward compatibility."""
        placeholders = get_placeholder_values(
            config_file_path="C:/Projects/config.toml", file_path="C:/Projects/docs/MyDocument.docx"
        )

        # Legacy format should also be resolved with same values
        assert placeholders[PLACEHOLDER_FILE_NAME_LEGACY] == "MyDocument"
        assert placeholders[PLACEHOLDER_FILE_FULLNAME_LEGACY] == "MyDocument.docx"
        assert placeholders[PLACEHOLDER_FILE_PATH_LEGACY] == "C:\\Projects\\docs"

    def test_new_placeholder_in_string_replacement(self):
        """Test replacing new simplified placeholders in strings."""
        placeholders = {
            PLACEHOLDER_FILE_NAME: "MyDoc",
            PLACEHOLDER_FILE_PATH: "C:/Projects",
        }

        test_string = "{file.path}/{file.name}-vba"
        result = resolve_placeholders_in_value(test_string, placeholders)

        assert result == "C:/Projects/MyDoc-vba"

    def test_legacy_placeholder_in_string_replacement(self):
        """Test replacing legacy placeholders in strings (backward compatibility)."""
        placeholders = {
            PLACEHOLDER_FILE_NAME_LEGACY: "MyDoc",
            PLACEHOLDER_FILE_PATH_LEGACY: "C:/Projects",
        }

        test_string = "{general.file.path}/{general.file.name}-vba"
        result = resolve_placeholders_in_value(test_string, placeholders)

        assert result == "C:/Projects/MyDoc-vba"

    def test_mixed_placeholder_formats(self):
        """Test that both old and new formats can coexist."""
        placeholders = get_placeholder_values(config_file_path="C:/config.toml", file_path="C:/docs/file.xlsx")

        # Both formats should resolve to same values
        new_format = "{file.name}"
        old_format = "{general.file.name}"

        new_result = resolve_placeholders_in_value(new_format, placeholders)
        old_result = resolve_placeholders_in_value(old_format, placeholders)

        assert new_result == old_result == "file"

    def test_vbaproject_placeholder_new_format(self):
        """Test new vbaproject placeholder format."""
        placeholders = {PLACEHOLDER_FILE_VBAPROJECT: "MyProject"}

        test_string = "Project: {file.vbaproject}"
        result = resolve_placeholders_in_value(test_string, placeholders)

        assert result == "Project: MyProject"

    def test_vbaproject_placeholder_legacy_format(self):
        """Test legacy vbaproject placeholder format."""
        placeholders = {PLACEHOLDER_VBA_PROJECT_LEGACY: "MyProject"}

        test_string = "Project: {vbaproject}"
        result = resolve_placeholders_in_value(test_string, placeholders)

        assert result == "Project: MyProject"

    def test_complex_path_with_new_placeholders(self):
        """Test complex path construction with new placeholders."""
        placeholders = get_placeholder_values(
            config_file_path="C:/Work/myproject/config.toml", file_path="C:/Work/myproject/data/spreadsheet.xlsm"
        )

        # Test various path combinations (note: forward slashes in templates are preserved)
        test_cases = [
            ("{file.path}/{file.name}_backup", "C:\\Work\\myproject\\data/spreadsheet_backup"),
            ("{config.path}/exports/{file.name}", "C:\\Work\\myproject/exports/spreadsheet"),
            ("{file.name}.{file.fullname}", "spreadsheet.spreadsheet.xlsm"),
        ]

        for template, expected in test_cases:
            result = resolve_placeholders_in_value(template, placeholders)
            assert result == expected, f"Failed for template: {template}, got: {result}"

    def test_placeholder_case_sensitivity(self):
        """Test that placeholder replacement is case-sensitive."""
        placeholders = {PLACEHOLDER_FILE_NAME: "TestFile"}

        # Exact match should work
        assert resolve_placeholders_in_value("{file.name}", placeholders) == "TestFile"

        # Wrong case should not be replaced
        assert resolve_placeholders_in_value("{FILE.NAME}", placeholders) == "{FILE.NAME}"
        assert resolve_placeholders_in_value("{File.Name}", placeholders) == "{File.Name}"

    def test_empty_placeholder_values(self):
        """Test behavior when placeholder values are empty."""
        placeholders = {
            PLACEHOLDER_FILE_NAME: "",
            PLACEHOLDER_FILE_PATH: "",
        }

        # Empty values should not replace (as per resolve_placeholders_in_value logic)
        test_string = "{file.path}/{file.name}-vba"
        result = resolve_placeholders_in_value(test_string, placeholders)

        # Should remain unchanged since values are empty
        assert result == "{file.path}/{file.name}-vba"

    def test_partial_placeholder_no_replacement(self):
        """Test that partial placeholder patterns are not replaced."""
        placeholders = {PLACEHOLDER_FILE_NAME: "MyDoc"}

        # Incomplete patterns should not be replaced
        test_cases = [
            "{file.name",  # Missing closing brace
            "file.name}",  # Missing opening brace
            "{file.nam}",  # Incomplete placeholder name
            "{file}",  # Missing property
        ]

        for test_string in test_cases:
            result = resolve_placeholders_in_value(test_string, placeholders)
            # Should remain unchanged
            assert test_string in result


class TestPlaceholderBackwardCompatibility:
    """Tests ensuring backward compatibility with legacy placeholder format."""

    def test_get_placeholder_values_returns_both_formats(self):
        """Test that get_placeholder_values() returns both new and legacy placeholders."""
        placeholders = get_placeholder_values(config_file_path="C:/test/config.toml", file_path="C:/test/doc.docx")

        # Should contain both new and legacy keys
        assert PLACEHOLDER_FILE_NAME in placeholders
        assert PLACEHOLDER_FILE_NAME_LEGACY in placeholders
        assert PLACEHOLDER_FILE_FULLNAME in placeholders
        assert PLACEHOLDER_FILE_FULLNAME_LEGACY in placeholders
        assert PLACEHOLDER_FILE_PATH in placeholders
        assert PLACEHOLDER_FILE_PATH_LEGACY in placeholders

    def test_both_formats_resolve_to_same_values(self):
        """Test that new and legacy formats resolve to identical values."""
        placeholders = get_placeholder_values(file_path="C:/documents/report.xlsx")

        # New and legacy should have same values
        assert placeholders[PLACEHOLDER_FILE_NAME] == placeholders[PLACEHOLDER_FILE_NAME_LEGACY]
        assert placeholders[PLACEHOLDER_FILE_FULLNAME] == placeholders[PLACEHOLDER_FILE_FULLNAME_LEGACY]
        assert placeholders[PLACEHOLDER_FILE_PATH] == placeholders[PLACEHOLDER_FILE_PATH_LEGACY]

    def test_migration_scenario_old_config_still_works(self):
        """Test that old config files with legacy placeholders still work."""
        placeholders = get_placeholder_values(
            config_file_path="C:/project/config.toml", file_path="C:/project/data/workbook.xlsm"
        )

        # Old-style path definition
        old_style_path = "{general.file.path}/{general.file.name}-modules"
        result = resolve_placeholders_in_value(old_style_path, placeholders)
        assert result == "C:\\project\\data/workbook-modules"

        # New-style path definition (should give same result)
        new_style_path = "{file.path}/{file.name}-modules"
        result_new = resolve_placeholders_in_value(new_style_path, placeholders)
        assert result_new == result

    def test_no_placeholder_values_without_file_path(self):
        """Test placeholder handling when no file path is provided."""
        placeholders = get_placeholder_values(config_file_path="C:/config.toml", file_path=None)

        # File-related placeholders should be empty
        assert placeholders[PLACEHOLDER_FILE_NAME] == ""
        assert placeholders[PLACEHOLDER_FILE_NAME_LEGACY] == ""
        assert placeholders[PLACEHOLDER_FILE_FULLNAME] == ""
        assert placeholders[PLACEHOLDER_FILE_FULLNAME_LEGACY] == ""

        # Config path should still be set (parent of config.toml is C:/)
        assert placeholders[PLACEHOLDER_CONFIG_PATH] == "C:\\"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

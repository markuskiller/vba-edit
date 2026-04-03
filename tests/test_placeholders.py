"""Tests for placeholder format (v0.5.0+, legacy placeholders removed)."""

import pytest

from vba_edit.cli_common import (
    PLACEHOLDER_CONFIG_PATH,
    PLACEHOLDER_FILE_FULLNAME,
    PLACEHOLDER_FILE_NAME,
    PLACEHOLDER_FILE_PATH,
    PLACEHOLDER_FILE_VBAPROJECT,
    get_placeholder_values,
    resolve_placeholders_in_value,
)


class TestSimplifiedPlaceholders:
    """Tests for placeholder format."""

    def test_placeholder_constants(self):
        """Test that placeholder constants are defined correctly."""
        assert PLACEHOLDER_FILE_NAME == "{file.name}"
        assert PLACEHOLDER_FILE_FULLNAME == "{file.fullname}"
        assert PLACEHOLDER_FILE_PATH == "{file.path}"
        assert PLACEHOLDER_FILE_VBAPROJECT == "{file.vbaproject}"
        assert PLACEHOLDER_CONFIG_PATH == "{config.path}"

    def test_placeholder_resolution(self):
        """Test that placeholders are resolved correctly."""
        placeholders = get_placeholder_values(
            config_file_path="C:/Projects/config.toml", file_path="C:/Projects/docs/MyDocument.docx"
        )

        # New format should be resolved
        assert placeholders[PLACEHOLDER_FILE_NAME] == "MyDocument"
        assert placeholders[PLACEHOLDER_FILE_FULLNAME] == "MyDocument.docx"
        assert placeholders[PLACEHOLDER_FILE_PATH] == "C:\\Projects\\docs"
        assert placeholders[PLACEHOLDER_CONFIG_PATH] == "C:\\Projects"

    def test_placeholder_in_string_replacement(self):
        """Test replacing placeholders in strings."""
        placeholders = {
            PLACEHOLDER_FILE_NAME: "MyDoc",
            PLACEHOLDER_FILE_PATH: "C:/Projects",
        }

        test_string = "{file.path}/{file.name}-vba"
        result = resolve_placeholders_in_value(test_string, placeholders)

        assert result == "C:/Projects/MyDoc-vba"

    def test_vbaproject_placeholder(self):
        """Test vbaproject placeholder format."""
        placeholders = {PLACEHOLDER_FILE_VBAPROJECT: "MyProject"}

        test_string = "Project: {file.vbaproject}"
        result = resolve_placeholders_in_value(test_string, placeholders)

        assert result == "Project: MyProject"

    def test_complex_path_with_placeholders(self):
        """Test complex path construction with placeholders."""
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


class TestPlaceholderNoLegacy:
    """Tests confirming legacy placeholders are no longer resolved."""

    def test_legacy_placeholders_not_in_values(self):
        """Legacy placeholder keys should NOT be present in get_placeholder_values()."""
        placeholders = get_placeholder_values(config_file_path="C:/test/config.toml", file_path="C:/test/doc.docx")

        assert "{general.file.name}" not in placeholders
        assert "{general.file.fullname}" not in placeholders
        assert "{general.file.path}" not in placeholders
        assert "{vbaproject}" not in placeholders

    def test_legacy_placeholder_strings_not_resolved(self):
        """Legacy placeholder strings should pass through unresolved."""
        placeholders = get_placeholder_values(
            config_file_path="C:/project/config.toml", file_path="C:/project/data/workbook.xlsm"
        )

        old_style_path = "{general.file.path}/{general.file.name}-modules"
        result = resolve_placeholders_in_value(old_style_path, placeholders)
        # Should remain unresolved
        assert result == old_style_path

    def test_no_placeholder_values_without_file_path(self):
        """Test placeholder handling when no file path is provided."""
        placeholders = get_placeholder_values(config_file_path="C:/config.toml", file_path=None)

        assert placeholders[PLACEHOLDER_FILE_NAME] == ""
        assert placeholders[PLACEHOLDER_FILE_FULLNAME] == ""

        # Config path should still be set
        assert placeholders[PLACEHOLDER_CONFIG_PATH] == "C:\\"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

"""Tests for enhanced help formatter."""

import argparse
import re

import pytest

from vba_edit.help_formatter import (
    CHECK_EXAMPLES,
    EDIT_EXAMPLES,
    EXPORT_EXAMPLES,
    IMPORT_EXAMPLES,
    EnhancedHelpFormatter,
    GroupedHelpFormatter,
    add_examples_epilog,
    create_parser_with_groups,
)


class TestEnhancedHelpFormatter:
    """Tests for EnhancedHelpFormatter class."""

    def test_formatter_initialization(self):
        """Test that formatter can be initialized."""
        formatter = EnhancedHelpFormatter("test-prog")
        assert formatter is not None
        assert formatter._prog == "test-prog"

    def test_simple_usage_formatting(self):
        """Test basic usage formatting."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=EnhancedHelpFormatter)
        parser.add_argument("command", choices=["edit", "export"])

        help_text = parser.format_help()

        # Should have usage line
        assert "usage:" in help_text.lower()
        assert "test-cmd" in help_text

    def test_usage_with_optional_arguments(self):
        """Test usage formatting with optional arguments."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=EnhancedHelpFormatter)
        parser.add_argument("--file", "-f", help="File path")
        parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")

        help_text = parser.format_help()

        # Should show help option
        assert "[-h]" in help_text or "[--help]" in help_text
        # Should mention optional args
        assert "--file" in help_text
        assert "--verbose" in help_text

    def test_usage_with_required_optional(self):
        """Test usage formatting with required optional argument."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=EnhancedHelpFormatter)
        parser.add_argument("--required", required=True, help="Required option")
        parser.add_argument("--optional", help="Optional option")

        help_text = parser.format_help()

        # Both should appear in help
        assert "--required" in help_text
        assert "--optional" in help_text

    def test_usage_with_positional_arguments(self):
        """Test usage formatting with positional arguments."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=EnhancedHelpFormatter)
        parser.add_argument("input_file", help="Input file")
        parser.add_argument("output_file", help="Output file")

        help_text = parser.format_help()

        # Positionals should be in usage
        assert "input_file" in help_text
        assert "output_file" in help_text

    def test_usage_with_choices(self):
        """Test usage formatting with choices."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=EnhancedHelpFormatter)
        parser.add_argument("action", choices=["start", "stop", "restart"])

        help_text = parser.format_help()

        # Should show choices in some form
        assert "start" in help_text
        assert "stop" in help_text
        assert "restart" in help_text

    def test_section_formatting(self):
        """Test section heading formatting."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=EnhancedHelpFormatter)
        parser.add_argument("--opt", help="Option")

        help_text = parser.format_help()

        # Should have standard sections (check for the text without worrying about markup tags)
        # Strip any rich markup tags for testing
        import re

        clean_text = re.sub(r"\[/?[^\]]+\]", "", help_text)
        assert (
            "positional arguments:" in clean_text.lower()
            or "optional arguments:" in clean_text.lower()
            or "options:" in clean_text.lower()
        )

    def test_version_action_formatting(self):
        """Test that version action is formatted correctly."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=EnhancedHelpFormatter)
        parser.add_argument("--version", "-V", action="version", version="1.0.0")

        help_text = parser.format_help()

        # Should show version option
        assert "--version" in help_text or "-V" in help_text


class TestGroupedHelpFormatter:
    """Tests for GroupedHelpFormatter class."""

    def test_grouped_formatter_initialization(self):
        """Test that grouped formatter can be initialized."""
        formatter = GroupedHelpFormatter("test-prog")
        assert formatter is not None

    def test_argument_group_creation(self):
        """Test creating argument groups."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=GroupedHelpFormatter)

        # Create groups
        required_group = parser.add_argument_group("Required Arguments", "These are required")
        optional_group = parser.add_argument_group("Optional Arguments", "These are optional")

        required_group.add_argument("--required", required=True, help="Required arg")
        optional_group.add_argument("--optional", help="Optional arg")

        help_text = parser.format_help()

        # Should show group titles
        assert "Required Arguments" in help_text
        assert "Optional Arguments" in help_text

    def test_multiple_groups_organization(self):
        """Test that multiple groups are organized correctly."""
        parser = argparse.ArgumentParser(prog="test-cmd", formatter_class=GroupedHelpFormatter)

        config_group = parser.add_argument_group("Configuration", "Config options")
        encoding_group = parser.add_argument_group("Encoding", "Encoding options")

        config_group.add_argument("--conf", help="Config file")
        encoding_group.add_argument("--encoding", help="Character encoding")

        help_text = parser.format_help()

        # Should show both groups
        assert "Configuration" in help_text
        assert "Encoding" in help_text
        assert "--conf" in help_text
        assert "--encoding" in help_text


class TestCreateParserWithGroups:
    """Tests for create_parser_with_groups helper function."""

    def test_creates_default_groups(self):
        """Test that default groups are created."""
        parser = argparse.ArgumentParser(formatter_class=GroupedHelpFormatter)
        groups = create_parser_with_groups(parser)

        # Should return dictionary of groups
        assert isinstance(groups, dict)
        assert "config" in groups
        assert "output" in groups

    def test_creates_optional_groups(self):
        """Test that optional groups are created when requested."""
        parser = argparse.ArgumentParser(formatter_class=GroupedHelpFormatter)
        groups = create_parser_with_groups(parser, include_encoding=True, include_headers=True, include_folders=True)

        assert "encoding" in groups
        assert "headers" in groups
        assert "folders" in groups

    def test_skips_optional_groups(self):
        """Test that optional groups are skipped when not requested."""
        parser = argparse.ArgumentParser(formatter_class=GroupedHelpFormatter)
        groups = create_parser_with_groups(parser, include_encoding=False, include_headers=False, include_folders=False)

        assert "encoding" not in groups
        assert "headers" not in groups
        assert "folders" not in groups

    def test_groups_have_correct_types(self):
        """Test that returned groups are argument groups."""
        parser = argparse.ArgumentParser(formatter_class=GroupedHelpFormatter)
        groups = create_parser_with_groups(parser)

        for name, group in groups.items():
            assert hasattr(group, "add_argument"), f"Group '{name}' should be an argument group"

    def test_can_add_arguments_to_groups(self):
        """Test that arguments can be added to created groups."""
        parser = argparse.ArgumentParser(formatter_class=GroupedHelpFormatter)
        groups = create_parser_with_groups(parser)

        # Should be able to add arguments to groups
        groups["config"].add_argument("--conf", help="Config file")
        groups["output"].add_argument("--verbose", action="store_true", help="Verbose")

        help_text = parser.format_help()
        assert "--conf" in help_text
        assert "--verbose" in help_text


class TestAddExamplesEpilog:
    """Tests for add_examples_epilog helper function."""

    def test_creates_examples_section(self):
        """Test that examples section is created."""
        examples = [
            ("Do something", "cmd --option"),
            ("Do something else", "cmd --other"),
        ]

        epilog = add_examples_epilog("test-cmd", examples)

        assert "Examples:" in epilog
        assert "Do something" in epilog
        assert "cmd --option" in epilog

    def test_formats_multiple_examples(self):
        """Test formatting of multiple examples."""
        examples = [
            ("First example", "cmd one"),
            ("Second example", "cmd two"),
            ("Third example", "cmd three"),
        ]

        epilog = add_examples_epilog("test-cmd", examples)

        for desc, cmd in examples:
            assert desc in epilog
            assert cmd in epilog

    def test_includes_comment_markers(self):
        """Test that example descriptions are formatted as comments."""
        examples = [("Test", "cmd")]
        epilog = add_examples_epilog("test-cmd", examples)

        # Should have comment marker for description
        assert "# Test" in epilog

    def test_empty_examples(self):
        """Test handling of empty examples list."""
        epilog = add_examples_epilog("test-cmd", [])

        # Should still have Examples header
        assert "Examples:" in epilog


class TestExampleConstants:
    """Tests for predefined example constants."""

    def test_edit_examples_exist(self):
        """Test that EDIT_EXAMPLES is defined."""
        assert EDIT_EXAMPLES is not None
        assert isinstance(EDIT_EXAMPLES, list)
        assert len(EDIT_EXAMPLES) > 0

    def test_import_examples_exist(self):
        """Test that IMPORT_EXAMPLES is defined."""
        assert IMPORT_EXAMPLES is not None
        assert isinstance(IMPORT_EXAMPLES, list)
        assert len(IMPORT_EXAMPLES) > 0

    def test_export_examples_exist(self):
        """Test that EXPORT_EXAMPLES is defined."""
        assert EXPORT_EXAMPLES is not None
        assert isinstance(EXPORT_EXAMPLES, list)
        assert len(EXPORT_EXAMPLES) > 0

    def test_check_examples_exist(self):
        """Test that CHECK_EXAMPLES is defined."""
        assert CHECK_EXAMPLES is not None
        assert isinstance(CHECK_EXAMPLES, list)
        assert len(CHECK_EXAMPLES) > 0

    def test_examples_have_correct_structure(self):
        """Test that examples have (description, command) structure."""
        for examples in [EDIT_EXAMPLES, IMPORT_EXAMPLES, EXPORT_EXAMPLES, CHECK_EXAMPLES]:
            for item in examples:
                assert isinstance(item, tuple), "Example should be a tuple"
                assert len(item) == 2, "Example should have 2 elements"
                desc, cmd = item
                assert isinstance(desc, str), "Description should be string"
                assert isinstance(cmd, str), "Command should be string"

    def test_examples_contain_commands(self):
        """Test that examples contain actual commands."""
        for examples in [EDIT_EXAMPLES, IMPORT_EXAMPLES, EXPORT_EXAMPLES, CHECK_EXAMPLES]:
            for desc, cmd in examples:
                # Should contain word-vba (or similar)
                assert "vba" in cmd.lower(), f"Command should mention vba: {cmd}"


class TestHelpOutputIntegration:
    """Integration tests for complete help output."""

    def test_complete_help_with_formatter(self):
        """Test complete help output with enhanced formatter."""
        parser = argparse.ArgumentParser(
            prog="test-vba", description="Test VBA tool", formatter_class=EnhancedHelpFormatter
        )

        subparsers = parser.add_subparsers(dest="command", required=True)

        # Edit command
        edit_parser = subparsers.add_parser("edit", help="Edit VBA")
        edit_parser.add_argument("--file", "-f", help="File path")
        edit_parser.add_argument("--verbose", "-v", action="store_true", help="Verbose")

        # Export command
        export_parser = subparsers.add_parser("export", help="Export VBA")
        export_parser.add_argument("--file", "-f", help="File path")

        help_text = parser.format_help()

        # Should have main sections
        assert "usage:" in help_text.lower()
        assert "test-vba" in help_text
        assert "edit" in help_text
        assert "export" in help_text

    def test_subcommand_help_with_formatter(self):
        """Test subcommand help output."""
        parser = argparse.ArgumentParser(prog="test-vba", formatter_class=EnhancedHelpFormatter)

        subparsers = parser.add_subparsers(dest="command")
        edit_parser = subparsers.add_parser("edit", help="Edit VBA", formatter_class=EnhancedHelpFormatter)
        edit_parser.add_argument("--file", "-f", help="File path")
        edit_parser.add_argument("--vba-directory", help="VBA directory")

        help_text = edit_parser.format_help()

        # Should show subcommand options
        assert "--file" in help_text
        assert "--vba-directory" in help_text

    def test_help_with_examples_epilog(self):
        """Test help output with examples epilog."""
        examples = [
            ("Edit active document", "word-vba edit"),
            ("Export with metadata", "word-vba export --save-metadata"),
        ]

        epilog = add_examples_epilog("word-vba", examples)

        parser = argparse.ArgumentParser(prog="word-vba", epilog=epilog, formatter_class=EnhancedHelpFormatter)
        parser.add_argument("--file", help="File path")

        help_text = parser.format_help()

        # Should include examples
        assert "Examples:" in help_text
        assert "Edit active document" in help_text
        assert "word-vba edit" in help_text


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

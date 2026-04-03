"""
Unit tests for reference classification, filtering, and CLI argument parsing.

These tests do NOT require Office — they test the pure-Python classification
and filtering logic in reference_manager.py, and CLI parser wiring.
"""

import pytest

from vba_edit.reference_manager import (
    THIRD_PARTY_GUIDS,
    THIRD_PARTY_NAME_PATTERNS,
    classify_reference,
    filter_references,
)
from vba_edit.excel_vba import create_cli_parser


# ---------------------------------------------------------------------------
# Test data helpers
# ---------------------------------------------------------------------------


def _make_ref(
    name="TestLib",
    guid="{00000000-0000-0000-0000-000000000001}",
    major=1,
    minor=0,
    builtin=False,
    broken=False,
    description="",
    path="",
):
    """Create a minimal reference dict for testing."""
    return {
        "name": name,
        "guid": guid,
        "major": major,
        "minor": minor,
        "priority": 1,
        "builtin": builtin,
        "broken": broken,
        "description": description,
        "path": path,
    }


# ---------------------------------------------------------------------------
# classify_reference
# ---------------------------------------------------------------------------


class TestClassifyReference:
    """Tests for classify_reference()."""

    def test_builtin_reference(self):
        ref = _make_ref(name="VBA", builtin=True)
        assert classify_reference(ref) == "builtin"

    def test_builtin_takes_precedence_over_guid(self):
        """Even if the GUID happens to be in the third-party list, BuiltIn wins."""
        guid = next(iter(THIRD_PARTY_GUIDS))
        ref = _make_ref(name="VBA", guid=guid, builtin=True)
        assert classify_reference(ref) == "builtin"

    @pytest.mark.parametrize("guid", list(THIRD_PARTY_GUIDS))
    def test_third_party_by_guid(self, guid):
        ref = _make_ref(name="SomeLib", guid=guid, builtin=False)
        assert classify_reference(ref) == "third-party"

    def test_third_party_guid_case_insensitive(self):
        guid = next(iter(THIRD_PARTY_GUIDS))
        ref = _make_ref(name="SomeLib", guid=guid.lower(), builtin=False)
        assert classify_reference(ref) == "third-party"

    @pytest.mark.parametrize("pattern", THIRD_PARTY_NAME_PATTERNS)
    def test_third_party_by_name_pattern(self, pattern):
        ref = _make_ref(name=f"My{pattern.title()}Lib", builtin=False)
        assert classify_reference(ref) == "third-party"

    def test_third_party_name_case_insensitive(self):
        ref = _make_ref(name="ACROBAT", builtin=False)
        assert classify_reference(ref) == "third-party"

    def test_custom_reference(self):
        ref = _make_ref(name="Scripting", guid="{420B2830-E718-11CF-893D-00A0C9054228}", builtin=False)
        assert classify_reference(ref) == "custom"

    def test_custom_reference_no_guid(self):
        ref = _make_ref(name="MyProject", guid="", builtin=False)
        assert classify_reference(ref) == "custom"

    def test_custom_reference_none_guid(self):
        ref = _make_ref(name="MyProject", guid=None, builtin=False)
        ref["guid"] = None
        assert classify_reference(ref) == "custom"


# ---------------------------------------------------------------------------
# filter_references
# ---------------------------------------------------------------------------


class TestFilterReferences:
    """Tests for filter_references()."""

    @pytest.fixture()
    def mixed_refs(self):
        """A list with one reference from each category."""
        return [
            _make_ref(name="VBA", builtin=True),
            _make_ref(name="Acrobat", guid=next(iter(THIRD_PARTY_GUIDS)), builtin=False),
            _make_ref(name="Scripting", guid="{420B2830-E718-11CF-893D-00A0C9054228}", builtin=False),
        ]

    def test_no_filters_returns_all(self, mixed_refs):
        result = filter_references(mixed_refs)
        assert len(result) == 3

    def test_no_builtins(self, mixed_refs):
        result = filter_references(mixed_refs, no_builtins=True)
        names = [r["name"] for r in result]
        assert "VBA" not in names
        assert len(result) == 2

    def test_no_third_party(self, mixed_refs):
        result = filter_references(mixed_refs, no_third_party=True)
        names = [r["name"] for r in result]
        assert "Acrobat" not in names
        assert len(result) == 2

    def test_no_custom(self, mixed_refs):
        result = filter_references(mixed_refs, no_custom=True)
        names = [r["name"] for r in result]
        assert "Scripting" not in names
        assert len(result) == 2

    def test_no_builtins_and_no_third_party(self, mixed_refs):
        result = filter_references(mixed_refs, no_builtins=True, no_third_party=True)
        assert len(result) == 1
        assert result[0]["name"] == "Scripting"

    def test_all_excluded_returns_empty(self, mixed_refs):
        result = filter_references(mixed_refs, no_builtins=True, no_third_party=True, no_custom=True)
        assert result == []

    def test_empty_input(self):
        result = filter_references([])
        assert result == []

    def test_multiple_custom_refs(self):
        refs = [
            _make_ref(name="Scripting", guid="{420B2830-E718-11CF-893D-00A0C9054228}"),
            _make_ref(name="MSXML2", guid="{F5078F18-C551-11D3-89B9-0000F81FE221}"),
        ]
        result = filter_references(refs, no_builtins=True, no_third_party=True)
        assert len(result) == 2


# ---------------------------------------------------------------------------
# CLI argument parsing for references subcommands
# ---------------------------------------------------------------------------


class TestReferencesCLIParsing:
    """Tests that references subcommand arguments are wired correctly."""

    @pytest.fixture()
    def parser(self):
        return create_cli_parser()

    def test_references_list_defaults(self, parser):
        args = parser.parse_args(["references", "list"])
        assert args.refs_subcommand == "list"
        assert args.no_builtins is False
        assert args.no_third_party is False
        assert args.no_custom is False

    def test_references_list_no_builtins(self, parser):
        args = parser.parse_args(["references", "list", "--no-builtins"])
        assert args.no_builtins is True

    def test_references_list_no_third_party(self, parser):
        args = parser.parse_args(["references", "list", "--no-third-party"])
        assert args.no_third_party is True

    def test_references_list_no_custom(self, parser):
        args = parser.parse_args(["references", "list", "--no-custom"])
        assert args.no_custom is True

    def test_references_list_combined_filters(self, parser):
        args = parser.parse_args(["references", "list", "--no-builtins", "--no-third-party"])
        assert args.no_builtins is True
        assert args.no_third_party is True
        assert args.no_custom is False

    def test_references_export_with_filters(self, parser):
        args = parser.parse_args(["references", "export", "--no-builtins", "--no-third-party"])
        assert args.refs_subcommand == "export"
        assert args.no_builtins is True
        assert args.no_third_party is True

    def test_references_export_with_refs_file(self, parser):
        args = parser.parse_args(["references", "export", "-r", "custom.toml"])
        assert args.refs_file == "custom.toml"

    def test_references_import_subcommand(self, parser):
        args = parser.parse_args(["references", "import", "-r", "refs.toml"])
        assert args.refs_subcommand == "import"
        assert args.refs_file == "refs.toml"

    def test_references_validate_subcommand(self, parser):
        args = parser.parse_args(["references", "validate"])
        assert args.refs_subcommand == "validate"

    def test_references_validate_with_file(self, parser):
        args = parser.parse_args(["references", "validate", "-f", "test.xlsm"])
        assert args.refs_subcommand == "validate"
        assert args.file == "test.xlsm"

    def test_references_add_subcommand(self, parser):
        args = parser.parse_args(["references", "add", "SharedLib.xlam"])
        assert args.refs_subcommand == "add"
        assert args.library == "SharedLib.xlam"

    def test_references_add_with_file(self, parser):
        args = parser.parse_args(["references", "add", "Lib.dotm", "-f", "doc.docm"])
        assert args.library == "Lib.dotm"
        assert args.file == "doc.docm"

    def test_references_remove_subcommand(self, parser):
        args = parser.parse_args(["references", "remove", "OldLibrary"])
        assert args.refs_subcommand == "remove"
        assert args.ref_name == "OldLibrary"

    def test_references_remove_with_file(self, parser):
        args = parser.parse_args(["references", "remove", "OldLib", "-f", "doc.xlsm"])
        assert args.ref_name == "OldLib"
        assert args.file == "doc.xlsm"

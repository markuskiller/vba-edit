"""
Unit tests for reference classification and filtering.

These tests do NOT require Office — they test the pure-Python classification
and filtering logic in reference_manager.py.
"""

import pytest

from vba_edit.reference_manager import (
    THIRD_PARTY_GUIDS,
    THIRD_PARTY_NAME_PATTERNS,
    classify_reference,
    filter_references,
)


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

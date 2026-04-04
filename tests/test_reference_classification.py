"""
Unit tests for reference classification, filtering, and CLI argument parsing.

These tests do NOT require Office — they test the pure-Python classification
and filtering logic in reference_manager.py, and CLI parser wiring.
"""

import pytest
from pathlib import Path

from vba_edit.reference_manager import (
    DEFAULT_GUIDS,
    DEFAULT_NAMES,
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

    def test_com_builtin_flag_returns_default(self):
        ref = _make_ref(name="VBA", builtin=True)
        assert classify_reference(ref) == "default"

    def test_com_builtin_flag_takes_precedence(self):
        """COM BuiltIn flag always wins, even with arbitrary GUID."""
        ref = _make_ref(name="VBA", guid="{99999999-9999-9999-9999-999999999999}", builtin=True)
        assert classify_reference(ref) == "default"

    @pytest.mark.parametrize("guid", list(DEFAULT_GUIDS))
    def test_default_by_guid(self, guid):
        """References with well-known Office framework GUIDs are default."""
        ref = _make_ref(name="SomeLib", guid=guid, builtin=False)
        assert classify_reference(ref) == "default"

    def test_default_guid_case_insensitive(self):
        guid = next(iter(DEFAULT_GUIDS))
        ref = _make_ref(name="SomeLib", guid=guid.lower(), builtin=False)
        assert classify_reference(ref) == "default"

    @pytest.mark.parametrize("name", list(DEFAULT_NAMES))
    def test_default_by_name(self, name):
        """References with well-known Office framework names are default."""
        ref = _make_ref(name=name, guid="", builtin=False)
        assert classify_reference(ref) == "default"

    def test_default_name_case_insensitive(self):
        ref = _make_ref(name="STDOLE", guid="", builtin=False)
        assert classify_reference(ref) == "default"

    def test_stdole_by_guid(self):
        """stdole (OLE Automation) should be classified as default."""
        ref = _make_ref(name="stdole", guid="{00020430-0000-0000-C000-000000000046}", builtin=False)
        assert classify_reference(ref) == "default"

    def test_office_library_by_guid(self):
        """Microsoft Office Object Library should be classified as default."""
        ref = _make_ref(name="Office", guid="{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", builtin=False)
        assert classify_reference(ref) == "default"

    def test_normal_template_by_name(self):
        """Word Normal.dotm reference should be classified as default."""
        ref = _make_ref(name="Normal", guid="", builtin=False)
        assert classify_reference(ref) == "default"

    def test_installed_reference_with_guid(self):
        """COM library with a GUID (not in DEFAULT_GUIDS) is installed."""
        ref = _make_ref(name="Scripting", guid="{420B2830-E718-11CF-893D-00A0C9054228}", builtin=False)
        assert classify_reference(ref) == "installed"

    def test_installed_reference_acrobat(self):
        """COM library like Acrobat is classified as installed."""
        ref = _make_ref(name="Acrobat", guid="{E64169B3-3592-47D2-816E-602C5C13F328}", builtin=False)
        assert classify_reference(ref) == "installed"

    def test_custom_reference_no_guid(self):
        """File-path reference without GUID is custom."""
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
            _make_ref(name="VBA", builtin=True),  # default
            _make_ref(name="Scripting", guid="{420B2830-E718-11CF-893D-00A0C9054228}", builtin=False),  # installed
            _make_ref(name="MyTemplate", guid="", builtin=False),  # custom
        ]

    def test_no_filters_returns_all(self, mixed_refs):
        result = filter_references(mixed_refs)
        assert len(result) == 3

    def test_no_default(self, mixed_refs):
        result = filter_references(mixed_refs, no_default=True)
        names = [r["name"] for r in result]
        assert "VBA" not in names
        assert len(result) == 2

    def test_no_installed(self, mixed_refs):
        result = filter_references(mixed_refs, no_installed=True)
        names = [r["name"] for r in result]
        assert "Scripting" not in names
        assert len(result) == 2

    def test_no_custom(self, mixed_refs):
        result = filter_references(mixed_refs, no_custom=True)
        names = [r["name"] for r in result]
        assert "MyTemplate" not in names
        assert len(result) == 2

    def test_no_default_and_no_installed(self, mixed_refs):
        result = filter_references(mixed_refs, no_default=True, no_installed=True)
        assert len(result) == 1
        assert result[0]["name"] == "MyTemplate"

    def test_all_excluded_returns_empty(self, mixed_refs):
        result = filter_references(mixed_refs, no_default=True, no_installed=True, no_custom=True)
        assert result == []

    def test_empty_input(self):
        result = filter_references([])
        assert result == []

    def test_multiple_installed_refs(self):
        refs = [
            _make_ref(name="MyLib", guid="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}"),
            _make_ref(name="MSXML2", guid="{F5078F18-C551-11D3-89B9-0000F81FE221}"),
        ]
        result = filter_references(refs, no_default=True, no_custom=True)
        assert len(result) == 2

    def test_default_guid_refs_filtered_by_no_default(self):
        """References classified as default via DEFAULT_GUIDS should be filtered."""
        refs = [
            _make_ref(name="VBA", builtin=True),
            _make_ref(name="stdole", guid="{00020430-0000-0000-C000-000000000046}", builtin=False),
            _make_ref(name="MyLib", guid="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}"),
        ]
        result = filter_references(refs, no_default=True)
        assert len(result) == 1
        assert result[0]["name"] == "MyLib"


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
        assert args.no_default is False
        assert args.no_installed is False
        assert args.no_custom is False

    def test_references_list_no_default(self, parser):
        args = parser.parse_args(["references", "list", "--no-default"])
        assert args.no_default is True

    def test_references_list_no_installed(self, parser):
        args = parser.parse_args(["references", "list", "--no-installed"])
        assert args.no_installed is True

    def test_references_list_no_custom(self, parser):
        args = parser.parse_args(["references", "list", "--no-custom"])
        assert args.no_custom is True

    def test_references_list_combined_filters(self, parser):
        args = parser.parse_args(["references", "list", "--no-default", "--no-installed"])
        assert args.no_default is True
        assert args.no_installed is True
        assert args.no_custom is False

    def test_references_export_with_filters(self, parser):
        args = parser.parse_args(["references", "export", "--no-default", "--no-installed"])
        assert args.refs_subcommand == "export"
        assert args.no_default is True
        assert args.no_installed is True

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


# ---------------------------------------------------------------------------
# --with-references flag parsing
# ---------------------------------------------------------------------------


class TestWithReferencesFlag:
    """Tests that --with-references is available on export, import, and edit commands."""

    @pytest.fixture()
    def parser(self):
        return create_cli_parser()

    def test_export_with_references_default(self, parser):
        args = parser.parse_args(["export"])
        assert args.with_references is False

    def test_export_with_references_set(self, parser):
        args = parser.parse_args(["export", "--with-references"])
        assert args.with_references is True

    def test_import_with_references_default(self, parser):
        args = parser.parse_args(["import"])
        assert args.with_references is False

    def test_import_with_references_set(self, parser):
        args = parser.parse_args(["import", "--with-references"])
        assert args.with_references is True

    def test_edit_with_references_default(self, parser):
        args = parser.parse_args(["edit"])
        assert args.with_references is False

    def test_edit_with_references_set(self, parser):
        args = parser.parse_args(["edit", "--with-references"])
        assert args.with_references is True


# ---------------------------------------------------------------------------
# TOML metadata and path normalization
# ---------------------------------------------------------------------------


class TestTomlMetadata:
    """Tests for TOML serialization: metadata section, path normalization, version checking."""

    def test_metadata_section_present(self):
        """Serialized TOML contains a [metadata] section."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        output = _serialize_references_to_toml([], document_name="Book1.xlsm")
        assert "[metadata]" in output
        assert 'generated_by = "vba-edit' in output
        assert "timestamp = " in output
        assert 'document = "Book1.xlsm"' in output

    def test_metadata_no_document(self):
        """Metadata omits document key when no name is provided."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        output = _serialize_references_to_toml([])
        assert "[metadata]" in output
        assert "document = " not in output

    def test_path_normalization_in_toml(self, tmp_path):
        """Paths are resolved to absolute before appearing in TOML."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        ref = _make_ref(name="TestLib", path=str(tmp_path / "lib.dll"))
        output = _serialize_references_to_toml([ref])
        # The path should be in the output and should be absolute
        assert "path = " in output
        assert "lib.dll" in output

    def test_version_in_metadata(self):
        """The generated_by field contains the current vba-edit version."""
        from vba_edit.reference_manager import _serialize_references_to_toml, _get_version

        output = _serialize_references_to_toml([])
        version = _get_version()
        assert f'generated_by = "vba-edit {version}"' in output

    def test_document_name_escaping(self):
        """Document names with special chars are escaped in TOML."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        output = _serialize_references_to_toml([], document_name='Book "1".xlsm')
        assert 'document = "Book \\"1\\".xlsm"' in output


# ---------------------------------------------------------------------------
# Auto-export / auto-import helpers
# ---------------------------------------------------------------------------


class TestAutoReferenceHelpers:
    """Tests for the OfficeVBACLI auto-export/import reference helper methods."""

    def test_get_refs_file_derives_correct_path(self):
        """_get_refs_file derives {stem}_refs.toml in vba_dir."""
        from unittest.mock import MagicMock
        from vba_edit.office_cli import OfficeVBACLI

        cli = OfficeVBACLI.__new__(OfficeVBACLI)
        handler = MagicMock()
        handler.doc_path = Path("C:/docs/MyWorkbook.xlsm")
        handler.vba_dir = Path("C:/docs/VBA-MyWorkbook")

        result = cli._get_refs_file(handler)
        assert result == Path("C:/docs/VBA-MyWorkbook/MyWorkbook_refs.toml")

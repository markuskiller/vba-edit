"""
Unit tests for reference classification, filtering, and CLI argument parsing.

These tests do NOT require Office — they test the pure-Python classification
and filtering logic in reference_manager.py, and CLI parser wiring.
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

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


# ---------------------------------------------------------------------------
# TOML filter metadata
# ---------------------------------------------------------------------------


class TestTomlFilterMetadata:
    """Tests for filter metadata in exported TOML files."""

    def test_no_filters_omits_key(self):
        """When no filters are active, the filters key is absent."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        output = _serialize_references_to_toml([], document_name="Book1.xlsm")
        assert "filters" not in output

    def test_single_filter_recorded(self):
        """A single active filter is recorded in metadata."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        output = _serialize_references_to_toml([], filters=["no_default"])
        assert 'filters = ["no_default"]' in output

    def test_multiple_filters_recorded(self):
        """Multiple active filters are recorded."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        output = _serialize_references_to_toml([], filters=["no_default", "no_installed"])
        assert 'filters = ["no_default", "no_installed"]' in output

    def test_filters_none_omits_key(self):
        """Passing filters=None omits the key."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        output = _serialize_references_to_toml([], filters=None)
        assert "filters" not in output

    def test_filters_parseable_as_toml(self):
        """The serialized filters value can be parsed back as valid TOML."""
        from vba_edit.reference_manager import _serialize_references_to_toml, _load_toml
        import tempfile

        output = _serialize_references_to_toml(
            [_make_ref()],
            document_name="Test.xlsm",
            filters=["no_default", "no_custom"],
        )
        with tempfile.NamedTemporaryFile(mode="w", suffix=".toml", delete=False, encoding="utf-8") as f:
            f.write(output)
            f.flush()
            data = _load_toml(Path(f.name))

        assert data["metadata"]["filters"] == ["no_default", "no_custom"]


# ---------------------------------------------------------------------------
# --sync CLI argument parsing
# ---------------------------------------------------------------------------


class TestSyncCLIParsing:
    """Tests that --sync and --force-overwrite are wired on references import."""

    @pytest.fixture()
    def parser(self):
        return create_cli_parser()

    def test_import_sync_default_false(self, parser):
        args = parser.parse_args(["references", "import", "-r", "refs.toml"])
        assert args.sync is False

    def test_import_sync_flag(self, parser):
        args = parser.parse_args(["references", "import", "-r", "refs.toml", "--sync"])
        assert args.sync is True

    def test_import_force_overwrite_default_false(self, parser):
        args = parser.parse_args(["references", "import", "-r", "refs.toml"])
        assert args.force_overwrite is False

    def test_import_force_overwrite_flag(self, parser):
        args = parser.parse_args(["references", "import", "-r", "refs.toml", "--force-overwrite"])
        assert args.force_overwrite is True

    def test_import_sync_with_force_overwrite(self, parser):
        args = parser.parse_args(["references", "import", "-r", "refs.toml", "--sync", "--force-overwrite"])
        assert args.sync is True
        assert args.force_overwrite is True

    def test_references_no_subcommand_exits_with_help(self, parser):
        """'references' without subcommand should exit (help display)."""
        # With required=False, parse succeeds but refs_subcommand is None
        args = parser.parse_args(["references"])
        assert args.refs_subcommand is None


# ---------------------------------------------------------------------------
# sync_from_toml logic (mocked — no Office needed)
# ---------------------------------------------------------------------------


class TestSyncFromToml:
    """Tests for ReferenceManager.sync_from_toml() with mocked COM objects."""

    def _make_toml_file(self, tmp_path, refs, metadata=None):
        """Create a TOML file with given references and optional metadata."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        filters = metadata.get("filters") if metadata else None
        content = _serialize_references_to_toml(refs, filters=filters)
        toml_path = tmp_path / "refs.toml"
        toml_path.write_text(content, encoding="utf-8")
        return toml_path

    def _make_mock_manager(self, existing_refs):
        """Create a ReferenceManager with mocked COM objects."""
        from vba_edit.reference_manager import ReferenceManager

        manager = ReferenceManager.__new__(ReferenceManager)
        manager.document = MagicMock()
        manager.vb_project = MagicMock()

        manager.list_references = MagicMock(return_value=existing_refs)
        manager.add_reference = MagicMock(return_value=True)
        manager.remove_reference = MagicMock(return_value=True)
        manager.reference_exists = MagicMock(return_value=False)

        return manager

    def test_sync_adds_missing_references(self, tmp_path):
        """Sync adds references present in TOML but not in document."""
        toml_refs = [_make_ref(name="NewLib", guid="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}")]
        toml_path = self._make_toml_file(tmp_path, toml_refs)

        manager = self._make_mock_manager(existing_refs=[])
        stats = manager.sync_from_toml(toml_path)

        assert stats["added"] == 1
        manager.add_reference.assert_called_once()

    def test_sync_removes_extra_references(self, tmp_path):
        """Sync removes references in document but not in TOML."""
        toml_path = self._make_toml_file(tmp_path, [])  # Empty TOML

        existing = [
            _make_ref(
                name="OldLib",
                guid="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}",
                builtin=False,
            )
        ]
        manager = self._make_mock_manager(existing_refs=existing)
        stats = manager.sync_from_toml(toml_path)

        assert stats["removed"] == 1
        manager.remove_reference.assert_called_once()

    def test_sync_protects_default_references(self, tmp_path):
        """Sync does not remove default references unless force_overwrite=True."""
        toml_path = self._make_toml_file(tmp_path, [])  # Empty TOML

        existing = [_make_ref(name="VBA", builtin=True)]
        manager = self._make_mock_manager(existing_refs=existing)
        stats = manager.sync_from_toml(toml_path)

        assert stats["protected"] == 1
        assert stats["removed"] == 0
        manager.remove_reference.assert_not_called()

    def test_sync_force_overwrite_removes_defaults(self, tmp_path):
        """With force_overwrite, default references can be removed."""
        # Need a TOML with a valid ref so it's not completely empty
        toml_path = self._make_toml_file(tmp_path, [])

        existing = [
            _make_ref(
                name="stdole",
                guid="{00020430-0000-0000-C000-000000000046}",
                builtin=False,
            )
        ]
        manager = self._make_mock_manager(existing_refs=existing)
        stats = manager.sync_from_toml(toml_path, force_overwrite=True)

        assert stats["removed"] == 1

    def test_sync_refuses_filtered_toml_without_force(self, tmp_path):
        """Sync refuses when TOML was exported with filters."""
        from vba_edit.exceptions import VBAReferenceError

        toml_path = self._make_toml_file(tmp_path, [], metadata={"filters": ["no_default", "no_installed"]})
        manager = self._make_mock_manager(existing_refs=[])

        with pytest.raises(VBAReferenceError, match="exported with filters"):
            manager.sync_from_toml(toml_path)

    def test_sync_allows_filtered_toml_with_force(self, tmp_path):
        """Sync proceeds on filtered TOML when force_overwrite=True."""
        toml_path = self._make_toml_file(tmp_path, [], metadata={"filters": ["no_default"]})
        manager = self._make_mock_manager(existing_refs=[])

        stats = manager.sync_from_toml(toml_path, force_overwrite=True)
        assert stats["added"] == 0  # Nothing to add

    def test_sync_skips_existing_references(self, tmp_path):
        """References already in the document are skipped (not re-added)."""
        guid = "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}"
        toml_refs = [_make_ref(name="MyLib", guid=guid)]
        toml_path = self._make_toml_file(tmp_path, toml_refs)

        existing = [_make_ref(name="MyLib", guid=guid)]
        manager = self._make_mock_manager(existing_refs=existing)
        manager.add_reference = MagicMock(return_value=False)  # Already exists

        stats = manager.sync_from_toml(toml_path)

        assert stats["skipped"] == 1
        assert stats["removed"] == 0

    def test_sync_file_not_found(self, tmp_path):
        """Sync raises FileNotFoundError for missing files."""
        manager = self._make_mock_manager(existing_refs=[])

        with pytest.raises(FileNotFoundError):
            manager.sync_from_toml(tmp_path / "nonexistent.toml")

    def test_sync_skips_refs_without_guid(self, tmp_path):
        """References without a GUID are protected (can't be managed by GUID)."""
        toml_path = self._make_toml_file(tmp_path, [])

        existing = [_make_ref(name="Normal", guid="", builtin=False)]
        manager = self._make_mock_manager(existing_refs=existing)
        stats = manager.sync_from_toml(toml_path)

        assert stats["protected"] == 1
        assert stats["removed"] == 0


# ---------------------------------------------------------------------------
# Import filter warning
# ---------------------------------------------------------------------------


class TestImportFilterWarning:
    """Tests that import_from_toml warns about filtered TOML files."""

    def _make_toml_with_filters(self, tmp_path, filters):
        """Create a TOML file with filter metadata and one valid reference."""
        from vba_edit.reference_manager import _serialize_references_to_toml

        ref = _make_ref(name="TestLib", guid="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}")
        content = _serialize_references_to_toml([ref], filters=filters)
        toml_path = tmp_path / "filtered_refs.toml"
        toml_path.write_text(content, encoding="utf-8")
        return toml_path

    def test_import_warns_on_filtered_toml(self, tmp_path, caplog):
        """import_from_toml logs a warning when TOML has filter metadata."""
        import logging
        from vba_edit.reference_manager import ReferenceManager

        toml_path = self._make_toml_with_filters(tmp_path, ["no_default", "no_installed"])

        manager = ReferenceManager.__new__(ReferenceManager)
        manager.document = MagicMock()
        manager.vb_project = MagicMock()
        manager.add_reference = MagicMock(return_value=True)
        manager.reference_exists = MagicMock(return_value=False)

        with caplog.at_level(logging.WARNING, logger="vba_edit.reference_manager"):
            manager.import_from_toml(toml_path)

        assert any("exported with filters" in msg for msg in caplog.messages)

    def test_import_no_warning_on_unfiltered_toml(self, tmp_path, caplog):
        """import_from_toml does not warn when TOML has no filter metadata."""
        import logging
        from vba_edit.reference_manager import ReferenceManager

        toml_path = self._make_toml_with_filters(tmp_path, None)

        manager = ReferenceManager.__new__(ReferenceManager)
        manager.document = MagicMock()
        manager.vb_project = MagicMock()
        manager.add_reference = MagicMock(return_value=True)
        manager.reference_exists = MagicMock(return_value=False)

        with caplog.at_level(logging.WARNING, logger="vba_edit.reference_manager"):
            manager.import_from_toml(toml_path)

        assert not any("exported with filters" in msg for msg in caplog.messages)


# ---------------------------------------------------------------------------
# Regression test: mutating references commands must call doc.Save()
# Issue #99: references add / import / remove did not persist changes to disk
# ---------------------------------------------------------------------------


class TestReferencesSavePersistence:
    """Regression tests for issue #99.

    Verifies that _handle_references_command calls doc.Save() after mutating
    subcommands (add, import, remove) so changes are persisted to disk.
    Read-only subcommands (list, export, validate) must NOT trigger a save.
    """

    def _make_args(self, subcommand, file_path, **kwargs):
        """Build a minimal argparse.Namespace for _handle_references_command."""
        import argparse

        defaults = dict(
            refs_subcommand=subcommand,
            file=str(file_path),
            refs_file=None,
            verbose=False,
            logfile=None,
            no_default=False,
            no_installed=False,
            no_custom=False,
            sync=False,
            force_overwrite=False,
        )
        defaults.update(kwargs)
        return argparse.Namespace(**defaults)

    def _run_with_mocks(self, subcommand, tmp_path, extra_args=None):
        """Run _handle_references_command with fully mocked COM layer.

        Returns the mock document so callers can assert on doc.Save().
        """
        from vba_edit.office_cli import OfficeVBACLI
        from vba_edit.reference_manager import ReferenceManager

        # Create a real (but empty) dotm file so Path(...).resolve() works
        fake_doc = tmp_path / "test.dotm"
        fake_doc.write_text("")

        args = self._make_args(subcommand, fake_doc, **(extra_args or {}))

        mock_doc = MagicMock()
        mock_app = MagicMock()
        mock_app.Documents.Open.return_value = mock_doc

        # Build a ReferenceManager whose COM methods are all no-ops
        mock_manager = MagicMock(spec=ReferenceManager)
        mock_manager.add_reference_by_path.return_value = True
        mock_manager.import_from_toml.return_value = {"added": 1, "skipped": 0, "failed": 0}
        mock_manager.remove_reference.return_value = True
        mock_manager.list_references.return_value = []
        mock_manager.export_to_toml.return_value = []

        cli = OfficeVBACLI("word")

        with (
            patch("win32com.client.Dispatch", return_value=mock_app),
            patch("vba_edit.office_cli.ReferenceManager", return_value=mock_manager),
            patch("vba_edit.office_cli.setup_logging"),
        ):
            try:
                cli._handle_references_command(args)
            except SystemExit as exc:
                # list / validate call sys.exit(0) on success — that's fine
                if exc.code != 0:
                    raise

        return mock_doc

    @pytest.mark.parametrize("subcommand", ["add", "import", "remove"])
    def test_mutating_subcommand_saves_document(self, subcommand, tmp_path):
        """add / import / remove must call doc.Save() to persist changes (issue #99)."""
        extra = {}
        if subcommand == "add":
            lib = tmp_path / "MyLib.dotm"
            lib.write_text("")
            extra["library"] = str(lib)
        elif subcommand == "import":
            refs = tmp_path / "refs.toml"
            refs.write_text(
                '[[references]]\nname = "TestLib"\nguid = "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}"\nmajor = 1\nminor = 0\n'
            )
            extra["refs_file"] = str(refs)
        elif subcommand == "remove":
            extra["ref_name"] = "OldLib"

        mock_doc = self._run_with_mocks(subcommand, tmp_path, extra)
        mock_doc.Save.assert_called_once()

    @pytest.mark.parametrize("subcommand", ["list", "validate"])
    def test_readonly_subcommand_does_not_save(self, subcommand, tmp_path):
        """list / validate must NOT call doc.Save() — they don't change anything."""
        mock_doc = self._run_with_mocks(subcommand, tmp_path)
        mock_doc.Save.assert_not_called()

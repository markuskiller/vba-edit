"""
Integration tests for VBA Reference Management.

These tests require Excel to be installed and VBA trust enabled.
They test the ReferenceManager with real Excel COM objects.

Run with: pytest tests/test_reference_manager_integration.py -v
Skip with: pytest tests/ -v (integration tests auto-skip without explicit run)
"""

import tempfile
from pathlib import Path

import pytest
import win32com.client

from vba_edit.exceptions import ReferenceError
from vba_edit.reference_manager import ReferenceManager


@pytest.fixture(scope="module")
def excel_app():
    """Create Excel application for integration tests."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    yield excel

    try:
        for i in range(excel.Workbooks.Count, 0, -1):
            try:
                excel.Workbooks(i).Saved = True
                excel.Workbooks(i).Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
    except Exception:
        pass


@pytest.fixture
def test_workbook(excel_app):
    """Create a test workbook for each test."""
    wb = excel_app.Workbooks.Add()
    yield wb

    try:
        wb.Saved = True
        wb.Close(SaveChanges=False)
    except Exception:
        pass


@pytest.mark.integration
@pytest.mark.skip(reason="Integration test - run explicitly with pytest tests/test_reference_manager_integration.py")
class TestReferenceManagerIntegration:
    """Integration tests for ReferenceManager with real Excel."""

    def test_list_references_real_excel(self, test_workbook):
        """Test listing references with real Excel workbook."""
        manager = ReferenceManager(test_workbook)
        refs = manager.list_references()

        assert len(refs) > 0

        first_ref = refs[0]
        assert "name" in first_ref
        assert "guid" in first_ref
        assert "major" in first_ref
        assert "minor" in first_ref
        assert "priority" in first_ref
        assert "builtin" in first_ref
        assert "broken" in first_ref

        builtin_refs = [r for r in refs if r["builtin"]]
        assert len(builtin_refs) > 0

        broken_refs = [r for r in refs if r["broken"]]
        assert len(broken_refs) == 0

    def test_add_and_remove_reference_real_excel(self, test_workbook):
        """Test adding and removing a reference with real Excel."""
        manager = ReferenceManager(test_workbook)

        refs_before = manager.list_references()
        initial_count = len(refs_before)

        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        added = manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)
        assert added is True

        refs_after_add = manager.list_references()
        assert len(refs_after_add) == initial_count + 1

        scripting_ref = next((r for r in refs_after_add if r["guid"].upper() == scripting_guid.upper()), None)
        assert scripting_ref is not None
        assert scripting_ref["name"] == "Scripting"
        assert scripting_ref["builtin"] is False

        removed = manager.remove_reference(guid=scripting_guid)
        assert removed is True

        refs_after_remove = manager.list_references()
        assert len(refs_after_remove) == initial_count

        scripting_ref = next((r for r in refs_after_remove if r["guid"].upper() == scripting_guid.upper()), None)
        assert scripting_ref is None

    def test_add_duplicate_reference_skips(self, test_workbook):
        """Test that adding duplicate reference is skipped."""
        manager = ReferenceManager(test_workbook)

        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        added_first = manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)
        assert added_first is True

        added_second = manager.add_reference(
            guid=scripting_guid, name="Scripting", major=1, minor=0, skip_if_exists=True
        )
        assert added_second is False

    def test_reference_exists_real_excel(self, test_workbook):
        """Test checking if reference exists with real Excel."""
        manager = ReferenceManager(test_workbook)

        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        assert manager.reference_exists(guid=scripting_guid) is False

        manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)
        assert manager.reference_exists(guid=scripting_guid) is True
        assert manager.reference_exists(name="Scripting") is True

    def test_cannot_remove_builtin_reference(self, test_workbook):
        """Test that built-in references cannot be removed."""
        manager = ReferenceManager(test_workbook)

        refs = manager.list_references()
        builtin_ref = next((r for r in refs if r["builtin"]), None)
        assert builtin_ref is not None

        with pytest.raises(ReferenceError, match="Cannot remove built-in reference"):
            manager.remove_reference(guid=builtin_ref["guid"])

    def test_export_import_toml_roundtrip(self, test_workbook):
        """Test exporting and importing references via TOML."""
        manager = ReferenceManager(test_workbook)

        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

        with tempfile.TemporaryDirectory() as tmpdir:
            toml_file = Path(tmpdir) / "refs.toml"
            manager.export_to_toml(toml_file)

            assert toml_file.exists()

            content = toml_file.read_text(encoding="utf-8")
            assert "Scripting" in content
            assert scripting_guid in content
            assert "[[references]]" in content

            manager.remove_reference(guid=scripting_guid)
            assert manager.reference_exists(guid=scripting_guid) is False

            stats = manager.import_from_toml(toml_file)
            assert stats["added"] >= 1
            assert manager.reference_exists(guid=scripting_guid) is True

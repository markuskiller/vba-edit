"""
Integration tests for VBA Reference Management with Word.

These tests focus on Word-specific scenarios.

Run with: pytest tests/test_reference_manager_word_integration.py -v
"""

import tempfile
from pathlib import Path

import pytest
import win32com.client

from vba_edit.exceptions import ReferenceError
from vba_edit.reference_manager import ReferenceManager


@pytest.fixture(scope="module")
def word_app():
    """Create Word application for integration tests."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone
    yield word

    try:
        for i in range(word.Documents.Count, 0, -1):
            try:
                word.Documents(i).Saved = True
                word.Documents(i).Close(SaveChanges=False)
            except Exception:
                pass
        word.Quit()
    except Exception:
        pass


@pytest.fixture
def test_document(word_app):
    """Create a test Word document for each test."""
    doc = word_app.Documents.Add()
    yield doc

    try:
        doc.Saved = True
        doc.Close(SaveChanges=False)
    except Exception:
        pass


@pytest.mark.integration
@pytest.mark.skip(
    reason="Integration test - run explicitly with pytest tests/test_reference_manager_word_integration.py"
)
class TestWordReferenceManagerIntegration:
    """Integration tests for ReferenceManager with real Word documents."""

    def test_list_references_word_document(self, test_document):
        """Test listing references in a Word document."""
        manager = ReferenceManager(test_document)
        refs = manager.list_references()

        assert len(refs) > 0

        for ref in refs:
            assert "name" in ref
            assert "guid" in ref
            assert "major" in ref
            assert "minor" in ref
            assert "priority" in ref
            assert "builtin" in ref
            assert "broken" in ref

        builtin_refs = [r for r in refs if r["builtin"]]
        assert len(builtin_refs) > 0

    def test_add_scripting_runtime_to_word(self, test_document):
        """Test adding Scripting Runtime to Word document (common use case)."""
        manager = ReferenceManager(test_document)

        refs_before = manager.list_references()
        initial_count = len(refs_before)

        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        added = manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)
        assert added is True

        refs_after = manager.list_references()
        assert len(refs_after) == initial_count + 1

        scripting_ref = next((r for r in refs_after if r["guid"].upper() == scripting_guid.upper()), None)
        assert scripting_ref is not None
        assert scripting_ref["name"] == "Scripting"
        assert scripting_ref["builtin"] is False
        assert scripting_ref["broken"] is False

    def test_add_multiple_references_to_word(self, test_document):
        """Test adding multiple common references to Word document."""
        manager = ReferenceManager(test_document)

        common_refs = [
            {"guid": "{420B2830-E718-11CF-893D-00A0C9054228}", "name": "Scripting", "major": 1, "minor": 0},
            {"guid": "{00020905-0000-0000-C000-000000000046}", "name": "Word", "major": 8, "minor": 7},
        ]

        refs_before = manager.list_references()
        initial_count = len(refs_before)
        added_count = 0

        for ref_info in common_refs:
            added = manager.add_reference(**ref_info)
            if added:
                added_count += 1

        refs_after = manager.list_references()
        assert len(refs_after) >= initial_count + added_count

    def test_word_document_reference_persistence(self, word_app, tmp_path):
        """Test that references persist when saving and reopening Word document."""
        doc_path = tmp_path / "test_refs.docm"
        doc = word_app.Documents.Add()

        try:
            manager = ReferenceManager(doc)

            scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
            manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

            doc.SaveAs2(str(doc_path), FileFormat=13)  # wdFormatXMLDocumentMacroEnabled
            doc.Close()

            doc_reopened = word_app.Documents.Open(str(doc_path))
            manager_reopened = ReferenceManager(doc_reopened)

            assert manager_reopened.reference_exists(guid=scripting_guid) is True
            doc_reopened.Close(SaveChanges=False)

        finally:
            try:
                if doc_path.exists():
                    doc_path.unlink()
            except Exception:
                pass

    def test_export_word_references_to_toml(self, test_document):
        """Test exporting Word document references to TOML configuration."""
        manager = ReferenceManager(test_document)

        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

        with tempfile.TemporaryDirectory() as tmpdir:
            toml_file = Path(tmpdir) / "word_refs.toml"
            manager.export_to_toml(toml_file)

            assert toml_file.exists()

            content = toml_file.read_text(encoding="utf-8")
            assert "Scripting" in content
            assert scripting_guid in content
            assert "[[references]]" in content

            refs = manager.list_references()
            builtin_refs = [r for r in refs if r["builtin"]]

            for builtin_ref in builtin_refs:
                assert builtin_ref["guid"] not in content, (
                    f"Built-in reference {builtin_ref['name']} should be filtered from export"
                )

    def test_import_references_to_new_word_document(self, word_app, tmp_path):
        """Test importing references from TOML to a new Word document."""
        toml_content = """
# Common VBA references for Word documents

[[references]]
name = "Scripting"
guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
major = 1
minor = 0
description = "Microsoft Scripting Runtime (FileSystemObject, Dictionary)"

[[references]]
name = "ADODB"
guid = "{00000201-0000-0010-8000-00AA006D2EA4}"
major = 2
minor = 8
description = "Microsoft ActiveX Data Objects 2.8 Library"
"""
        toml_file = tmp_path / "word_common_refs.toml"
        toml_file.write_text(toml_content, encoding="utf-8")

        doc = word_app.Documents.Add()
        try:
            manager = ReferenceManager(doc)
            stats = manager.import_from_toml(toml_file)

            assert stats["added"] >= 0
            assert stats["failed"] == 0
            assert (stats["added"] + stats["skipped"]) >= 1
            assert manager.reference_exists(guid="{420B2830-E718-11CF-893D-00A0C9054228}") is True
        finally:
            doc.Close(SaveChanges=False)

    def test_word_reference_round_trip_workflow(self, word_app, tmp_path):
        """Test complete workflow: export from one doc, import to another."""
        source_doc = word_app.Documents.Add()
        source_manager = ReferenceManager(source_doc)

        refs_to_add = [
            {"guid": "{420B2830-E718-11CF-893D-00A0C9054228}", "name": "Scripting", "major": 1, "minor": 0},
            {"guid": "{00000201-0000-0010-8000-00AA006D2EA4}", "name": "ADODB", "major": 2, "minor": 8},
        ]

        for ref_info in refs_to_add:
            source_manager.add_reference(**ref_info)

        toml_file = tmp_path / "template_refs.toml"
        source_manager.export_to_toml(toml_file)
        source_doc.Close(SaveChanges=False)

        target_doc = word_app.Documents.Add()
        try:
            target_manager = ReferenceManager(target_doc)
            stats = target_manager.import_from_toml(toml_file)

            assert stats["added"] >= 0
            assert stats["failed"] == 0

            # Both references should be present
            for ref_info in refs_to_add:
                assert target_manager.reference_exists(guid=ref_info["guid"]) is True
        finally:
            target_doc.Close(SaveChanges=False)

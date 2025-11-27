"""
Integration tests for VBA Reference Management with Word.

These tests focus on Word-specific scenarios since Word is the primary use case
for the reference management feature (work requirement).

Run with: pytest tests/test_reference_manager_word_integration.py -v
"""

import pytest
import win32com.client
from pathlib import Path
import tempfile

from vba_edit.reference_manager import ReferenceManager, ReferenceError


@pytest.fixture(scope="module")
def word_app():
    """Create Word application for integration tests."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone
    yield word

    # Cleanup
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

        # Word always has some built-in references
        assert len(refs) > 0

        # Check structure
        for ref in refs:
            assert "name" in ref
            assert "guid" in ref
            assert "major" in ref
            assert "minor" in ref
            assert "priority" in ref
            assert "builtin" in ref
            assert "broken" in ref

        # Should have at least one built-in reference (VBA, Word)
        builtin_refs = [r for r in refs if r["builtin"]]
        assert len(builtin_refs) > 0

        # Print for debugging
        print("\nWord document references:")
        for ref in refs:
            status = "⚠ BROKEN" if ref["broken"] else "✓"
            builtin = " [BUILTIN]" if ref["builtin"] else ""
            print(f"{status} {ref['name']} v{ref['major']}.{ref['minor']}{builtin}")
            print(f"   GUID: {ref['guid']}")

    def test_add_scripting_runtime_to_word(self, test_document):
        """Test adding Scripting Runtime to Word document (common use case)."""
        manager = ReferenceManager(test_document)

        # Get initial count
        refs_before = manager.list_references()
        initial_count = len(refs_before)

        # Add Scripting Runtime (commonly used for FileSystemObject, Dictionary)
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        added = manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

        assert added is True

        # Verify it was added
        refs_after = manager.list_references()
        assert len(refs_after) == initial_count + 1

        # Find the added reference
        scripting_ref = next((r for r in refs_after if r["guid"].upper() == scripting_guid.upper()), None)
        assert scripting_ref is not None
        assert scripting_ref["name"] == "Scripting"
        assert scripting_ref["builtin"] is False
        assert scripting_ref["broken"] is False

    def test_add_multiple_references_to_word(self, test_document):
        """Test adding multiple common references to Word document."""
        manager = ReferenceManager(test_document)

        # Common references used in Word VBA projects
        common_refs = [
            {
                "guid": "{420B2830-E718-11CF-893D-00A0C9054228}",
                "name": "Scripting",
                "major": 1,
                "minor": 0,
                "description": "Microsoft Scripting Runtime",
            },
            {
                "guid": "{00020905-0000-0000-C000-000000000046}",
                "name": "Word",
                "major": 8,
                "minor": 7,
                "description": "Microsoft Word Object Library",
            },
        ]

        refs_before = manager.list_references()
        initial_count = len(refs_before)
        added_count = 0

        for ref_info in common_refs:
            added = manager.add_reference(
                guid=ref_info["guid"], name=ref_info["name"], major=ref_info["major"], minor=ref_info["minor"]
            )
            if added:
                added_count += 1

        refs_after = manager.list_references()
        assert len(refs_after) >= initial_count + added_count

    def test_word_document_reference_persistence(self, word_app, tmp_path):
        """Test that references persist when saving and reopening Word document."""
        # Create document with reference
        doc_path = tmp_path / "test_refs.docm"
        doc = word_app.Documents.Add()

        try:
            manager = ReferenceManager(doc)

            # Add reference
            scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
            manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

            # Save as macro-enabled document
            doc.SaveAs2(str(doc_path), FileFormat=13)  # wdFormatXMLDocumentMacroEnabled
            doc.Close()

            # Reopen and verify reference still exists
            doc_reopened = word_app.Documents.Open(str(doc_path))
            manager_reopened = ReferenceManager(doc_reopened)

            assert manager_reopened.reference_exists(guid=scripting_guid) is True

            refs = manager_reopened.list_references()
            scripting_ref = next((r for r in refs if r["guid"].upper() == scripting_guid.upper()), None)
            assert scripting_ref is not None

            doc_reopened.Close(SaveChanges=False)

        finally:
            # Cleanup
            try:
                if doc_path.exists():
                    doc_path.unlink()
            except Exception:
                pass

    def test_export_word_references_to_toml(self, test_document):
        """Test exporting Word document references to TOML configuration."""
        manager = ReferenceManager(test_document)

        # Add some references
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

        # Export to TOML
        with tempfile.TemporaryDirectory() as tmpdir:
            toml_file = Path(tmpdir) / "word_refs.toml"
            manager.export_to_toml(toml_file)

            assert toml_file.exists()

            # Read and verify content
            content = toml_file.read_text(encoding="utf-8")
            print("\nExported TOML content:")
            print(content)

            assert "Scripting" in content
            assert scripting_guid in content
            assert "[[references]]" in content

            # Built-in Word references should NOT be in the export
            refs = manager.list_references()
            builtin_refs = [r for r in refs if r["builtin"]]

            for builtin_ref in builtin_refs:
                # GUID should not appear in exported TOML
                assert builtin_ref["guid"] not in content, (
                    f"Built-in reference {builtin_ref['name']} should be filtered from export"
                )

    def test_import_references_to_new_word_document(self, word_app, tmp_path):
        """Test importing references from TOML to a new Word document."""
        # Create TOML with common Word references
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

        # Create new document and import
        doc = word_app.Documents.Add()

        try:
            manager = ReferenceManager(doc)

            # Import references
            stats = manager.import_from_toml(toml_file)

            print(f"\nImport statistics: {stats}")

            # Should have added references (or skipped if already present)
            assert stats["added"] >= 0
            assert stats["failed"] == 0  # Should not fail for valid references

            # Verify at least one was added or exists
            assert (stats["added"] + stats["skipped"]) >= 1

            # Verify Scripting Runtime exists
            assert manager.reference_exists(guid="{420B2830-E718-11CF-893D-00A0C9054228}") is True

        finally:
            doc.Close(SaveChanges=False)

    def test_word_reference_round_trip_workflow(self, word_app, tmp_path):
        """Test complete workflow: export from one doc, import to another (work use case)."""
        # Scenario: User has a Word template with references, wants to apply to new documents

        # Step 1: Create source document with references
        source_doc = word_app.Documents.Add()
        source_manager = ReferenceManager(source_doc)

        # Add references to source
        refs_to_add = [
            {"guid": "{420B2830-E718-11CF-893D-00A0C9054228}", "name": "Scripting", "major": 1, "minor": 0},
            {"guid": "{00000201-0000-0010-8000-00AA006D2EA4}", "name": "ADODB", "major": 2, "minor": 8},
        ]

        for ref_info in refs_to_add:
            source_manager.add_reference(**ref_info)

        # Step 2: Export to TOML
        toml_file = tmp_path / "template_refs.toml"
        source_manager.export_to_toml(toml_file)
        source_doc.Close(SaveChanges=False)

        # Step 3: Create new document and import
        target_doc = word_app.Documents.Add()
        target_manager = ReferenceManager(target_doc)

        # Before import - should not have these refs
        assert target_manager.reference_exists(guid=refs_to_add[0]["guid"]) is False

        # Import
        stats = target_manager.import_from_toml(toml_file)

        print(f"\nRound-trip import stats: {stats}")

        # After import - should have the refs
        for ref_info in refs_to_add:
            exists = target_manager.reference_exists(guid=ref_info["guid"])
            assert exists is True, f"Reference {ref_info['name']} should exist after import"

        target_doc.Close(SaveChanges=False)

    def test_word_builtin_references_cannot_be_removed(self, test_document):
        """Test that Word built-in references are protected from removal."""
        manager = ReferenceManager(test_document)

        # Get Word's built-in references
        refs = manager.list_references()
        builtin_refs = [r for r in refs if r["builtin"]]

        assert len(builtin_refs) > 0, "Word should have built-in references"

        # Try to remove a built-in reference
        builtin_ref = builtin_refs[0]

        with pytest.raises(ReferenceError, match="Cannot remove built-in reference"):
            manager.remove_reference(guid=builtin_ref["guid"])

        # Verify it still exists
        assert manager.reference_exists(guid=builtin_ref["guid"]) is True

    def test_word_reference_manager_with_template(self, word_app, tmp_path):
        """Test using reference manager with Word template (.dotm)."""
        template_path = tmp_path / "test_template.dotm"

        # Create template
        doc = word_app.Documents.Add()
        manager = ReferenceManager(doc)

        # Add reference
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

        # Save as template
        try:
            doc.SaveAs2(str(template_path), FileFormat=15)  # wdFormatXMLTemplateMacroEnabled
            doc.Close()

            # Open template and verify reference exists
            template = word_app.Documents.Open(str(template_path))
            template_manager = ReferenceManager(template)

            assert template_manager.reference_exists(guid=scripting_guid) is True

            template.Close(SaveChanges=False)

        finally:
            # Cleanup
            try:
                if template_path.exists():
                    template_path.unlink()
            except Exception:
                pass

    def test_word_broken_reference_detection(self, test_document):
        """Test that broken references are properly detected in Word."""
        manager = ReferenceManager(test_document)

        # Add a valid reference first
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        manager.add_reference(guid=scripting_guid, name="Scripting", major=1, minor=0)

        # List references and check none are broken
        refs = manager.list_references()
        scripting_ref = next((r for r in refs if r["guid"].upper() == scripting_guid.upper()), None)

        assert scripting_ref is not None
        assert scripting_ref["broken"] is False, "Scripting Runtime should not be broken"

        # All references in a fresh document should be unbroken
        broken_refs = [r for r in refs if r["broken"]]
        assert len(broken_refs) == 0, "Fresh Word document should have no broken references"

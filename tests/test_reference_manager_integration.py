"""
Integration tests for VBA Reference Management.

These tests require Excel to be installed and VBA trust enabled.
They test the ReferenceManager with real Excel COM objects.

Run with: pytest tests/test_reference_manager_integration.py -v
Skip with: pytest tests/ -v (integration tests auto-skip without explicit run)
"""

import pytest
import win32com.client
from pathlib import Path
import tempfile

from vba_edit.reference_manager import ReferenceManager, ReferenceError


@pytest.fixture(scope="module")
def excel_app():
    """Create Excel application for integration tests."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    yield excel
    
    # Cleanup
    try:
        for i in range(excel.Workbooks.Count, 0, -1):
            try:
                excel.Workbooks(i).Saved = True
                excel.Workbooks(i).Close(SaveChanges=False)
            except:
                pass
        excel.Quit()
    except:
        pass


@pytest.fixture
def test_workbook(excel_app):
    """Create a test workbook for each test."""
    wb = excel_app.Workbooks.Add()
    yield wb
    
    try:
        wb.Saved = True
        wb.Close(SaveChanges=False)
    except:
        pass


@pytest.mark.integration
@pytest.mark.skip(reason="Integration test - run explicitly with pytest tests/test_reference_manager_integration.py")
class TestReferenceManagerIntegration:
    """Integration tests for ReferenceManager with real Excel."""
    
    def test_list_references_real_excel(self, test_workbook):
        """Test listing references with real Excel workbook."""
        manager = ReferenceManager(test_workbook)
        refs = manager.list_references()
        
        # Excel always has some built-in references
        assert len(refs) > 0
        
        # Check structure of first reference
        first_ref = refs[0]
        assert 'name' in first_ref
        assert 'guid' in first_ref
        assert 'major' in first_ref
        assert 'minor' in first_ref
        assert 'priority' in first_ref
        assert 'builtin' in first_ref
        assert 'broken' in first_ref
        
        # Should have at least one built-in reference
        builtin_refs = [r for r in refs if r['builtin']]
        assert len(builtin_refs) > 0
        
        # None should be broken in a fresh workbook
        broken_refs = [r for r in refs if r['broken']]
        assert len(broken_refs) == 0
    
    def test_add_and_remove_reference_real_excel(self, test_workbook):
        """Test adding and removing a reference with real Excel."""
        manager = ReferenceManager(test_workbook)
        
        # Get initial count
        refs_before = manager.list_references()
        initial_count = len(refs_before)
        
        # Add Scripting Runtime
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        added = manager.add_reference(
            guid=scripting_guid,
            name="Scripting",
            major=1,
            minor=0
        )
        
        assert added is True
        
        # Verify it was added
        refs_after_add = manager.list_references()
        assert len(refs_after_add) == initial_count + 1
        
        # Find the added reference
        scripting_ref = next((r for r in refs_after_add if r['guid'].upper() == scripting_guid.upper()), None)
        assert scripting_ref is not None
        assert scripting_ref['name'] == "Scripting"
        assert scripting_ref['builtin'] is False
        
        # Remove it
        removed = manager.remove_reference(guid=scripting_guid)
        assert removed is True
        
        # Verify it was removed
        refs_after_remove = manager.list_references()
        assert len(refs_after_remove) == initial_count
        
        # Verify it's gone
        scripting_ref = next((r for r in refs_after_remove if r['guid'].upper() == scripting_guid.upper()), None)
        assert scripting_ref is None
    
    def test_add_duplicate_reference_skips(self, test_workbook):
        """Test that adding duplicate reference is skipped."""
        manager = ReferenceManager(test_workbook)
        
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        
        # Add first time
        added_first = manager.add_reference(
            guid=scripting_guid,
            name="Scripting",
            major=1,
            minor=0
        )
        assert added_first is True
        
        # Add second time (should skip)
        added_second = manager.add_reference(
            guid=scripting_guid,
            name="Scripting",
            major=1,
            minor=0,
            skip_if_exists=True
        )
        assert added_second is False  # Should return False (skipped)
    
    def test_reference_exists_real_excel(self, test_workbook):
        """Test checking if reference exists with real Excel."""
        manager = ReferenceManager(test_workbook)
        
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        
        # Should not exist initially
        assert manager.reference_exists(guid=scripting_guid) is False
        
        # Add it
        manager.add_reference(
            guid=scripting_guid,
            name="Scripting",
            major=1,
            minor=0
        )
        
        # Should exist now
        assert manager.reference_exists(guid=scripting_guid) is True
        assert manager.reference_exists(name="Scripting") is True
    
    def test_cannot_remove_builtin_reference(self, test_workbook):
        """Test that built-in references cannot be removed."""
        manager = ReferenceManager(test_workbook)
        
        # Get a built-in reference
        refs = manager.list_references()
        builtin_ref = next((r for r in refs if r['builtin']), None)
        assert builtin_ref is not None
        
        # Try to remove it
        with pytest.raises(ReferenceError, match="Cannot remove built-in reference"):
            manager.remove_reference(guid=builtin_ref['guid'])
    
    def test_export_import_toml_roundtrip(self, test_workbook):
        """Test exporting and importing references via TOML."""
        manager = ReferenceManager(test_workbook)
        
        # Add some test references
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        manager.add_reference(
            guid=scripting_guid,
            name="Scripting",
            major=1,
            minor=0
        )
        
        # Export to TOML
        with tempfile.TemporaryDirectory() as tmpdir:
            toml_file = Path(tmpdir) / "refs.toml"
            manager.export_to_toml(toml_file)
            
            assert toml_file.exists()
            
            # Read and verify content
            content = toml_file.read_text(encoding='utf-8')
            assert "Scripting" in content
            assert scripting_guid in content
            assert "[[references]]" in content
            
            # Remove the reference
            manager.remove_reference(guid=scripting_guid)
            assert manager.reference_exists(guid=scripting_guid) is False
            
            # Import it back
            stats = manager.import_from_toml(toml_file)
            
            assert stats['added'] == 1
            assert stats['skipped'] == 0
            assert stats['failed'] == 0
            
            # Verify it's back
            assert manager.reference_exists(guid=scripting_guid) is True
    
    def test_export_filters_builtin_references(self, test_workbook):
        """Test that export excludes built-in references."""
        manager = ReferenceManager(test_workbook)
        
        # Add a user reference
        scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        manager.add_reference(
            guid=scripting_guid,
            name="Scripting",
            major=1,
            minor=0
        )
        
        # Export
        with tempfile.TemporaryDirectory() as tmpdir:
            toml_file = Path(tmpdir) / "refs.toml"
            manager.export_to_toml(toml_file)
            
            content = toml_file.read_text(encoding='utf-8')
            
            # Should have Scripting
            assert "Scripting" in content
            
            # Get built-in references
            refs = manager.list_references()
            builtin_refs = [r for r in refs if r['builtin']]
            
            # Built-in references should NOT be in exported file
            for builtin_ref in builtin_refs:
                # Check that the GUID is not in the file
                assert builtin_ref['guid'] not in content
    
    def test_invalid_guid_raises_error(self, test_workbook):
        """Test that invalid GUID raises appropriate error."""
        manager = ReferenceManager(test_workbook)
        
        # Try to add with invalid GUID
        with pytest.raises(ValueError, match="Invalid GUID format"):
            manager.add_reference(
                guid="not-a-valid-guid",
                name="Invalid",
                major=1,
                minor=0
            )
    
    def test_remove_nonexistent_reference_with_skip(self, test_workbook):
        """Test removing non-existent reference with skip_if_missing."""
        manager = ReferenceManager(test_workbook)
        
        fake_guid = "{00000000-0000-0000-0000-000000000000}"
        
        # Should return False (skipped) not raise error
        removed = manager.remove_reference(guid=fake_guid, skip_if_missing=True)
        assert removed is False
    
    def test_remove_nonexistent_reference_without_skip(self, test_workbook):
        """Test removing non-existent reference without skip_if_missing."""
        manager = ReferenceManager(test_workbook)
        
        fake_guid = "{00000000-0000-0000-0000-000000000000}"
        
        # Should raise error
        with pytest.raises(ReferenceError, match="Reference not found"):
            manager.remove_reference(guid=fake_guid, skip_if_missing=False)

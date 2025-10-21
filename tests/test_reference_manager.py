"""
Tests for VBA reference management functionality (Phase 1: Core).

NOTE: These tests require Excel to be installed and VBA trust enabled.
Many tests are marked as skip by default - run with pytest -v -k "not skip" or individually.
"""

import pytest
from pathlib import Path

# NOTE: xlwings is optional dependency - tests will be skipped if not available
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

from vba_edit.reference_manager import ReferenceManager, ReferenceError


@pytest.fixture
def test_workbook():
    """Create a temporary test workbook."""
    if not XLWINGS_AVAILABLE:
        pytest.skip("xlwings not available")
    
    app = xw.App(visible=False)
    wb = app.books.add()
    wb.api.VBProject.References.AddFromGuid(
        "{00020813-0000-0000-C000-000000000046}",  # Excel type library
        1, 9  # Version 1.9
    )
    yield wb
    wb.close()
    app.quit()


@pytest.mark.skipif(not XLWINGS_AVAILABLE, reason="xlwings not available")
@pytest.mark.skipif(not XLWINGS_AVAILABLE, reason="xlwings not available")
@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_list_references(test_workbook):
    """Test listing all references in a workbook."""
    from vba_edit.reference_manager import ReferenceManager
    
    manager = ReferenceManager(test_workbook)
    refs = manager.list_references()
    
    assert isinstance(refs, list)
    assert len(refs) > 0
    
    # Check structure of first reference
    first_ref = refs[0]
    assert "name" in first_ref
    assert "guid" in first_ref
    assert "major" in first_ref
    assert "minor" in first_ref
    assert "priority" in first_ref
    assert "builtin" in first_ref
    assert "broken" in first_ref


@pytest.mark.skipif(not XLWINGS_AVAILABLE, reason="xlwings not available")
@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_add_reference(test_workbook):
    """Test adding a new reference."""
    from vba_edit.reference_manager import ReferenceManager
    
    manager = ReferenceManager(test_workbook)
    refs_before = manager.list_references()
    
    # Add Scripting Runtime
    manager.add_reference(
        guid="{420B2830-E718-11CF-893D-00A0C9054228}",
        name="Scripting",
        major=1,
        minor=0
    )
    
    refs_after = manager.list_references()
    assert len(refs_after) == len(refs_before) + 1
    
    # Verify the added reference
    scripting_ref = next((r for r in refs_after if r["name"] == "Scripting"), None)
    assert scripting_ref is not None
    assert scripting_ref["guid"] == "{420B2830-E718-11CF-893D-00A0C9054228}"


@pytest.mark.skipif(not XLWINGS_AVAILABLE, reason="xlwings not available")
@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_remove_reference(test_workbook):
    """Test removing a reference."""
    from vba_edit.reference_manager import ReferenceManager
    
    manager = ReferenceManager(test_workbook)
    
    # Add a reference first
    manager.add_reference(
        guid="{420B2830-E718-11CF-893D-00A0C9054228}",
        name="Scripting",
        major=1,
        minor=0
    )
    
    refs_before = manager.list_references()
    
    # Remove it
    manager.remove_reference(guid="{420B2830-E718-11CF-893D-00A0C9054228}")
    
    refs_after = manager.list_references()
    assert len(refs_after) == len(refs_before) - 1
    
    # Verify it's gone
    scripting_ref = next((r for r in refs_after if r["name"] == "Scripting"), None)
    assert scripting_ref is None


@pytest.mark.skipif(not XLWINGS_AVAILABLE, reason="xlwings not available")
@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_reference_exists(test_workbook):
    """Test checking if a reference exists."""
    from vba_edit.reference_manager import ReferenceManager
    
    manager = ReferenceManager(test_workbook)
    
    # Excel type library should exist (added in fixture)
    assert manager.reference_exists(guid="{00020813-0000-0000-C000-000000000046}") is True
    
    # Non-existent reference
    assert manager.reference_exists(guid="{00000000-0000-0000-0000-000000000000}") is False


@pytest.mark.skipif(not XLWINGS_AVAILABLE, reason="xlwings not available")
@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_export_references_to_toml(test_workbook, tmp_path):
    """Test exporting references to TOML file."""
    from vba_edit.reference_manager import ReferenceManager
    
    manager = ReferenceManager(test_workbook)
    output_file = tmp_path / "references.toml"
    
    manager.export_to_toml(output_file)
    
    assert output_file.exists()
    
    # Read and verify TOML structure
    try:
        import tomli
    except ImportError:
        pytest.skip("tomli not available")
    
    with open(output_file, "rb") as f:
        data = tomli.load(f)
    
    assert "references" in data
    assert isinstance(data["references"], list)
    assert len(data["references"]) > 0


@pytest.mark.skipif(not XLWINGS_AVAILABLE, reason="xlwings not available")
@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_import_references_from_toml(test_workbook, tmp_path):
    """Test importing references from TOML file."""
    from vba_edit.reference_manager import ReferenceManager
    
    # Create a TOML file with test references
    toml_content = """
[[references]]
name = "Scripting"
guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
major = 1
minor = 0
description = "Microsoft Scripting Runtime"
"""
    
    toml_file = tmp_path / "test_references.toml"
    toml_file.write_text(toml_content)
    
    manager = ReferenceManager(test_workbook)
    refs_before = manager.list_references()
    
    manager.import_from_toml(toml_file)
    
    refs_after = manager.list_references()
    assert len(refs_after) > len(refs_before)
    
    # Verify Scripting Runtime was added
    scripting_ref = next((r for r in refs_after if r["name"] == "Scripting"), None)
    assert scripting_ref is not None


def test_reference_priority_validation():
    """Test that priority values are validated correctly."""
    # This test doesn't require Excel - just validates logic
    # Implement once ReferenceManager is created
    pass


def test_guid_format_validation():
    """Test that GUID format is validated."""
    # This test doesn't require Excel - just validates logic
    # Implement once ReferenceManager is created
    pass

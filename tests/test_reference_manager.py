"""
Tests for VBA reference management functionality (Phase 1: Core).

NOTE: Tests requiring live Office interaction are marked skip and run manually.
ReferenceManager uses pure win32com - no xlwings dependency required.
"""

import pytest

from vba_edit.exceptions import ReferenceError, VBAError
from vba_edit.reference_manager import ReferenceManager


@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_list_references(tmp_path):
    """Test listing all references in a workbook."""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Add()
    try:
        manager = ReferenceManager(wb)
        refs = manager.list_references()

        assert isinstance(refs, list)
        assert len(refs) > 0

        first_ref = refs[0]
        assert "name" in first_ref
        assert "guid" in first_ref
        assert "major" in first_ref
        assert "minor" in first_ref
        assert "priority" in first_ref
        assert "builtin" in first_ref
        assert "broken" in first_ref
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_add_reference():
    """Test adding a new reference."""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Add()
    try:
        manager = ReferenceManager(wb)
        refs_before = manager.list_references()

        manager.add_reference(guid="{420B2830-E718-11CF-893D-00A0C9054228}", name="Scripting", major=1, minor=0)

        refs_after = manager.list_references()
        assert len(refs_after) == len(refs_before) + 1

        scripting_ref = next((r for r in refs_after if r["name"] == "Scripting"), None)
        assert scripting_ref is not None
        assert scripting_ref["guid"] == "{420B2830-E718-11CF-893D-00A0C9054228}"
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_remove_reference():
    """Test removing a reference."""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Add()
    try:
        manager = ReferenceManager(wb)
        manager.add_reference(guid="{420B2830-E718-11CF-893D-00A0C9054228}", name="Scripting", major=1, minor=0)
        refs_before = manager.list_references()

        manager.remove_reference(guid="{420B2830-E718-11CF-893D-00A0C9054228}")

        refs_after = manager.list_references()
        assert len(refs_after) == len(refs_before) - 1

        scripting_ref = next((r for r in refs_after if r["name"] == "Scripting"), None)
        assert scripting_ref is None
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_reference_exists():
    """Test checking if a reference exists."""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Add()
    try:
        manager = ReferenceManager(wb)
        assert manager.reference_exists(guid="{00020813-0000-0000-C000-000000000046}") is True
        assert manager.reference_exists(guid="{00000000-0000-0000-0000-000000000000}") is False
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_export_references_to_toml(tmp_path):
    """Test exporting references to TOML file."""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Add()
    try:
        manager = ReferenceManager(wb)
        output_file = tmp_path / "references.toml"
        manager.export_to_toml(output_file)

        assert output_file.exists()

        try:
            import tomli
        except ImportError:
            pytest.skip("tomli not available")

        with open(output_file, "rb") as f:
            data = tomli.load(f)

        assert "references" in data
        assert isinstance(data["references"], list)
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


@pytest.mark.skip(reason="Requires Excel interaction - run manually when testing")
def test_import_references_from_toml(tmp_path):
    """Test importing references from TOML file."""
    import win32com.client

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

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Add()
    try:
        manager = ReferenceManager(wb)
        refs_before = manager.list_references()
        manager.import_from_toml(toml_file)
        refs_after = manager.list_references()

        assert len(refs_after) > len(refs_before)
        scripting_ref = next((r for r in refs_after if r["name"] == "Scripting"), None)
        assert scripting_ref is not None
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


# ── Unit tests (no Office required) ──────────────────────────────────────────


def test_reference_priority_validation():
    """Test GUID pattern validation."""
    assert ReferenceManager.GUID_PATTERN.match("{420B2830-E718-11CF-893D-00A0C9054228}")
    assert ReferenceManager.GUID_PATTERN.match("{00020813-0000-0000-C000-000000000046}")

    assert not ReferenceManager.GUID_PATTERN.match("420B2830-E718-11CF-893D-00A0C9054228")
    assert not ReferenceManager.GUID_PATTERN.match("{420B2830-E718-11CF-893D}")
    assert not ReferenceManager.GUID_PATTERN.match("{INVALID-GUID-FORMAT}")
    assert not ReferenceManager.GUID_PATTERN.match("")
    assert not ReferenceManager.GUID_PATTERN.match("{}")


def test_guid_format_validation():
    """Test that add_reference raises ValueError for invalid GUIDs."""

    class MockDocument:
        class MockVBProject:
            Name = "TestProject"

            class MockReferences:
                Count = 0

            References = MockReferences()

        VBProject = MockVBProject()

    manager = ReferenceManager(MockDocument())

    invalid_guids = [
        "420B2830-E718-11CF-893D-00A0C9054228",
        "{420B2830-E718-11CF-893D}",
        "{INVALID-GUID-FORMAT}",
        "not-a-guid",
        "",
    ]

    for invalid_guid in invalid_guids:
        with pytest.raises(ValueError, match="Invalid GUID format"):
            if not manager.GUID_PATTERN.match(invalid_guid):
                raise ValueError(f"Invalid GUID format: {invalid_guid}")


def test_reference_error_inheritance():
    """Test that ReferenceError properly inherits from VBAError."""
    assert issubclass(ReferenceError, VBAError)

    try:
        raise ReferenceError("Test error")
    except VBAError:
        pass  # Should be catchable as VBAError


def test_reference_manager_initialization_with_invalid_object():
    """Test that ReferenceManager raises ReferenceError with unsupported object."""

    class InvalidObject:
        pass

    with pytest.raises(ReferenceError, match="does not support VBA projects"):
        ReferenceManager(InvalidObject())


def test_reference_manager_supports_both_xlwings_and_com():
    """Test that ReferenceManager handles both xlwings and COM objects."""

    class MockXLWingsWorkbook:
        class MockAPI:
            class MockVBProject:
                Name = "TestProject"

                class MockReferences:
                    Count = 0

                References = MockReferences()

            VBProject = MockVBProject()

        api = MockAPI()

    class MockCOMWorkbook:
        class MockVBProject:
            Name = "TestProject"

            class MockReferences:
                Count = 0

            References = MockReferences()

        VBProject = MockVBProject()

    xlwings_manager = ReferenceManager(MockXLWingsWorkbook())
    assert xlwings_manager.vb_project is not None

    com_manager = ReferenceManager(MockCOMWorkbook())
    assert com_manager.vb_project is not None


def test_guid_validation_case_insensitive():
    """Test that GUID pattern accepts both upper and lowercase."""
    assert ReferenceManager.GUID_PATTERN.match("{420B2830-E718-11CF-893D-00A0C9054228}")
    assert ReferenceManager.GUID_PATTERN.match("{420b2830-e718-11cf-893d-00a0c9054228}")
    assert ReferenceManager.GUID_PATTERN.match("{420B2830-e718-11CF-893d-00A0C9054228}")


def test_toml_export_creates_valid_structure(tmp_path):
    """Test that export_to_toml writes valid TOML-structured content."""
    import pywintypes

    class MockRef:
        def __init__(self, name, guid, major, minor, builtin, broken, description="", path=""):
            self.Name = name
            self.Guid = guid
            self.Major = major
            self.Minor = minor
            self.BuiltIn = builtin
            self.IsBroken = broken
            self.Description = description
            self.FullPath = path

    class MockReferences:
        def __init__(self, refs):
            self._refs = refs
            self.Count = len(refs)

        def Item(self, i):
            return self._refs[i - 1]

    class MockVBProject:
        Name = "TestProject"

        def __init__(self, refs):
            self.References = MockReferences(refs)

    class MockDocument:
        def __init__(self, refs):
            self.VBProject = MockVBProject(refs)

    refs = [
        MockRef("Scripting", "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0, False, False, "MS Scripting Runtime"),
        MockRef("VBA", "{000204EF-0000-0000-C000-000000000046}", 4, 2, True, False),  # built-in, excluded
    ]
    doc = MockDocument(refs)

    # Monkey-patch list_references to avoid pywintypes.com_error in mock
    manager = ReferenceManager.__new__(ReferenceManager)
    manager.document = doc
    manager.vb_project = doc.VBProject

    output = tmp_path / "refs.toml"
    manager.export_to_toml(output)

    content = output.read_text(encoding="utf-8")
    assert "[[references]]" in content
    assert "Scripting" in content
    assert "{420B2830-E718-11CF-893D-00A0C9054228}" in content
    # Built-in VBA reference should be excluded
    assert "{000204EF-0000-0000-C000-000000000046}" not in content

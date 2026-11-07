"""Tests: UserForms are imported through a real Import in both header modes, never re-created."""

from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import ExcelVBAHandler

FORM_HEADER = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestForm
   Caption         =   "TestForm"
   ClientHeight    =   3015
   OleObjectBlob   =   "TestForm.frx":0000
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False"""

CODE = "Private Sub UserForm_Click()\nEnd Sub"


class FakeComponents:
    def __init__(self, *components):
        self.components = {c.Name: c for c in components}
        self.removed = []
        self.imported = []  # (frm text, frx bytes or None)
        self.added = []

    def __call__(self, name):
        return self.components[name]

    def Remove(self, component):
        self.removed.append(component.Name)

    def Import(self, path):
        frm = Path(path)
        frx = frm.with_suffix(".frx")
        self.imported.append((frm.read_text(encoding="cp1252"), frx.read_bytes() if frx.exists() else None))

    def Add(self, kind):
        component = Mock(Type=kind)
        self.added.append(component)
        return component


@pytest.fixture
def handler(tmp_path):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        h = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"), save_headers=True)
    h.vba_dir.mkdir(exist_ok=True)
    h.doc = Mock(FullName=str(workbook))
    return h


def write_form(handler, with_header=True, with_frx=True):
    frm = handler.vba_dir / "TestForm.frm"
    frm.write_text(CODE + "\n", encoding="cp1252")
    if with_header:
        (handler.vba_dir / "TestForm.header").write_text(FORM_HEADER + "\n", encoding="cp1252")
    if with_frx:
        (handler.vba_dir / "TestForm.frx").write_bytes(b"\x00\x01binary")
    return frm


def test_existing_form_is_replaced_through_import_with_header_and_frx(handler):
    existing = Mock(Name="TestForm", Type=3)
    components = FakeComponents(existing)

    handler.import_component(write_form(handler), components)

    assert components.removed == ["TestForm"]
    assert components.added == []
    ((frm_text, frx_bytes),) = components.imported
    assert frm_text.startswith("VERSION 5.00")
    assert 'Attribute VB_Name = "TestForm"' in frm_text and "UserForm_Click" in frm_text
    assert frx_bytes == b"\x00\x01binary"
    assert not any(p.name.endswith("_import_tmp") for p in handler.vba_dir.iterdir())  # temp folder removed


def test_new_form_is_imported_not_added(handler):
    components = FakeComponents()

    handler.import_component(write_form(handler), components)

    assert components.added == [] and components.removed == []
    assert len(components.imported) == 1


def test_form_without_any_header_is_refused_and_left_alone(handler):
    existing = Mock(Name="TestForm", Type=3)
    components = FakeComponents(existing)

    with pytest.raises(VBAError, match="TestForm.frm"):
        handler.import_component(write_form(handler, with_header=False), components)

    assert components.removed == [] and components.added == [] and components.imported == []


def test_preparation_failure_leaves_the_existing_form_in_place(handler, monkeypatch):
    existing = Mock(Name="TestForm", Type=3)
    components = FakeComponents(existing)
    write_form(handler)

    def broken_copy(*_args, **_kwargs):
        raise OSError("disk full")

    monkeypatch.setattr("vba_edit.office_vba.shutil.copy2", broken_copy)

    with pytest.raises(VBAError):
        handler.import_component(handler.vba_dir / "TestForm.frm", components)

    assert components.removed == []  # removal happens only after the files are prepared

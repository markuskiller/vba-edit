"""Tests for issue #122: class modules are removed and re-imported only when that is needed."""

from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from vba_edit.office_vba import (
    AccessVBAHandler,
    ExcelVBAHandler,
    _has_member_attributes,
    _normalized_header,
)

HEADER = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Widget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False"""

PREDECLARED_HEADER = HEADER.replace("VB_PredeclaredId = False", "VB_PredeclaredId = True")


class FakeCodeModule:
    def __init__(self, body=""):
        self.body = body

    @property
    def CountOfLines(self):
        return len(self.body.splitlines())

    def Lines(self, _start, _count):
        return self.body

    def DeleteLines(self, _start, _count):
        self.body = ""

    def AddFromString(self, code):
        self.body = code


class FakeComponent:
    """A class module whose Export writes its current header and code, like the VBE does."""

    def __init__(self, name, header, body, encoding="cp1252"):
        self.Name = name
        self.Type = 2
        self.header = header
        self.CodeModule = FakeCodeModule(body)
        self.encoding = encoding

    def Export(self, path):
        Path(path).write_text(f"{self.header}\n{self.CodeModule.body}\n", encoding=self.encoding)


class FakeComponents:
    def __init__(self, *components):
        self.components = {c.Name: c for c in components}
        self.removed = []
        self.imported = []

    def __call__(self, name):
        return self.components[name]

    def Remove(self, component):
        self.removed.append(component.Name)
        del self.components[component.Name]

    def Import(self, path):
        self.imported.append(Path(path).read_text(encoding="cp1252"))


@pytest.fixture
def handler(tmp_path):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        h = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"), in_file_headers=True)
    h.vba_dir.mkdir(exist_ok=True)
    return h


def write_class(handler, header, code, in_file=True):
    path = handler.vba_dir / "Widget.cls"
    if in_file:
        path.write_text(f"{header}\n{code}\n", encoding="cp1252")
    else:
        path.write_text(f"{code}\n", encoding="cp1252")
        if header:
            path.with_suffix(".header").write_text(header + "\n", encoding="cp1252")
    return path


# --- helpers -----------------------------------------------------------------------------------------


def test_member_attribute_detection():
    assert _has_member_attributes("Public Property Get Item()\nAttribute Item.VB_UserMemId = 0\nEnd Property")
    assert _has_member_attributes("Private WithEvents Btn As Button\nAttribute Btn.VB_VarHelpID = -1")
    assert not _has_member_attributes('Attribute VB_Name = "Widget"\nSub A()\nEnd Sub')
    assert not _has_member_attributes("")


def test_header_comparison_ignores_blank_lines_and_indentation():
    assert _normalized_header(HEADER) == _normalized_header("\n" + HEADER.replace("\n", "\n\n") + "  \n")
    assert _normalized_header(HEADER) != _normalized_header(PREDECLARED_HEADER)


# --- in-file headers ------------------------------------------------------------------------------------


def test_unchanged_header_updates_the_code_in_place(handler):
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(handler, HEADER, "Sub New()\nEnd Sub")

    handler.import_component(path, components)

    assert widget.CodeModule.body == "Sub New()\nEnd Sub"
    assert components.removed == [] and components.imported == []
    assert not list(handler.vba_dir.glob("*_current.tmp"))  # temp export cleaned up


def test_changed_header_removes_and_imports(handler):
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(handler, PREDECLARED_HEADER, "Sub New()\nEnd Sub")

    handler.import_component(path, components)

    assert components.removed == ["Widget"]
    assert len(components.imported) == 1
    assert "VB_PredeclaredId = True" in components.imported[0]
    assert "Sub New()" in components.imported[0]


def test_member_attributes_force_a_full_import_even_with_the_same_header(handler):
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(handler, HEADER, "Public Property Get Item()\nAttribute Item.VB_UserMemId = 0\nEnd Property")

    handler.import_component(path, components)

    assert components.removed == ["Widget"]
    assert "VB_UserMemId = 0" in components.imported[0]


def test_unreadable_current_header_keeps_the_old_behaviour(handler):
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    widget.Export = Mock(side_effect=RuntimeError("Export failed"))
    components = FakeComponents(widget)
    path = write_class(handler, HEADER, "Sub New()\nEnd Sub")

    handler.import_component(path, components)

    assert components.removed == ["Widget"]


def test_a_new_class_module_is_still_imported_with_its_header(handler):
    components = FakeComponents()
    path = write_class(handler, PREDECLARED_HEADER, "Sub New()\nEnd Sub")

    handler.import_component(path, components)

    assert components.removed == []
    assert "VB_PredeclaredId = True" in components.imported[0]


# --- separate header files ------------------------------------------------------------------------------


@pytest.fixture
def separate_handler(tmp_path):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        h = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"), save_headers=True)
    h.vba_dir.mkdir(exist_ok=True)
    return h


def test_separate_header_unchanged_updates_in_place(separate_handler):
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(separate_handler, HEADER, "Sub New()\nEnd Sub", in_file=False)

    separate_handler.import_component(path, components)

    assert widget.CodeModule.body == "Sub New()\nEnd Sub"
    assert components.removed == []


def test_separate_header_changed_is_applied_through_a_real_import(separate_handler):
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(separate_handler, PREDECLARED_HEADER, "Sub New()\nEnd Sub", in_file=False)

    separate_handler.import_component(path, components)

    assert components.removed == ["Widget"]
    assert "VB_PredeclaredId = True" in components.imported[0]
    assert "Sub New()" in components.imported[0]


def test_class_without_any_header_file_keeps_its_attributes(separate_handler):
    """No header to apply: the module is updated in place instead of being re-created with defaults."""
    widget = FakeComponent("Widget", PREDECLARED_HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(separate_handler, "", "Sub New()\nEnd Sub", in_file=False)

    separate_handler.import_component(path, components)

    assert widget.CodeModule.body == "Sub New()\nEnd Sub"
    assert components.removed == []


# --- Access saves the re-imported module ------------------------------------------------------------------


def test_access_saves_a_re_imported_module(tmp_path):
    database = tmp_path / "db.accdb"
    database.touch()
    with patch("vba_edit.office_vba.win32com.client.GetActiveObject", side_effect=RuntimeError("not running")):
        with patch("vba_edit.office_vba.win32com.client.Dispatch"):
            with patch("vba_edit.office_vba.AccessVBAHandler._open_document_impl"):
                handler = AccessVBAHandler(doc_path=str(database), vba_dir=str(tmp_path / "src"), in_file_headers=True)
    handler.vba_dir.mkdir(exist_ok=True)
    handler.app = Mock()
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(handler, PREDECLARED_HEADER, "Sub New()\nEnd Sub")

    handler.import_component(path, components)

    handler.app.DoCmd.Save.assert_called_once_with(5, "Widget")


def test_access_does_not_save_when_the_code_was_replaced_in_place(tmp_path):
    database = tmp_path / "db.accdb"
    database.touch()
    with patch("vba_edit.office_vba.win32com.client.GetActiveObject", side_effect=RuntimeError("not running")):
        with patch("vba_edit.office_vba.win32com.client.Dispatch"):
            with patch("vba_edit.office_vba.AccessVBAHandler._open_document_impl"):
                handler = AccessVBAHandler(doc_path=str(database), vba_dir=str(tmp_path / "src"), in_file_headers=True)
    handler.vba_dir.mkdir(exist_ok=True)
    handler.app = Mock()
    widget = FakeComponent("Widget", HEADER, "Sub Old()\nEnd Sub")
    components = FakeComponents(widget)
    path = write_class(handler, HEADER, "Sub New()\nEnd Sub")

    handler.import_component(path, components)

    handler.app.DoCmd.Save.assert_not_called()
    assert widget.CodeModule.body == "Sub New()\nEnd Sub"

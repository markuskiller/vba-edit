"""Tests: replacing a class module or UserForm never loses the original, and complete imports are complete.

A class module whose header changed and every UserForm are replaced by removing the
component and importing the file. Office applies the removal at once, so the original
is exported first and put back when the import fails. The same file also pins what a
complete import must carry: the attributes from a separate ``.header`` file for a new
class, and the Rubberduck ``@Folder`` annotation that belongs to the file's folder.
"""

import logging
import re
import shutil
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import ExcelVBAHandler

CLASS_HEADER = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Widget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False"""
PREDECLARED_HEADER = CLASS_HEADER.replace("VB_PredeclaredId = False", "VB_PredeclaredId = True")

FORM_HEADER = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog
   Caption         =   "Dialog"
   OleObjectBlob   =   "Dialog.frx":0000
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False"""


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
    """A component that exports like the VBE: header and code, plus a .frx for a UserForm."""

    def __init__(self, name, header, body, layout=None):
        self.Name = name
        self.Type = 3 if layout is not None else 2
        self.header = header
        self.layout = layout
        self.CodeModule = FakeCodeModule(body)

    def Export(self, path):
        Path(path).write_text(f"{self.header}\n{self.CodeModule.body}\n", encoding="cp1252")
        if self.layout is not None:
            Path(path).with_suffix(".frx").write_bytes(self.layout)


class FakeProject:
    """A VBComponents collection that really removes and imports, and can be told to fail."""

    def __init__(self, *components):
        self.components = {c.Name: c for c in components}
        self.fail_when = None  # callable(text) -> bool: Import raises for this file
        self.leave_behind = None  # name of a component a failed Import leaves in the project
        self.imported_paths = []

    def __call__(self, name):
        return self.components[name]

    def __iter__(self):
        return iter(list(self.components.values()))

    def Remove(self, component):
        del self.components[component.Name]

    def Add(self, kind):
        component = FakeComponent("Added", "", "")
        component.Type = kind
        self.components[f"added-{len(self.components)}"] = component
        return component

    def Import(self, path):
        path = Path(path)
        text = path.read_text(encoding="cp1252")
        self.imported_paths.append(path)
        if self.fail_when and self.fail_when(text):
            if self.leave_behind:
                self.components[self.leave_behind] = FakeComponent(self.leave_behind, "", "")
            raise RuntimeError("the VBA editor rejected the file")
        name = re.search(r'Attribute VB_Name = "([^"]+)"', text).group(1)
        boundary = text.rindex("Attribute VB_Exposed")
        boundary = text.index("\n", boundary)
        frx = path.with_suffix(".frx")
        layout = frx.read_bytes() if frx.exists() else (b"" if path.suffix == ".frm" else None)
        self.components[name] = FakeComponent(name, text[:boundary], text[boundary:].strip(), layout)


def make_handler(tmp_path, **options):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        handler = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"), **options)
    handler.vba_dir.mkdir(exist_ok=True)
    handler.office_encoding = "cp1252"
    handler.doc = Mock(FullName=str(workbook))
    return handler


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text + "\n", encoding="cp1252")
    return path


def safety_copies():
    return list(Path(tempfile.gettempdir()).glob("vba-edit-backup-*"))


@pytest.fixture(autouse=True)
def no_safety_copies_left_behind():
    before = set(safety_copies())
    yield
    for folder in set(safety_copies()) - before:
        shutil.rmtree(folder, ignore_errors=True)
        pytest.fail(f"safety copy was not cleaned up: {folder}")


# --- a failed replacement puts the original back -----------------------------------------------------


def test_class_with_in_file_header_is_restored_when_the_new_version_cannot_be_imported(tmp_path, caplog):
    handler = make_handler(tmp_path, in_file_headers=True)
    project = FakeProject(FakeComponent("Widget", CLASS_HEADER, "Sub Old()\nEnd Sub"))
    project.fail_when = lambda text: "Sub New()" in text
    path = write(handler.vba_dir / "Widget.cls", f"{PREDECLARED_HEADER}\nSub New()\nEnd Sub")

    with caplog.at_level(logging.WARNING), pytest.raises(VBAError):
        handler.import_component(path, project)

    widget = project("Widget")
    assert widget.CodeModule.body == "Sub Old()\nEnd Sub"
    assert "VB_PredeclaredId = False" in widget.header
    assert "the original Widget was put back" in caplog.text


def test_class_with_separate_header_is_restored_too(tmp_path):
    handler = make_handler(tmp_path, save_headers=True)
    project = FakeProject(FakeComponent("Widget", CLASS_HEADER, "Sub Old()\nEnd Sub"))
    project.fail_when = lambda text: "Sub New()" in text
    path = write(handler.vba_dir / "Widget.cls", "Sub New()\nEnd Sub")
    write(handler.vba_dir / "Widget.header", PREDECLARED_HEADER)

    with pytest.raises(VBAError):
        handler.import_component(path, project)

    assert project("Widget").CodeModule.body == "Sub Old()\nEnd Sub"


@pytest.mark.parametrize("options", [{"in_file_headers": True}, {"save_headers": True}])
def test_userform_is_restored_with_its_layout(tmp_path, options):
    handler = make_handler(tmp_path, **options)
    project = FakeProject(FakeComponent("Dialog", FORM_HEADER, "'old code", layout=b"old layout"))
    project.fail_when = lambda text: "'new code" in text
    if "in_file_headers" in options:
        path = write(handler.vba_dir / "Dialog.frm", f"{FORM_HEADER}\n'new code")
    else:
        path = write(handler.vba_dir / "Dialog.frm", "'new code")
        write(handler.vba_dir / "Dialog.header", FORM_HEADER)
    (handler.vba_dir / "Dialog.frx").write_bytes(b"new layout")

    with pytest.raises(VBAError):
        handler.import_component(path, project)

    dialog = project("Dialog")
    assert dialog.CodeModule.body == "'old code"
    assert dialog.layout == b"old layout"


def test_what_a_failed_import_left_behind_is_removed_before_restoring(tmp_path):
    handler = make_handler(tmp_path, in_file_headers=True)
    bystander = FakeComponent("Module1", "", "Sub Untouched()\nEnd Sub")
    project = FakeProject(FakeComponent("Dialog", FORM_HEADER, "'old code", layout=b"old layout"), bystander)
    project.fail_when = lambda text: "'new code" in text
    project.leave_behind = "UserForm1"  # what the VBE creates when it gives up on a .frm
    path = write(handler.vba_dir / "Dialog.frm", f"{FORM_HEADER}\n'new code")

    with pytest.raises(VBAError):
        handler.import_component(path, project)

    assert sorted(project.components) == ["Dialog", "Module1"]
    assert project("Module1") is bystander


def test_a_replacement_that_works_leaves_no_safety_copy(tmp_path):
    handler = make_handler(tmp_path, in_file_headers=True)
    project = FakeProject(FakeComponent("Widget", CLASS_HEADER, "Sub Old()\nEnd Sub"))
    path = write(handler.vba_dir / "Widget.cls", f"{PREDECLARED_HEADER}\nSub New()\nEnd Sub")

    handler.import_component(path, project)

    assert project("Widget").CodeModule.body == "Sub New()\nEnd Sub"
    assert "VB_PredeclaredId = True" in project("Widget").header


def test_when_the_original_cannot_be_put_back_the_message_says_where_the_copy_is(tmp_path, caplog):
    handler = make_handler(tmp_path, in_file_headers=True)
    project = FakeProject(FakeComponent("Widget", CLASS_HEADER, "Sub Old()\nEnd Sub"))
    project.fail_when = lambda text: True  # neither the new version nor the safety copy can be imported
    path = write(handler.vba_dir / "Widget.cls", f"{PREDECLARED_HEADER}\nSub New()\nEnd Sub")

    with caplog.at_level(logging.ERROR), pytest.raises(VBAError):
        handler.import_component(path, project)

    (copy,) = [folder for folder in safety_copies() if (folder / "Widget.cls").exists()]
    try:
        assert str(copy) in caplog.text
        assert "Sub Old()" in (copy / "Widget.cls").read_text(encoding="cp1252")
    finally:
        shutil.rmtree(copy)


def test_batch_import_saves_the_restored_original_next_to_what_succeeded(tmp_path):
    handler = make_handler(tmp_path, in_file_headers=True)
    project = FakeProject(
        FakeComponent("Widget", CLASS_HEADER, "Sub Old()\nEnd Sub"),
        FakeComponent("Dialog", FORM_HEADER, "'old code", layout=b"old layout"),
    )
    project.fail_when = lambda text: "'new code" in text
    write(handler.vba_dir / "Widget.cls", f"{PREDECLARED_HEADER}\nSub New()\nEnd Sub")
    write(handler.vba_dir / "Dialog.frm", f"{FORM_HEADER}\n'new code")
    handler.get_vba_project = lambda: Mock(VBComponents=project)
    handler.doc.Name = "Book.xlsm"

    with pytest.raises(VBAError, match="Dialog.frm"):
        handler.import_vba()

    handler.doc.Save.assert_called_once()
    assert project("Widget").CodeModule.body == "Sub New()\nEnd Sub"
    assert project("Dialog").layout == b"old layout"


# --- a new class with a separate header is a complete import -------------------------------------------


def test_new_class_gets_the_attributes_from_its_header_file(tmp_path):
    handler = make_handler(tmp_path, save_headers=True)
    project = FakeProject()
    path = write(handler.vba_dir / "Widget.cls", "Sub New()\nEnd Sub")
    write(handler.vba_dir / "Widget.header", PREDECLARED_HEADER)

    handler.import_component(path, project)

    assert sorted(project.components) == ["Widget"]  # imported, not added and filled in
    assert "VB_PredeclaredId = True" in project("Widget").header
    assert project("Widget").CodeModule.body == "Sub New()\nEnd Sub"


def test_new_class_without_a_header_file_is_still_created_directly(tmp_path):
    handler = make_handler(tmp_path)
    project = FakeProject()
    path = write(handler.vba_dir / "Widget.cls", "Sub New()\nEnd Sub")

    handler.import_component(path, project)

    assert project.imported_paths == []
    (added,) = project.components.values()
    assert added.Name == "Widget" and added.CodeModule.body == "Sub New()\nEnd Sub"


# --- the Rubberduck folder annotation reaches complete imports -------------------------------------------


def rubberduck_handler(tmp_path):
    return make_handler(tmp_path, in_file_headers=True, use_rubberduck_folders=True)


def test_new_class_in_a_folder_is_imported_with_that_folder(tmp_path):
    handler = rubberduck_handler(tmp_path)
    project = FakeProject()
    code = '\'@Folder("Old.Place")\nSub New()\nEnd Sub'
    path = write(handler.vba_dir / "Domain" / "Billing" / "Widget.cls", f"{CLASS_HEADER}\n{code}")

    handler.import_component(path, project)

    body = project("Widget").CodeModule.body
    assert '\'@Folder("Domain.Billing")' in body
    assert "Old.Place" not in body
    assert "VB_PredeclaredId = False" in project("Widget").header  # the header is still complete


def test_reimported_class_gets_the_folder_as_well(tmp_path):
    handler = rubberduck_handler(tmp_path)
    project = FakeProject(FakeComponent("Widget", CLASS_HEADER, "Sub Old()\nEnd Sub"))
    path = write(handler.vba_dir / "Domain" / "Widget.cls", f"{PREDECLARED_HEADER}\nSub New()\nEnd Sub")

    handler.import_component(path, project)

    assert project("Widget").CodeModule.body.startswith('\'@Folder("Domain")')


def test_userform_moved_to_another_folder_is_imported_with_the_new_folder_and_its_layout(tmp_path):
    handler = rubberduck_handler(tmp_path)
    project = FakeProject(FakeComponent("Dialog", FORM_HEADER, "'old code", layout=b"old layout"))
    path = write(handler.vba_dir / "Forms" / "Dialog.frm", f"{FORM_HEADER}\n'@Folder(\"Elsewhere\")\n'new code")
    path.with_suffix(".frx").write_bytes(b"new layout")

    handler.import_component(path, project)

    dialog = project("Dialog")
    assert dialog.CodeModule.body.startswith('\'@Folder("Forms")')
    assert dialog.layout == b"new layout"
    assert project.imported_paths[-1] != path  # a staged copy: the file itself still names the old folder
    assert "Elsewhere" in path.read_text(encoding="cp1252")  # and is not rewritten


def test_userform_whose_annotation_already_fits_is_imported_from_the_file_itself(tmp_path):
    handler = rubberduck_handler(tmp_path)
    project = FakeProject()
    path = write(handler.vba_dir / "Forms" / "Dialog.frm", f"{FORM_HEADER}\n'@Folder(\"Forms\")\n'new code")
    path.with_suffix(".frx").write_bytes(b"new layout")

    handler.import_component(path, project)

    assert project.imported_paths == [path]

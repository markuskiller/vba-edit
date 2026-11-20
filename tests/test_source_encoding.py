"""Tests: the encoding of the source files is separate from the code page of the VBA editor.

``VBComponent.Export`` writes, and ``VBComponents.Import`` reads, the Windows code page.
``--encoding`` names the encoding of the files in the VBA directory, so text is converted
between the two; ``--detect-encoding`` works the encoding out for each file.
"""

import json
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import ExcelVBAHandler

CODE = 'Sub Grüße()\n    MsgBox "Preis: 5 €"\nEnd Sub'
CLASS_HEADER = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Widget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False"""
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


class OfficeComponent:
    """Exports in the Windows code page, as the VBA editor does."""

    def __init__(self, name, kind, text, extension):
        self.Name = name
        self.Type = kind
        self.text = text
        self.extension = extension
        self.CodeModule = FakeCodeModule(text)

    def Export(self, path):
        Path(path).write_bytes((self.text + "\n").replace("\n", "\r\n").encode("cp1252"))


class OfficeProject:
    """Imports like the VBA editor: it reads the bytes of the file as Windows code page text."""

    def __init__(self, *components):
        self.components = {c.Name: c for c in components}
        self.imported = []  # (path, text as the VBA editor sees it)
        self.removed = []

    def __call__(self, name):
        return self.components[name]

    def __iter__(self):
        return iter(list(self.components.values()))

    def Remove(self, component):
        self.removed.append(component.Name)
        del self.components[component.Name]

    def Add(self, kind):
        component = OfficeComponent("Added", kind, "", ".bas")
        self.components[f"added-{len(self.components)}"] = component
        return component

    def Import(self, path):
        self.imported.append((Path(path), Path(path).read_bytes().decode("cp1252")))


def make_handler(tmp_path, encoding, **options):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        handler = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"), encoding=encoding, **options)
    handler.vba_dir.mkdir(exist_ok=True)
    handler.office_encoding = "cp1252"
    handler.doc = Mock(FullName=str(workbook))
    return handler


def module(name="Module1", text=CODE):
    return OfficeComponent(name, 1, f'Attribute VB_Name = "{name}"\n{text}', ".bas")


def info_for(component):
    return {
        "name": component.Name,
        "type": component.Type,
        "type_name": "Standard Module",
        "extension": component.extension,
        "code_lines": 3,
        "has_cls_header": False,
    }


def export(handler, component):
    with patch.object(handler.component_handler, "get_component_info", return_value=info_for(component)):
        return handler.export_component(component, handler.vba_dir)


# --- --encoding ---------------------------------------------------------------------------------------


def test_export_converts_the_code_page_of_the_vba_editor_to_the_encoding_of_the_files(tmp_path):
    handler = make_handler(tmp_path, "utf-8")

    used = export(handler, module())

    raw = (handler.vba_dir / "Module1.bas").read_bytes()
    assert used == "utf-8"
    assert "Grüße".encode("utf-8") in raw and "€".encode("utf-8") in raw
    assert "Grüße".encode("cp1252") not in raw


def test_default_encoding_writes_the_code_page_as_before(tmp_path):
    handler = make_handler(tmp_path, "cp1252")

    export(handler, module())

    assert "Grüße".encode("cp1252") in (handler.vba_dir / "Module1.bas").read_bytes()


def test_import_of_a_complete_class_hands_the_vba_editor_its_own_code_page(tmp_path):
    handler = make_handler(tmp_path, "utf-8", in_file_headers=True)
    project = OfficeProject()
    path = handler.vba_dir / "Widget.cls"
    path.write_text(f"{CLASS_HEADER}\n{CODE}\n", encoding="utf-8")

    handler.import_component(path, project)

    ((_path, seen),) = project.imported
    assert "Sub Grüße()" in seen and "5 €" in seen


def test_utf8_userform_is_staged_in_the_code_page_instead_of_imported_as_it_is(tmp_path):
    handler = make_handler(tmp_path, "utf-8", in_file_headers=True)
    project = OfficeProject()
    path = handler.vba_dir / "Dialog.frm"
    path.write_text(f"{FORM_HEADER}\n{CODE}\n", encoding="utf-8")
    path.with_suffix(".frx").write_bytes(b"layout")

    handler.import_component(path, project)

    ((imported_path, seen),) = project.imported
    assert imported_path != path
    assert "Sub Grüße()" in seen and "5 €" in seen


def test_userform_the_vba_editor_can_read_as_it_is_is_imported_directly(tmp_path):
    handler = make_handler(tmp_path, "utf-8", in_file_headers=True)
    project = OfficeProject()
    path = handler.vba_dir / "Dialog.frm"
    path.write_text(f"{FORM_HEADER}\nSub Plain()\nEnd Sub\n", encoding="utf-8")  # ASCII only

    handler.import_component(path, project)

    assert [imported_path for imported_path, _seen in project.imported] == [path]


def test_code_update_passes_unicode_text_to_office(tmp_path):
    handler = make_handler(tmp_path, "utf-8")
    existing = OfficeComponent("Module1", 1, "Sub Old()\nEnd Sub", ".bas")
    project = OfficeProject(existing)
    path = handler.vba_dir / "Module1.bas"
    path.write_text(CODE + "\n", encoding="utf-8")

    handler.import_component(path, project)

    assert existing.CodeModule.body == CODE


def test_character_outside_the_code_page_is_refused_before_anything_is_removed(tmp_path):
    handler = make_handler(tmp_path, "utf-8", in_file_headers=True)
    existing = OfficeComponent("Widget", 2, CLASS_HEADER.replace("= True", "= False") + "\nSub Old()\nEnd Sub", ".cls")
    project = OfficeProject(existing)
    path = handler.vba_dir / "Widget.cls"
    path.write_text(f"{CLASS_HEADER}\n' 日本語\nSub New()\nEnd Sub\n", encoding="utf-8")

    with pytest.raises(VBAError) as failure:
        handler.import_component(path, project)

    assert "cp1252" in str(failure.value.__cause__)
    assert project.removed == [] and project("Widget") is existing


# --- --detect-encoding ------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "file_encoding, expected",
    [("utf-8", "utf-8"), ("cp1252", "cp1252"), ("utf-8-sig", "utf-8-sig"), ("utf-16", "utf-16")],
)
def test_detection_reads_each_file_in_its_own_encoding(tmp_path, file_encoding, expected):
    handler = make_handler(tmp_path, None)
    with patch("vba_edit.utils.get_windows_ansi_codepage", return_value="cp1252"):
        existing = OfficeComponent("Module1", 1, "Sub Old()\nEnd Sub", ".bas")
        path = handler.vba_dir / "Module1.bas"
        path.write_text(CODE + "\n", encoding=file_encoding)

        assert handler._file_encoding(path) == expected
        handler.import_component(path, OfficeProject(existing))

    assert existing.CodeModule.body == CODE


def test_detection_keeps_the_encoding_of_an_existing_file_on_export(tmp_path):
    handler = make_handler(tmp_path, None)
    with patch("vba_edit.utils.get_windows_ansi_codepage", return_value="cp1252"):
        (handler.vba_dir / "Module1.bas").write_text("' früher\n", encoding="utf-8")

        used_existing = export(handler, module("Module1"))
        used_new = export(handler, module("Module2"))

    assert used_existing == "utf-8"
    assert "Grüße".encode("utf-8") in (handler.vba_dir / "Module1.bas").read_bytes()
    assert used_new == "cp1252"  # a new file gets the Windows code page, like every file by default
    assert "Grüße".encode("cp1252") in (handler.vba_dir / "Module2.bas").read_bytes()


def test_metadata_records_the_mode_and_the_encoding_of_each_file(tmp_path):
    handler = make_handler(tmp_path, None)
    component = module("Module1")
    (handler.vba_dir / "Module1.bas").write_text("' früher\n", encoding="utf-8")
    with (
        patch("vba_edit.utils.get_windows_ansi_codepage", return_value="cp1252"),
        patch.object(handler.component_handler, "get_component_info", return_value=info_for(component)),
    ):
        encodings = handler._export_components([component], overwrite=True)
    handler._save_metadata(encodings)

    metadata = json.loads((handler.vba_dir / "vba_metadata.json").read_text(encoding="utf-8"))
    assert metadata["encoding_mode"] == "detect"
    assert metadata["encodings"]["Module1"]["encoding"] == "utf-8"

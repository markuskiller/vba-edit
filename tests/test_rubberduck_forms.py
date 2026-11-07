"""Tests: with --rubberduck-folders a UserForm's .frm, .header and .frx stay together in its @Folder.

Before 0.6.0 export wrote the .frm and .header into the module's @Folder but left the .frx in the
VBA directory root, where import never looked for it. These tests pin the file layout on export,
the lookup on import (including the fallback for trees exported by older versions) and the
discovery helpers that have to look into the subfolders.
"""

import os
import stat
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from tests.test_form_import import CODE, FORM_HEADER, FakeComponents
from vba_edit.office_vba import ExcelVBAHandler

FOLDER_CODE = '\'@Folder("Forms.Dialogs")\n' + CODE
FRX = b"\x00\x01binary"


@pytest.fixture
def handler(tmp_path):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        h = ExcelVBAHandler(
            doc_path=str(workbook), vba_dir=str(tmp_path / "src"), save_headers=True, use_rubberduck_folders=True
        )
    h.vba_dir.mkdir(exist_ok=True)
    h.doc = Mock(FullName=str(workbook))
    return h


def form_component():
    """A fake VBComponent whose Export writes the .frm text and, like the VBE, a .frx next to it."""
    component = Mock(Name="TestForm", Type=3)
    component.CodeModule.CountOfLines = 2

    def export(path):
        Path(path).write_text(f"{FORM_HEADER}\n{FOLDER_CODE}\n", encoding="cp1252")
        Path(path).with_suffix(".frx").write_bytes(FRX)

    component.Export = export
    return component


def write_form(handler, folder="Forms/Dialogs", inline=False, frx_in="folder"):
    target = handler.vba_dir / folder
    target.mkdir(parents=True, exist_ok=True)
    frm = target / "TestForm.frm"
    frm.write_text((f"{FORM_HEADER}\n" if inline else "") + CODE + "\n", encoding="cp1252")
    if not inline:
        (target / "TestForm.header").write_text(FORM_HEADER + "\n", encoding="cp1252")
    if frx_in == "folder":
        (target / "TestForm.frx").write_bytes(FRX)
    elif frx_in == "root":
        (handler.vba_dir / "TestForm.frx").write_bytes(FRX)
    return frm


def test_export_puts_frm_header_and_frx_into_the_rubberduck_folder(handler):
    handler.export_component(form_component(), handler.vba_dir)

    folder = handler.vba_dir / "Forms" / "Dialogs"
    assert "UserForm_Click" in (folder / "TestForm.frm").read_text(encoding="cp1252")
    assert (folder / "TestForm.header").read_text(encoding="cp1252").startswith("VERSION 5.00")
    assert (folder / "TestForm.frx").read_bytes() == FRX
    assert not (handler.vba_dir / "TestForm.frx").exists()
    assert not os.stat(folder / "TestForm.frx").st_mode & stat.S_IWRITE  # read-only, as in the root case


def test_re_export_replaces_the_read_only_frx_in_the_folder(handler):
    handler.export_component(form_component(), handler.vba_dir)
    handler.export_component(form_component(), handler.vba_dir)

    assert (handler.vba_dir / "Forms" / "Dialogs" / "TestForm.frx").read_bytes() == FRX
    assert not (handler.vba_dir / "TestForm.frx").exists()


def test_import_finds_header_and_frx_beside_the_form_in_its_folder(handler):
    existing = Mock(Name="TestForm", Type=3)
    components = FakeComponents(existing)

    handler.import_component(write_form(handler), components)

    assert components.removed == ["TestForm"]
    ((frm_text, frx_bytes),) = components.imported
    assert frm_text.startswith("VERSION 5.00") and "UserForm_Click" in frm_text
    assert frx_bytes == FRX


def test_import_falls_back_to_a_frx_left_in_the_root_by_older_exports(handler):
    components = FakeComponents()

    handler.import_component(write_form(handler, frx_in="root"), components)

    ((_frm_text, frx_bytes),) = components.imported
    assert frx_bytes == FRX


def test_inline_header_form_with_root_frx_is_staged_together(handler):
    components = FakeComponents()

    handler.import_component(write_form(handler, inline=True, frx_in="root"), components)

    ((frm_text, frx_bytes),) = components.imported
    assert frm_text.startswith("VERSION 5.00") and frx_bytes == FRX
    assert not list(handler.vba_dir.glob(".*_import_tmp"))  # staging folder removed


def test_edit_start_skips_modules_already_exported_into_their_folder(handler):
    write_form(handler)
    component = Mock(Name="TestForm", Type=3)
    component.CodeModule.CountOfLines = 2
    handler.export_component = Mock()

    handler._export_components([component], overwrite=False)
    handler.export_component.assert_not_called()

    handler._export_components([component], overwrite=True)
    handler.export_component.assert_called_once()


def test_header_cleanup_removes_each_header_once(handler, caplog):
    write_form(handler)
    (handler.vba_dir / "Root.header").write_text("header", encoding="cp1252")

    with caplog.at_level("WARNING", logger="vba_edit.office_vba"):
        handler._cleanup_old_header_files()

    assert not list(handler.vba_dir.rglob("*.header"))
    assert "Could not remove" not in caplog.text


def test_hidden_folders_are_ignored_by_discovery(handler):
    write_form(handler)
    staging = handler.vba_dir / ".TestForm_import_tmp"
    staging.mkdir()
    (staging / "TestForm.frm").write_text("code", encoding="cp1252")
    state = handler.vba_dir / ".vba-edit" / "conflicts"
    state.mkdir(parents=True)
    (state / "Module1.bas").write_text("code", encoding="cp1252")

    assert [p.name for p in handler._find_vba_files("*.frm")] == ["TestForm.frm"]
    assert handler._find_vba_files("*.bas") == []
    assert handler._check_existing_vba_files() == [handler.vba_dir / "Forms" / "Dialogs" / "TestForm.frm"]

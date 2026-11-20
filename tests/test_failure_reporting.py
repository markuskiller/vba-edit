"""Tests: a command that did not do all of its work says so, in its message and in its exit code.

Covers an export in which components fail, reference imports with failed entries, a
config file that cannot be read, the header clean-up under a module filter, and the save
after a live reference change.
"""

import argparse
import logging
import sys
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import pytest

from vba_edit.exceptions import VBAError
from vba_edit.office_cli import OfficeVBACLI
from vba_edit.office_vba import ExcelVBAHandler
from vba_edit.reference_manager import ReferenceManager


def make_handler(tmp_path, **options):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        handler = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"), **options)
    handler.vba_dir.mkdir(exist_ok=True)
    handler.office_encoding = "cp1252"
    return handler


# --- export ---------------------------------------------------------------------------------------------


class Module:
    def __init__(self, name, broken=False):
        self.Name = name
        self.Type = 1
        self.broken = broken
        self.CodeModule = Mock(CountOfLines=2)

    def Export(self, path):
        if self.broken:
            raise RuntimeError("the VBA editor could not export this module")
        Path(path).write_text(f'Attribute VB_Name = "{self.Name}"\nSub A()\nEnd Sub\n', encoding="cp1252")


def export(handler, *modules):
    components = MagicMock()
    components.Count = len(modules)
    components.__iter__.side_effect = lambda: iter(modules)
    handler.is_document_open = Mock(return_value=True)
    handler.get_vba_project = Mock(return_value=Mock(VBComponents=components))
    handler.close_document = Mock()
    handler.export_vba(interactive=False)


def test_export_names_the_components_that_failed_and_still_writes_the_others(tmp_path):
    handler = make_handler(tmp_path)

    with pytest.raises(VBAError) as failure:
        export(handler, Module("Good"), Module("Broken", broken=True))

    assert "1 of 2 VBA components could not be exported: Broken" in str(failure.value)
    assert (handler.vba_dir / "Good.bas").exists()
    assert not (handler.vba_dir / "Broken.bas").exists()
    handler.close_document.assert_called_once()  # a document opened for the export is not left open


def test_export_in_which_every_component_fails_is_an_error(tmp_path):
    handler = make_handler(tmp_path)

    with pytest.raises(VBAError, match="2 of 2 VBA components could not be exported"):
        export(handler, Module("One", broken=True), Module("Two", broken=True))

    assert list(handler.vba_dir.glob("*.bas")) == []


def test_export_without_failures_returns_normally(tmp_path):
    handler = make_handler(tmp_path)

    export(handler, Module("One"), Module("Two"))

    assert sorted(p.name for p in handler.vba_dir.glob("*.bas")) == ["One.bas", "Two.bas"]


def test_export_command_exits_with_an_error_code(tmp_path):
    handler = MagicMock()
    args = argparse.Namespace(
        command="export", with_references=False, keep_open=False, save_metadata=False, force_overwrite=False
    )
    with patch("vba_edit.office_cli.handle_export_with_warnings", side_effect=VBAError("1 of 2 ... could not")):
        with pytest.raises(SystemExit) as stop:
            OfficeVBACLI("excel")._run_vba_command(handler, args)

    assert stop.value.code == 1


# --- references ------------------------------------------------------------------------------------------


def run_references_import(tmp_path, stats, subcommand_args=None):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    refs = tmp_path / "Book_refs.toml"
    refs.write_text("")
    document = MagicMock(FullName=str(workbook.resolve()))
    app = MagicMock()
    app.Workbooks.Count = 0
    app.Workbooks.Open.return_value = document
    manager = MagicMock(spec=ReferenceManager)
    manager.import_from_toml.return_value = stats
    manager.sync_from_toml.return_value = stats
    args = argparse.Namespace(
        refs_subcommand="import", file=str(workbook), refs_file=None, verbose=False, logfile=None, sync=False
    )
    for key, value in (subcommand_args or {}).items():
        setattr(args, key, value)
    steps = []
    document.Save.side_effect = lambda: steps.append("save")
    document.Close.side_effect = lambda **_kwargs: steps.append("close")
    with (
        patch("win32com.client.Dispatch", return_value=app),
        patch("vba_edit.office_cli.ReferenceManager", return_value=manager),
        patch("vba_edit.office_cli.setup_logging"),
    ):
        try:
            OfficeVBACLI("excel")._handle_references_command(args)
            code = 0
        except SystemExit as stop:
            code = stop.code
    return code, steps


def test_references_import_with_a_failed_entry_exits_with_an_error_code_after_saving_the_rest(tmp_path):
    code, steps = run_references_import(tmp_path, {"added": 1, "skipped": 0, "failed": 1})

    assert code == 1
    assert steps == ["save", "close"]  # the reference that could be added is not thrown away


def test_references_sync_with_a_failed_entry_exits_with_an_error_code(tmp_path):
    stats = {"added": 0, "skipped": 2, "removed": 1, "protected": 0, "failed": 1}

    code, _steps = run_references_import(tmp_path, stats, {"sync": True})

    assert code == 1


def test_references_import_without_failures_exits_normally(tmp_path):
    code, steps = run_references_import(tmp_path, {"added": 1, "skipped": 1, "failed": 0})

    assert code == 0
    assert steps == ["save", "close"]


def run_with_references(command, tmp_path, manager):
    handler = MagicMock()
    handler.vba_dir = tmp_path
    handler.doc_path = tmp_path / "Book.xlsm"
    (tmp_path / "Book_refs.toml").write_text("")
    args = argparse.Namespace(
        command=command, with_references=True, keep_open=False, save_metadata=False, force_overwrite=False
    )
    with (
        patch("vba_edit.office_cli.ReferenceManager", return_value=manager),
        patch("vba_edit.office_cli.handle_export_with_warnings"),
    ):
        try:
            OfficeVBACLI("excel")._run_vba_command(handler, args)
            code = 0
        except SystemExit as stop:
            code = stop.code
    return code, handler


def test_import_with_references_imports_the_code_and_reports_the_failed_reference(tmp_path):
    manager = MagicMock()
    manager.import_from_toml.return_value = {"added": 0, "skipped": 0, "failed": 1}

    code, handler = run_with_references("import", tmp_path, manager)

    handler.import_vba.assert_called_once()
    assert code == 1


def test_import_with_references_that_cannot_be_read_is_an_error(tmp_path):
    manager = MagicMock()
    manager.import_from_toml.side_effect = RuntimeError("no VBA project")

    code, handler = run_with_references("import", tmp_path, manager)

    handler.import_vba.assert_called_once()
    assert code == 1


def test_export_with_references_that_cannot_be_written_is_an_error_and_closes_the_document(tmp_path):
    manager = MagicMock()
    manager.export_to_toml.side_effect = RuntimeError("no VBA project")

    code, handler = run_with_references("export", tmp_path, manager)

    assert code == 1
    handler.close_document.assert_called_once()


# --- config file -------------------------------------------------------------------------------------------


@pytest.mark.parametrize("content", ["[general\nfile = 'Book.xlsm'\n", None])
def test_a_config_file_that_cannot_be_read_stops_the_command(content, tmp_path, monkeypatch):
    """Otherwise the command would run on the defaults: the active document and its folder."""
    config = tmp_path / "vba-config.toml"
    if content is not None:
        config.write_text(content, encoding="utf-8")  # malformed; None: the file does not exist
    monkeypatch.setattr(sys, "argv", ["excel-vba", "import", "--conf", str(config)])
    monkeypatch.setattr("vba_edit.office_cli.setup_logging", lambda **_kwargs: None)
    messages = []
    monkeypatch.setattr("vba_edit.office_cli.error", messages.append)
    reached = []
    monkeypatch.setattr(OfficeVBACLI, "handle_office_vba_command", lambda self, args: reached.append(args))
    monkeypatch.setattr(OfficeVBACLI, "_handle_references_command", lambda self, args: reached.append(args))

    with pytest.raises(SystemExit) as stop:
        OfficeVBACLI("excel").main()

    assert stop.value.code == 1
    assert reached == []
    assert "Error loading configuration file" in messages[0]
    assert "Nothing was done" in messages[-1]


# --- header clean-up under a module filter ---------------------------------------------------------------------


def test_header_cleanup_leaves_the_headers_of_filtered_out_modules(tmp_path):
    handler = make_handler(tmp_path, in_file_headers=True, exclude=["Vendor_*", "LockedForm.frm"])
    for name in ("Mine.cls", "Mine.header", "Vendor_Lib.cls", "Vendor_Lib.header", "LockedForm.frm"):
        (handler.vba_dir / name).write_text("")
    (handler.vba_dir / "LockedForm.header").write_text("VERSION 5.00\n")
    (handler.vba_dir / "Orphan.header").write_text("")  # no code file: cannot be matched, stays

    handler._cleanup_old_header_files()

    assert sorted(p.name for p in handler.vba_dir.glob("*.header")) == [
        "LockedForm.header",
        "Orphan.header",
        "Vendor_Lib.header",
    ]


def test_header_cleanup_follows_an_include_filter(tmp_path):
    handler = make_handler(tmp_path, in_file_headers=True, include=["Mine"])
    for name in ("Mine.cls", "Mine.header", "Other.cls", "Other.header"):
        (handler.vba_dir / name).write_text("")

    handler._cleanup_old_header_files()

    assert [p.name for p in handler.vba_dir.glob("*.header")] == ["Other.header"]


def test_header_cleanup_without_a_filter_removes_every_header(tmp_path):
    handler = make_handler(tmp_path, in_file_headers=True)
    for name in ("Mine.cls", "Mine.header", "Orphan.header"):
        (handler.vba_dir / name).write_text("")

    handler._cleanup_old_header_files()

    assert list(handler.vba_dir.glob("*.header")) == []


# --- live reference changes in an edit session ----------------------------------------------------------------------


@pytest.mark.parametrize(
    "stats, saves",
    [
        ({"added": 1, "skipped": 2, "failed": 0}, 1),
        ({"added": 0, "skipped": 3, "failed": 0}, 0),  # nothing changed: nothing to save
        ({"added": 1, "skipped": 0, "failed": 1}, 1),
    ],
)
def test_a_saved_references_file_is_followed_by_a_save_of_the_document(tmp_path, stats, saves, caplog):
    handler = make_handler(tmp_path)
    handler.doc = Mock()
    handler.get_vba_project = Mock()
    manager = Mock()
    manager.import_from_toml.return_value = stats

    with caplog.at_level(logging.INFO), patch("vba_edit.reference_manager.ReferenceManager", return_value=manager):
        handler._handle_toml_change(str(tmp_path / "Book_refs.toml"))

    assert handler.doc.Save.call_count == saves
    level = logging.WARNING if stats["failed"] else logging.INFO
    assert any(r.levelno == level and "References re-imported" in r.message for r in caplog.records)

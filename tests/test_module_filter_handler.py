"""Handler-level tests for --include / --exclude (GitHub issue #77).

Covers export, import, the edit-mode watcher and the existing-files check in
OfficeVBAHandler with the Office COM layer mocked out.
"""

import tempfile
import time
from pathlib import Path
from unittest.mock import Mock, patch

import pytest
from watchfiles import Change

from vba_edit.office_vba import ExcelVBAHandler, WordVBAHandler


@pytest.fixture
def temp_dir():
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


def _make_component(name: str, comp_type: int, lines: int = 3, code: str = "") -> Mock:
    comp = Mock()
    comp.Name = name
    comp.Type = comp_type
    comp.CodeModule.CountOfLines = lines
    comp.CodeModule.Lines.return_value = code
    return comp


def _export_handler(temp_dir, **kwargs):
    doc_path = temp_dir / "test.xlsm"
    doc_path.touch()
    with patch("win32com.client.Dispatch"):
        handler = ExcelVBAHandler(doc_path=str(doc_path), vba_dir=str(temp_dir), **kwargs)
    handler.is_document_open = Mock(return_value=True)
    handler.export_component = Mock()
    handler._check_existing_vba_files = Mock(return_value=[])
    handler._check_header_mode_change = Mock(return_value=False)
    handler._check_form_safety = Mock()
    handler._save_metadata = Mock()
    handler.close_document = Mock()
    return handler


def _run_export(handler, comps):
    mock_components = Mock()
    mock_components.Count = len(comps)
    mock_components.__iter__ = Mock(return_value=iter(comps))
    mock_vbproject = Mock()
    mock_vbproject.VBComponents = mock_components
    handler.get_vba_project = Mock(return_value=mock_vbproject)
    handler.export_vba(save_metadata=False, overwrite=True, interactive=False)
    return [call.args[0].Name for call in handler.export_component.call_args_list]


class TestModuleFilterExport:
    def test_include_limits_exported_components(self, temp_dir):
        handler = _export_handler(temp_dir, include=["Module1", "frm*"])
        exported = _run_export(
            handler,
            [_make_component("Sheet1", 100), _make_component("Module1", 1), _make_component("frmMain", 3)],
        )
        assert exported == ["Module1", "frmMain"]

    def test_exclude_removes_components_including_forms(self, temp_dir):
        handler = _export_handler(temp_dir, exclude=["Vendor_*", "LockedForm.frm"])
        exported = _run_export(
            handler,
            [
                _make_component("Vendor_Json", 1),
                _make_component("LockedForm", 3),
                _make_component("MyForm", 3),
                _make_component("Module1", 1),
            ],
        )
        assert exported == ["MyForm", "Module1"]

    def test_no_filter_exports_everything(self, temp_dir):
        handler = _export_handler(temp_dir)
        exported = _run_export(handler, [_make_component("A", 1), _make_component("B", 2)])
        assert exported == ["A", "B"]

    def test_nothing_matches_exports_nothing(self, temp_dir, caplog):
        handler = _export_handler(temp_dir, include=["DoesNotExist"])
        with caplog.at_level("WARNING", logger="vba_edit.office_vba"):
            exported = _run_export(handler, [_make_component("A", 1)])
        assert exported == []
        assert "No VBA components matched" in caplog.text

    def test_rubberduck_folder_pattern_reads_component_code(self, temp_dir):
        handler = _export_handler(temp_dir, use_rubberduck_folders=True, include=["Utils/*"])
        in_utils = _make_component("Strings", 1, code='\'@Folder("Utils")\nSub A()\nEnd Sub')
        nested = _make_component("Deep", 1, code='\'@Folder("Utils.Sub")\nSub B()\nEnd Sub')
        at_root = _make_component("Main", 1, code="Sub C()\nEnd Sub")
        elsewhere = _make_component("Other", 1, code='\'@Folder("Reports")\nSub D()\nEnd Sub')
        assert _run_export(handler, [in_utils, nested, at_root, elsewhere]) == ["Strings", "Deep"]

    def test_folder_not_read_when_pattern_has_no_separator(self, temp_dir):
        handler = _export_handler(temp_dir, use_rubberduck_folders=True, include=["Strings"])
        comp = _make_component("Strings", 1)
        _run_export(handler, [comp])
        comp.CodeModule.Lines.assert_not_called()


def _import_handler(temp_dir, **kwargs):
    doc_path = temp_dir / "test.xlsm"
    doc_path.touch()
    with patch("win32com.client.Dispatch"):
        handler = ExcelVBAHandler(doc_path=str(doc_path), vba_dir=str(temp_dir), **kwargs)
    handler.doc = Mock()
    handler.doc.Name = "test.xlsm"
    mock_vbproject = Mock()
    mock_vbproject.VBComponents = Mock()
    handler.get_vba_project = Mock(return_value=mock_vbproject)
    handler.import_component = Mock()
    handler.save_document = Mock()
    handler.is_document_open = Mock(return_value=True)
    return handler


def _write_module(path: Path) -> None:
    path.write_text(f'Attribute VB_Name = "{path.stem}"\nSub X()\nEnd Sub\n')


class TestModuleFilterImport:
    def test_exclude_skips_matching_files(self, temp_dir):
        for name in ("Module1.bas", "Vendor_Lib.bas", "Form1.frm"):
            _write_module(temp_dir / name)
        handler = _import_handler(temp_dir, exclude=["Vendor_*"])

        handler.import_vba()

        imported = sorted(call.args[0].name for call in handler.import_component.call_args_list)
        assert imported == ["Form1.frm", "Module1.bas"]
        handler.save_document.assert_called_once()

    def test_include_pattern_relative_to_vba_directory(self, temp_dir):
        (temp_dir / "Utils").mkdir()
        _write_module(temp_dir / "Utils" / "Strings.bas")
        _write_module(temp_dir / "Main.bas")
        handler = _import_handler(temp_dir, use_rubberduck_folders=True, include=["Utils/*"])

        handler.import_vba()

        imported = [call.args[0].name for call in handler.import_component.call_args_list]
        assert imported == ["Strings.bas"]

    def test_nothing_matches_imports_nothing(self, temp_dir, caplog):
        _write_module(temp_dir / "Main.bas")
        handler = _import_handler(temp_dir, include=["mod_10*.bas"])

        with caplog.at_level("WARNING", logger="vba_edit.office_vba"):
            handler.import_vba()

        handler.import_component.assert_not_called()
        handler.save_document.assert_not_called()
        assert "No VBA files matched" in caplog.text


class TestModuleFilterWatch:
    def test_watcher_ignores_filtered_files_including_deletions(self, temp_dir):
        handler = _import_handler(temp_dir, exclude=["Vendor_*"])
        handler._handle_vba_change = Mock()
        changes = {
            (Change.modified, str(temp_dir / "Vendor_A.bas")),
            (Change.deleted, str(temp_dir / "Vendor_B.cls")),
            (Change.modified, str(temp_dir / "Module1.bas")),
            (Change.deleted, str(temp_dir / "Module2.bas")),
        }

        handler._process_changes_batch(changes, "test_refs.toml", False, time.time(), 5, {".bas", ".cls", ".frm"})

        handled = sorted(call.args[1].name for call in handler._handle_vba_change.call_args_list)
        assert handled == ["Module1.bas", "Module2.bas"]

    def test_watcher_without_filter_handles_everything(self, temp_dir):
        handler = _import_handler(temp_dir)
        handler._handle_vba_change = Mock()
        changes = {(Change.modified, str(temp_dir / "Vendor_A.bas")), (Change.modified, str(temp_dir / "M.bas"))}

        handler._process_changes_batch(changes, "test_refs.toml", False, time.time(), 5, {".bas", ".cls", ".frm"})

        assert handler._handle_vba_change.call_count == 2

    def test_existing_files_check_respects_filter(self, temp_dir):
        for name in ("Module1.bas", "Vendor_Lib.bas"):
            (temp_dir / name).write_text("")
        doc_path = temp_dir / "test.xlsm"
        doc_path.touch()
        with patch("win32com.client.Dispatch"):
            handler = ExcelVBAHandler(doc_path=str(doc_path), vba_dir=str(temp_dir), include=["Module1"])
        assert [f.name for f in handler._check_existing_vba_files()] == ["Module1.bas"]


class TestModuleFilterHandlerInit:
    def test_filter_attribute_set_from_kwargs(self, temp_dir):
        doc_path = temp_dir / "test.docm"
        doc_path.touch()
        with patch("win32com.client.Dispatch"):
            handler = WordVBAHandler(doc_path=str(doc_path), vba_dir=str(temp_dir), include=["A"], exclude=["B"])
        assert handler.module_filter.include == ("A",)
        assert handler.module_filter.exclude == ("B",)

    def test_filter_inactive_by_default(self, temp_dir):
        doc_path = temp_dir / "test.docm"
        doc_path.touch()
        with patch("win32com.client.Dispatch"):
            handler = WordVBAHandler(doc_path=str(doc_path), vba_dir=str(temp_dir))
        assert handler.module_filter.active is False

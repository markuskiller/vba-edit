"""Tests: import failures surface instead of producing stray modules or a silent exit code 0."""

import logging
from unittest.mock import Mock, patch

import pytest

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import ExcelVBAHandler


class FakeCodeModule:
    def __init__(self, body="", fail_on=None):
        self.body = body
        self.fail_on = fail_on

    @property
    def CountOfLines(self):
        return len(self.body.splitlines())

    def Lines(self, _start, _count):
        return self.body

    def DeleteLines(self, _start, _count):
        self.body = ""

    def AddFromString(self, code):
        if self.fail_on is not None and self.fail_on in code:
            raise RuntimeError("VBE refused the code")
        self.body = code


class FakeComponents:
    """VBComponents stand-in that records additions and removals."""

    def __init__(self, **existing):
        self.existing = existing
        self.added = []
        self.removed = []

    def __call__(self, name):
        if name not in self.existing:
            raise KeyError(name)
        return self.existing[name]

    def Add(self, kind):
        component = Mock(Type=kind, CodeModule=FakeCodeModule())
        self.added.append(component)
        return component

    def Remove(self, component):
        self.removed.append(component)


@pytest.fixture
def handler(tmp_path):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        h = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"))
    h.vba_dir.mkdir(exist_ok=True)
    return h


def test_a_failed_update_of_an_existing_module_surfaces_and_creates_nothing(handler):
    module = Mock(Name="Module1", Type=1, CodeModule=FakeCodeModule("Sub Old()\nEnd Sub", fail_on="New"))
    components = FakeComponents(Module1=module)
    code_file = handler.vba_dir / "Module1.bas"
    code_file.write_text("Sub New()\nEnd Sub\n", encoding="cp1252")

    with pytest.raises(VBAError, match="Module1.bas"):
        handler.import_component(code_file, components)

    assert components.added == []  # no stray Module2
    assert module.CodeModule.body == "Sub Old()\nEnd Sub"  # previous code restored


def test_a_missing_module_is_created(handler):
    components = FakeComponents()
    code_file = handler.vba_dir / "Helper.bas"
    code_file.write_text("Sub Help()\nEnd Sub\n", encoding="cp1252")

    handler.import_component(code_file, components)

    assert len(components.added) == 1
    assert components.added[0].Type == 1
    assert "Sub Help()" in components.added[0].CodeModule.body


def test_an_existing_standard_module_is_updated_in_place(handler):
    module = Mock(Name="Module1", Type=1, CodeModule=FakeCodeModule("Sub Old()\nEnd Sub"))
    components = FakeComponents(Module1=module)
    code_file = handler.vba_dir / "Module1.bas"
    code_file.write_text("Sub New()\nEnd Sub\n", encoding="cp1252")

    handler.import_component(code_file, components)

    assert module.CodeModule.body == "Sub New()\nEnd Sub"
    assert components.added == [] and components.removed == []


def prepare_import(handler, *names):
    for name in names:
        (handler.vba_dir / f"{name}.bas").write_text(f"Sub {name}()\nEnd Sub\n", encoding="cp1252")
    handler.doc = Mock()
    handler.get_vba_project = Mock(return_value=Mock(VBComponents=FakeComponents()))
    handler.save_document = Mock()


def test_import_vba_reports_failures_and_raises(handler, caplog):
    prepare_import(handler, "Good", "Bad", "AlsoGood")

    def import_component(path, _components):
        if path.stem == "Bad":
            raise VBAError("Failed to handle Bad.bas")

    handler.import_component = Mock(side_effect=import_component)

    with caplog.at_level(logging.ERROR), pytest.raises(VBAError) as error:
        handler.import_vba()

    assert "Imported 2 of 3 VBA files" in str(error.value)
    assert "1 failed: Bad.bas" in str(error.value)
    handler.save_document.assert_called_once()  # the good ones are kept
    assert any("Failed to import Bad.bas" in record.message for record in caplog.records)


def test_import_vba_summarises_a_clean_run(handler, caplog):
    prepare_import(handler, "One", "Two")
    handler.import_component = Mock()

    with caplog.at_level(logging.INFO):
        handler.import_vba()

    assert any(record.message == "Imported 2 of 2 VBA files" for record in caplog.records)


def test_import_vba_counts_skipped_empty_files_separately(handler, caplog):
    prepare_import(handler, "Full")
    (handler.vba_dir / "Empty.bas").write_text('Attribute VB_Name = "Empty"\n', encoding="cp1252")
    handler.skip_empty = True
    handler.import_component = Mock()

    with caplog.at_level(logging.INFO):
        handler.import_vba()

    assert any(record.message == "Imported 1 of 2 VBA files (1 empty file(s) skipped)" for record in caplog.records)
    assert handler.import_component.call_count == 1

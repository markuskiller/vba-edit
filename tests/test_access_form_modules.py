"""Tests for importing the code modules of Access forms and reports (PR #117)."""

import logging
from pathlib import Path
from types import SimpleNamespace

import pytest

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import (
    AccessVBAHandler,
    OfficeVBAHandler,
    VBAComponentHandler,
    VBADocumentNames,
    VBAModuleType,
)

DOCUMENT, CLASS, STANDARD = VBAModuleType.DOCUMENT, VBAModuleType.CLASS, VBAModuleType.STANDARD


class FakeCodeModule:
    """Minimal stand-in for a VBIDE CodeModule holding its code as text."""

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
            raise RuntimeError("Access rejected the code")
        self.body = code


def project(**component_types):
    """Return a VBComponents stand-in holding components of the given VBIDE types."""
    components = {
        name: SimpleNamespace(Name=name, Type=kind, CodeModule=FakeCodeModule())
        for name, kind in component_types.items()
    }

    def lookup(name):
        if name not in components:
            raise KeyError(name)  # COM raises for an unknown component as well.
        return components[name]

    return lookup


def corrected(app_name, name, module_type, components):
    handler = SimpleNamespace(app_name=app_name)
    return OfficeVBAHandler._module_type_for_existing_component(handler, name, module_type, components)


def update_access_module(name, code, components):
    handler = AccessVBAHandler.__new__(AccessVBAHandler)  # No COM; only the update methods are used.
    handler._update_document_module(name, code, components)


# --- Names -----------------------------------------------------------------------------------------


def test_access_form_and_report_names_only_count_in_access():
    assert VBADocumentNames.is_document_module("Form_Customers", "Access")
    assert VBADocumentNames.is_document_module("Report_Sales", "Access")

    for app_name in ("Excel", "Word", "PowerPoint", None):
        assert not VBADocumentNames.is_document_module("Form_Customers", app_name)
        assert not VBADocumentNames.is_document_module("Report_Sales", app_name)


def test_access_prefix_is_not_localized():
    """A German Access names the module of "Formular1" Form_Formular1, never Formular_..."""
    assert VBADocumentNames.is_document_module("Form_Formular1", "Access")
    assert VBADocumentNames.is_document_module("Report_Bericht1", "Access")

    for name in ("Formular_Helper", "Bericht_Export", "Rapport_Mensuel", "Informe_Ventas", "Formulaire_Client"):
        assert not VBADocumentNames.is_document_module(name, "Access")


def test_names_need_the_underscore_separator():
    for name in ("FormHelper", "Reporting", "Form", "Report"):
        assert not VBADocumentNames.is_document_module(name, "Access")


def test_sheet_and_slide_names_count_in_their_own_application():
    assert VBADocumentNames.is_document_module("Sheet1", "Excel")
    assert not VBADocumentNames.is_document_module("Sheet1", "Access")
    assert VBADocumentNames.is_document_module("Slide1", "PowerPoint")
    assert not VBADocumentNames.is_document_module("Slide1", "Excel")


def test_without_an_application_names_are_matched_as_before():
    """Callers that do not pass an application keep the behaviour they had."""
    for name in ("ThisWorkbook", "ThisDocument", "Sheet1", "Tabelle2", "Slide1"):
        assert VBADocumentNames.is_document_module(name)

    assert VBAComponentHandler().get_module_type(Path("Sheet1.cls")) == DOCUMENT
    assert VBAComponentHandler(True).use_rubberduck_folders is True
    assert VBAComponentHandler(True).app_name is None


def test_access_form_name_on_a_standard_module_file_stays_a_standard_module():
    handler = VBAComponentHandler(app_name="Access")

    assert handler.get_module_type(Path("Form_Customers.cls")) == DOCUMENT
    assert handler.get_module_type(Path("Form_Utils.bas")) == STANDARD
    assert handler.get_module_type(Path("Report_Utils.bas")) == STANDARD


# --- The existing component settles it ---------------------------------------------------------------


def test_existing_form_module_is_updated_in_place_whatever_the_file_suggests():
    components = project(Form_Customers=100, frmOddName=100)

    assert corrected("Access", "Form_Customers", DOCUMENT, components) == DOCUMENT
    assert corrected("Access", "frmOddName", CLASS, components) == DOCUMENT


def test_existing_document_module_wins_in_every_application():
    """For example an Excel sheet with a custom code name and no header file."""
    components = project(shtData=100)

    assert corrected("Excel", "shtData", CLASS, components) == DOCUMENT


def test_ordinary_access_module_with_a_form_name_is_not_a_document_module():
    components = project(Form_Helper=2, Report_Utils=1)

    assert corrected("Access", "Form_Helper", DOCUMENT, components) == CLASS
    assert corrected("Access", "Report_Utils", DOCUMENT, components) == STANDARD


def test_missing_component_leaves_the_decision_to_the_file():
    components = project()

    assert corrected("Access", "Form_Customers", DOCUMENT, components) == DOCUMENT
    assert corrected("Access", "Helper", CLASS, components) == CLASS


def test_other_applications_keep_the_file_based_type_for_ordinary_components():
    components = project(Sheet1=2, Module1=1)

    assert corrected("Excel", "Sheet1", DOCUMENT, components) == DOCUMENT
    assert corrected("Excel", "Module1", STANDARD, components) == STANDARD


# --- Updating the module -----------------------------------------------------------------------------


def test_form_module_code_is_replaced_in_place():
    components = project(Form_Customers=100)
    components("Form_Customers").CodeModule.body = "Sub Old()\nEnd Sub"

    update_access_module("Form_Customers", "Sub New()\nEnd Sub", components)

    assert components("Form_Customers").CodeModule.body == "Sub New()\nEnd Sub"


def test_missing_form_gives_a_clear_error_instead_of_a_stray_class_module():
    with pytest.raises(VBAError) as error:
        update_access_module("Form_Missing", "Sub New()\nEnd Sub", project())

    message = str(error.value)
    assert "Form_Missing" in message
    assert "Has Module" in message


def test_failed_update_restores_the_previous_form_code(caplog):
    components = project(Form_Customers=100)
    module = components("Form_Customers").CodeModule
    module.body = "Sub Old()\nEnd Sub"
    module.fail_on = "Broken"

    with caplog.at_level(logging.WARNING), pytest.raises(VBAError):
        update_access_module("Form_Customers", "Sub Broken()\nEnd Sub", components)

    assert module.body == "Sub Old()\nEnd Sub"
    assert any("previous code was restored" in record.message for record in caplog.records)

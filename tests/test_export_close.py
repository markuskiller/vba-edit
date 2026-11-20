"""Tests: export only closes what vba-edit opened itself."""

import logging
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pytest

from vba_edit.exceptions import VBAError
from vba_edit.office_vba import ExcelVBAHandler, PowerPointVBAHandler, WordVBAHandler


class FakeDocument:
    def __init__(self, full_name, saved=True):
        self.FullName = full_name
        self.Name = full_name.rsplit("\\", 1)[-1]
        self.Saved = saved
        self.closed_with = None

    def Close(self, SaveChanges):
        self.closed_with = SaveChanges


class FakeCollection:
    """Workbooks/Documents collection: 1-based, callable like COM."""

    def __init__(self, *documents):
        self.documents = list(documents)

    @property
    def Count(self):
        return len(self.documents)

    def __call__(self, index):
        return self.documents[index - 1]

    def Open(self, path):
        document = FakeDocument(path)
        self.documents.append(document)
        return document


class FakeApp:
    def __init__(self, visible, *documents):
        self.Visible = visible
        self.Workbooks = FakeCollection(*documents)
        self.quit = False

    def Quit(self):
        self.quit = True


@pytest.fixture
def handler(tmp_path):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        h = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"))
    return h


def attach(handler, app):
    with patch("vba_edit.office_vba.win32com.client.Dispatch", return_value=app):
        handler.open_document()


def test_a_document_the_user_had_open_is_attached_and_left_open(handler, caplog):
    user_document = FakeDocument(str(handler.doc_path), saved=False)
    app = FakeApp(True, FakeDocument("C:\\other\\Budget.xlsx"), user_document)

    attach(handler, app)
    assert handler.doc is user_document  # attached, not re-opened
    assert handler.app.Workbooks.Count == 2

    with caplog.at_level(logging.INFO):
        handler.close_document()

    assert user_document.closed_with is None
    assert user_document.Saved is False
    assert app.quit is False
    assert handler.doc is None
    assert any("already open" in record.message for record in caplog.records)


def test_a_document_vba_edit_opened_in_the_users_excel_is_closed_but_excel_stays(handler):
    app = FakeApp(True, FakeDocument("C:\\other\\Budget.xlsx"))

    attach(handler, app)
    opened = handler.doc
    assert opened.FullName == str(handler.doc_path)

    handler.close_document()

    assert opened.closed_with is False
    assert app.quit is False  # Excel was running before vba-edit, with the user's other workbook


def test_an_excel_started_by_vba_edit_is_quit_again(handler):
    app = FakeApp(False)  # A fresh automation instance is invisible

    attach(handler, app)
    assert handler.app.Visible is True
    opened = handler.doc

    handler.close_document()

    assert opened.closed_with is False
    assert app.quit is True


def test_a_running_but_empty_excel_is_not_quit(handler):
    """The user has Excel open without a workbook; vba-edit opens and closes one."""
    app = FakeApp(True)

    attach(handler, app)
    handler.close_document()

    assert handler.doc is None
    assert app.quit is False


def test_matching_is_case_insensitive(tmp_path):
    workbook = tmp_path / "Book.xlsm"
    workbook.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        handler = ExcelVBAHandler(doc_path=str(workbook), vba_dir=str(tmp_path / "src"))
    same_file = FakeDocument(str(workbook).upper())
    attach(handler, FakeApp(True, same_file))

    assert handler.doc is same_file
    assert handler._document_was_open is True


def test_word_uses_the_documents_collection(tmp_path):
    document_path = tmp_path / "Report.docm"
    document_path.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        handler = WordVBAHandler(doc_path=str(document_path), vba_dir=str(tmp_path / "src"))
    app = SimpleNamespace(Visible=True, Documents=FakeCollection(FakeDocument(str(document_path))))

    with patch("vba_edit.office_vba.win32com.client.Dispatch", return_value=app):
        handler.open_document()

    assert handler._document_was_open is True
    assert handler.doc.FullName == str(document_path)


def test_mock_applications_fall_back_to_opening(handler):
    """Mocks (as in the older tests) have no usable collection; the document is opened as before."""
    app = Mock()
    app.Visible = Mock()  # not a real True: counts as "started by vba-edit"
    attach(handler, app)

    app.Workbooks.Open.assert_called_once_with(str(handler.doc_path))
    assert handler._document_was_open is False
    assert handler._app_was_running is False


class RefusingCollection(FakeCollection):
    """Excel returns nothing, without an error, for a second workbook with the same file name."""

    def Open(self, path):
        return None


def test_a_second_workbook_with_the_same_name_is_explained(handler, caplog):
    app = FakeApp(True)
    app.Workbooks = RefusingCollection(FakeDocument("C:\\other\\BOOK.xlsm"))

    with caplog.at_level(logging.ERROR), pytest.raises(VBAError):
        attach(handler, app)

    assert handler.doc is None
    assert "C:\\other\\BOOK.xlsm" in caplog.text
    assert "cannot open a second one with the same name" in caplog.text


def test_a_document_that_was_not_opened_is_an_error_not_a_none(handler, caplog):
    app = FakeApp(True)
    app.Workbooks = RefusingCollection(FakeDocument("C:\\other\\Unrelated.xlsm"))

    with caplog.at_level(logging.ERROR), pytest.raises(VBAError):
        attach(handler, app)

    assert "Excel did not open the workbook" in caplog.text


class FakePresentation:
    def __init__(self, full_name):
        self.FullName = full_name
        self.Name = full_name.rsplit("\\", 1)[-1]
        self.Saved = False
        self.closed = False

    def Close(self):  # unlike Workbook.Close and Document.Close, this one takes no arguments
        self.closed = True


class FakePresentations(FakeCollection):
    def __iter__(self):
        return iter(self.documents)

    def Open(self, path):
        presentation = FakePresentation(path)
        self.documents.append(presentation)
        return presentation


def test_a_presentation_vba_edit_opened_is_closed_again(tmp_path):
    deck = tmp_path / "Deck.pptm"
    deck.touch()
    with patch("vba_edit.office_vba.win32com.client.Dispatch"):
        handler = PowerPointVBAHandler(doc_path=str(deck), vba_dir=str(tmp_path / "src"))
    app = FakeApp(False)
    app.Presentations = FakePresentations()

    attach(handler, app)
    opened = handler.doc
    handler.close_document()

    assert opened.closed is True
    assert opened.Saved is True  # so that PowerPoint does not ask
    assert handler.doc is None
    assert app.quit is True

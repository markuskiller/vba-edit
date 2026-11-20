"""Tests: the `check` command works on a scratch file of its own and closes only what it opened.

It used to create ``test.accdb`` / ``test.xlsm`` / ... in the current directory, deleting an
existing ``test.accdb`` first and the file of that name afterwards, and it quit the application
even when the user had it open.
"""

import os
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from vba_edit.utils import CheckAccessApp, CheckExcelApp, CheckPowerPointApp, CheckWordApp, check_office_app

USER_FILES = ["test.accdb", "test.xlsm", "test.docm", "test.pptm"]


@pytest.fixture
def working_directory(tmp_path, monkeypatch):
    """A current directory that already holds files with the names the check used to claim."""
    for name in USER_FILES:
        (tmp_path / name).write_text("the user's own file")
    monkeypatch.chdir(tmp_path)
    return tmp_path


def fake_app(visible, open_database=None):
    app = MagicMock()
    app.Visible = visible
    app.CurrentDb.return_value = open_database

    def create(path, *_args, **_kwargs):
        Path(path).write_text("scratch")

    app.NewCurrentDatabase.side_effect = create
    for collection in (app.Workbooks, app.Documents, app.Presentations):
        collection.Add.return_value.SaveAs.side_effect = create
    return app


def run_check(check_class, app):
    check = check_class()
    with patch("vba_edit.utils.win32com.client.Dispatch", return_value=app):
        check_office_app(check)
    return check


@pytest.mark.parametrize("check_class", [CheckAccessApp, CheckExcelApp, CheckWordApp, CheckPowerPointApp])
def test_files_in_the_current_directory_are_neither_overwritten_nor_deleted(check_class, working_directory):
    check = run_check(check_class, fake_app(visible=False))

    for name in USER_FILES:
        assert (working_directory / name).read_text() == "the user's own file"
    scratch = Path(check.test_file)
    assert Path(tempfile.gettempdir()) in scratch.parents
    assert working_directory not in scratch.parents
    assert not os.path.exists(check.test_dir)  # the scratch folder is gone again


def test_each_check_gets_a_folder_of_its_own(working_directory):
    first = run_check(CheckExcelApp, fake_app(visible=False))
    second = run_check(CheckExcelApp, fake_app(visible=False))

    assert first.test_dir != second.test_dir


@pytest.mark.parametrize("check_class", [CheckExcelApp, CheckWordApp, CheckPowerPointApp])
def test_an_application_the_user_has_open_is_not_quit(check_class, working_directory):
    app = fake_app(visible=True)

    run_check(check_class, app)

    app.Quit.assert_not_called()


@pytest.mark.parametrize("check_class", [CheckAccessApp, CheckExcelApp, CheckWordApp, CheckPowerPointApp])
def test_an_application_the_check_started_is_quit_again(check_class, working_directory):
    app = fake_app(visible=False)

    run_check(check_class, app)

    app.Quit.assert_called_once()


def test_access_with_an_open_database_is_tested_as_it_is(working_directory):
    """Access shows one database at a time; the user's database is neither replaced nor closed."""
    app = fake_app(visible=True, open_database=MagicMock())

    check = run_check(CheckAccessApp, app)

    app.NewCurrentDatabase.assert_not_called()
    app.CloseCurrentDatabase.assert_not_called()
    app.Quit.assert_not_called()
    assert check.test_file is None  # no scratch database was created

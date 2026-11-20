"""Tests: what clears each kind of pause in bidirectional editing, as the guide describes it.

``--sync-resolve`` settles a disagreement about the code of a module. A pause about where a
module is, what it is called or whether it exists needs a step by the user first; some need
the session state to be reset. These tests pin the recovery steps documented under
"Clearing a pause", so the guide cannot drift from the engine again.
"""

import pytest

from tests.test_bidirectional import Component, change_file, session, settle, source  # noqa: F401
from vba_edit.bidirectional import OfficeAdapter, SyncSession
from vba_edit.office_vba import VBAComponentHandler

IN_MOVED = '\'@Folder("Moved")\n'


def components(session):
    return session.adapter.handler.get_vba_project().VBComponents


def restart(session, **resolution):
    restarted = SyncSession(session.adapter, **resolution)
    restarted.startup()
    return restarted


def reset_state(session):
    (session.directory / "state.json").unlink()
    return restart(session)


def with_folders(session):
    handler = session.adapter.handler
    handler.use_rubberduck_folders = True
    handler.component_handler = VBAComponentHandler(True)
    return SyncSession(OfficeAdapter(handler))


def paused(session):
    return sorted(session.state["paused"])


# --- restoring one side from the other ----------------------------------------------------------------


def test_a_deleted_file_comes_back_with_resolve_office(session):
    session.startup()
    source(session).unlink()
    settle(session)
    assert paused(session) == ["Module1"]

    resolved = restart(session, resolve="office", component="Module1")

    assert paused(resolved) == []
    assert "Sub Example()" in source(session).read_text()


def test_a_module_removed_in_office_comes_back_with_resolve_files(session):
    components(session).append(Component("Other"))
    session.startup()
    components(session).remove(components(session)("Module1"))
    settle(session)
    assert paused(session) == ["Module1"]

    resolved = restart(session, resolve="files", component="Module1")

    assert paused(resolved) == []
    assert any(component.Name == "Module1" for component in components(session))


# --- deleting and renaming need a reset --------------------------------------------------------------------


@pytest.mark.parametrize("resolution", [{}, {"resolve": "office"}, {"resolve": "files"}])
def test_a_module_deleted_on_both_sides_stays_paused_whatever_is_resolved(session, resolution):
    components(session).append(Component("Other"))
    session.startup()
    source(session).unlink()
    components(session).remove(components(session)("Module1"))
    settle(session)
    if resolution:
        resolution["component"] = "Module1"

    assert paused(restart(session, **resolution)) == ["Module1"]


def test_resetting_the_state_forgets_a_module_deleted_on_both_sides(session):
    components(session).append(Component("Other"))
    session.startup()
    source(session).unlink()
    components(session).remove(components(session)("Module1"))
    settle(session)

    fresh = reset_state(session)

    assert paused(fresh) == []
    assert sorted(fresh.state["baseline"]) == ["Other"]


def test_rename_in_the_vba_editor_then_delete_the_old_file_and_reset(session):
    session.startup()
    components(session)("Module1").Name = "Renamed"
    settle(session)
    assert source(session, "Renamed").exists()  # the new name arrives as a new file
    assert paused(session) == ["Module1"]
    source(session).unlink()
    assert paused(restart(session)) == ["Module1"]  # not cleared by a restart alone

    fresh = reset_state(session)

    assert paused(fresh) == []
    assert sorted(fresh.state["baseline"]) == ["Renamed"]


def test_a_renamed_file_is_renamed_back_and_the_state_reset(session):
    session.startup()
    source(session).rename(source(session, "Renamed"))
    settle(session)
    assert paused(session) == ["Module1", "Renamed"]
    assert [component.Name for component in components(session)] == ["Module1"]  # nothing was created
    source(session, "Renamed").rename(source(session))

    assert paused(reset_state(session)) == []


def test_a_reset_pauses_what_differs_and_resolve_settles_it(session):
    components(session).append(Component("OnlyInOffice"))
    session.startup()
    change_file(session, "Sub Different()\nEnd Sub")
    source(session, "OnlyInOffice").unlink()

    fresh = reset_state(session)
    assert paused(fresh) == ["Module1", "OnlyInOffice"]

    restart(session, resolve="files", component="Module1")
    resolved = restart(session, resolve="office", component="OnlyInOffice")

    assert paused(resolved) == []
    assert components(session)("Module1").CodeModule.body == "Sub Different()\nEnd Sub"
    assert source(session, "OnlyInOffice").exists()


# --- moves ---------------------------------------------------------------------------------------------------


def move_file(session):
    folder = session.adapter.root / "Moved"
    folder.mkdir()
    source(session).rename(folder / "Module1.bas")
    return folder / "Module1.bas"


@pytest.mark.parametrize("side", ["office", "files"])
def test_a_moved_file_is_not_settled_by_resolve_alone(session, side):
    session = with_folders(session)
    session.startup()
    move_file(session)
    settle(session)
    assert paused(session) == ["Module1"]

    assert paused(restart(session, resolve=side, component="Module1")) == ["Module1"]


def test_a_move_is_undone_by_moving_the_file_back_and_resolving_office(session):
    session = with_folders(session)
    session.startup()
    moved = move_file(session)
    settle(session)
    moved.rename(source(session))

    assert paused(restart(session, resolve="office", component="Module1")) == []


def test_a_move_is_made_real_by_the_annotation_in_the_vba_editor(session):
    session = with_folders(session)
    session.startup()
    moved = move_file(session)
    settle(session)
    module = components(session)("Module1")
    module.CodeModule.body = IN_MOVED + module.CodeModule.body

    resolved = restart(session, resolve="office", component="Module1")

    assert paused(resolved) == []
    assert '\'@Folder("Moved")' in moved.read_text()
    assert not source(session).exists()


def test_annotation_changed_in_the_vba_editor_needs_the_old_file_out_of_the_way(session):
    session = with_folders(session)
    session.startup()
    module = components(session)("Module1")
    module.CodeModule.body = IN_MOVED + module.CodeModule.body
    settle(session)
    assert paused(session) == ["Module1"]
    assert paused(restart(session, resolve="office", component="Module1")) == ["Module1"]
    source(session).unlink()

    resolved = restart(session, resolve="office", component="Module1")

    assert paused(resolved) == []
    assert (session.adapter.root / "Moved" / "Module1.bas").exists()


# --- headers and attributes -------------------------------------------------------------------------------------


def test_a_changed_header_is_only_cleared_by_taking_the_office_version(session):
    session.startup()
    text = source(session).read_text()
    source(session).write_text(text.replace("\n", '\nAttribute VB_Description = "x"\n', 1))
    settle(session)
    assert paused(session) == ["Module1"]
    assert paused(restart(session, resolve="files", component="Module1")) == ["Module1"]

    resolved = restart(session, resolve="office", component="Module1")

    assert paused(resolved) == []
    assert "VB_Description" not in source(session).read_text()


def test_a_module_with_member_attributes_is_only_cleared_by_taking_the_office_version(session):
    components(session)("Module1").attributes = 'Attribute Example.VB_Description = "Description"'
    session.startup()
    source(session).write_text(source(session).read_text().replace("Sub Example", "Sub Changed"))
    settle(session)
    assert paused(session) == ["Module1"]
    assert paused(restart(session, resolve="files", component="Module1")) == ["Module1"]

    resolved = restart(session, resolve="office", component="Module1")

    assert paused(resolved) == []
    assert "Sub Example" in source(session).read_text()
    # the edited version is still in the conflict folder
    saved = [p.read_text() for p in (session.directory / "conflicts").rglob("filesystem.txt")]
    assert any("Sub Changed" in text for text in saved)


# --- encoding ------------------------------------------------------------------------------------------------------


def test_office_exports_are_read_in_the_code_page_whatever_the_files_use(session):
    handler = session.adapter.handler
    handler.encoding = "utf-8"
    handler.office_encoding = "cp1252"
    session = SyncSession(OfficeAdapter(handler))
    components(session)("Module1").CodeModule.body = 'Sub Grüße()\n    MsgBox "5 €"\nEnd Sub'

    session.startup()

    assert paused(session) == []
    assert "Grüße".encode("utf-8") in source(session).read_bytes()

"""Deterministic regression checks for two-way synchronization and recovery."""

from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pytest

from vba_edit.bidirectional import OfficeAdapter, SessionLock, Snapshot, SyncSession, plan, run_bidirectional
from vba_edit.cli_common import merge_config_with_args, validate_bidirectional_options
from vba_edit.exceptions import VBAError
from vba_edit.office_cli import OfficeVBACLI
from vba_edit.office_vba import VBAComponentHandler


class Code:
    def __init__(self, body="Sub Example()\nEnd Sub"):
        self.body = body

    @property
    def CountOfLines(self):
        return len(self.body.splitlines())

    def Lines(self, start, count):
        assert start == 1 and count == self.CountOfLines
        return self.body


class Component:
    def __init__(self, name="Module1", kind=1, body="Sub Example()\nEnd Sub", attributes=""):
        self.Name = name
        self.Type = kind
        self.CodeModule = Code(body)
        self.attributes = attributes

    def Export(self, path):
        Path(path).write_text(
            f'Attribute VB_Name = "{self.Name}"\n{self.CodeModule.body}\n{self.attributes}', encoding="cp1252"
        )


class Components(list):
    def __call__(self, name):
        return next(c for c in self if c.Name == name)

    def Add(self, kind):
        component = Component("Added", kind, "")
        self.append(component)
        return component


@pytest.fixture
def session(tmp_path):
    document = tmp_path / "document.xlsm"
    document.touch()
    root = tmp_path / "source"
    root.mkdir()
    components = Components([Component()])
    project = SimpleNamespace(Name="Project", Mode=2, VBComponents=components)
    handler = SimpleNamespace(
        vba_dir=root,
        doc_path=document,
        app_name="Excel",
        document_type="workbook",
        encoding="cp1252",
        in_file_headers=True,
        save_headers=False,
        use_rubberduck_folders=False,
        doc=SimpleNamespace(Save=Mock()),
        get_vba_project=lambda: project,
        is_document_open=lambda: True,
        component_handler=VBAComponentHandler(),
        sync_save_metadata=False,
        _save_metadata=Mock(),
        save_document=Mock(),
    )

    def update(component, body):
        component.CodeModule.body = body

    handler._update_module_content = Mock(side_effect=update)
    handler._update_document_module = Mock(side_effect=lambda name, body, cs: update(cs(name), body))
    return SyncSession(OfficeAdapter(handler))


def settle(session):
    session.scan()
    session.scan()


def source(session, name="Module1"):
    return session.adapter.root / f"{name}.bas"


def change_file(session, body, name="Module1"):
    source(session, name).write_text(f'Attribute VB_Name = "{name}"\n{body}\n', encoding="cp1252")


@pytest.mark.parametrize(
    "office_body,file_body,action",
    [
        ("A", "A", "adopt"),
        ("B", "A", "export"),
        ("A", "B", "import"),
        ("B", "B", "adopt"),
        ("B", "C", "conflict"),
    ],
)
def test_planner(office_body, file_body, action):
    baseline = Snapshot("M", 1, "M.bas", "A")
    assert plan(baseline, replace(baseline, body=office_body), replace(baseline, body=file_body)) == action


def test_startup_rejects_equal_divergence_and_missing_files():
    baseline = Snapshot("M", 1, "M.bas", "A")
    changed = replace(baseline, body="B")
    assert plan(baseline, changed, changed, startup=True) == "conflict"
    assert plan(baseline, changed, None, startup=True) == "conflict"


def test_both_directions_and_no_feedback(session):
    session.startup()
    component = session.adapter.handler.get_vba_project().VBComponents("Module1")
    component.CodeModule.body = "Sub VBE()\nEnd Sub"
    settle(session)
    assert "Sub VBE()" in source(session).read_text()
    assert session.adapter.handler.doc.Save.call_count == 0
    change_file(session, "Sub Editor()\nEnd Sub")
    settle(session)
    assert component.CodeModule.body == "Sub Editor()\nEnd Sub"
    assert session.adapter.handler.doc.Save.call_count == 1
    settle(session)
    assert session.adapter.handler.doc.Save.call_count == 1
    assert not session.state["paused"]


def test_conflict_preserves_both_versions_and_other_modules_progress(session):
    cs = session.adapter.handler.get_vba_project().VBComponents
    cs.append(Component("Other"))
    session.startup()
    cs("Module1").CodeModule.body = "Office"
    change_file(session, "Filesystem")
    cs("Other").CodeModule.body = "Independent"
    settle(session)
    assert cs("Module1").CodeModule.body == "Office"
    assert "Filesystem" in source(session).read_text()
    assert "Independent" in source(session, "Other").read_text()
    artifacts = Path(session.state["paused"]["Module1"]["artifacts"])
    assert "Office" in (artifacts / "office.txt").read_text()
    assert "Filesystem" in (artifacts / "filesystem.txt").read_text()
    assert "Sub Example" in (artifacts / "baseline.txt").read_text()


def test_close_without_save_and_explicit_resolution(session):
    session.startup()
    cs = session.adapter.handler.get_vba_project().VBComponents
    cs("Module1").CodeModule.body = "Unsaved code"
    settle(session)
    cs("Module1").CodeModule.body = "Sub Example()\nEnd Sub"  # Reopened saved document.
    restarted = SyncSession(session.adapter)
    restarted.startup()
    assert "Module1" in restarted.state["paused"]
    assert "Unsaved code" in source(session).read_text()
    resolved = SyncSession(session.adapter, resolve="files", component="Module1")
    resolved.startup()
    assert cs("Module1").CodeModule.body == "Unsaved code"
    assert not resolved.state["paused"] and not resolved.state["pending"]


def test_file_deletion_does_not_delete_office(session):
    session.startup()
    source(session).unlink()
    settle(session)
    assert len(session.adapter.handler.get_vba_project().VBComponents) == 1
    assert "Module1" in session.state["paused"]


def test_save_failure_retries_without_reimport(session):
    session.startup()
    change_file(session, "Changed")
    handler = session.adapter.handler
    handler.doc.Save.side_effect = [OSError("save unavailable"), None]
    settle(session)
    assert session.state["pending"]["Module1"]["phase"] == "save"
    assert session.baseline("Module1").body != "Changed"
    settle(session)
    assert handler._update_module_content.call_count == 1
    assert handler.doc.Save.call_count == 2
    assert session.baseline("Module1").body == "Changed"
    assert not session.state["pending"]


def test_read_error_and_break_mode_are_not_deletions(session):
    session.startup()
    project = session.adapter.handler.get_vba_project()
    project.Mode = 1
    change_file(session, "Pending")
    settle(session)
    assert not session.state["paused"]
    assert session.read_failed
    project.Mode = 2
    session.scan()
    assert session.adapter.handler._update_module_content.call_count == 0
    session.scan()
    assert session.adapter.handler._update_module_content.call_count == 1


def test_member_attributes_are_retained_but_never_imported(session):
    cs = session.adapter.handler.get_vba_project().VBComponents
    cs("Module1").attributes = 'Attribute Example.VB_Description = "Description"'
    session.startup()
    assert "VB_Description" in source(session).read_text()
    source(session).write_text(source(session).read_text().replace("Sub Example", "Sub Changed"))
    settle(session)
    assert "Module1" in session.state["paused"]
    assert session.adapter.handler._update_module_content.call_count == 0


def test_export_encoding_loss_is_refused_before_writing(session):
    component = session.adapter.handler.get_vba_project().VBComponents("Module1")
    component.CodeModule.body = 'Sub Unicode()\n    Debug.Print "\U0001f642"\nEnd Sub'
    component.Export = lambda path: Path(path).write_text(
        'Attribute VB_Name = "Module1"\nSub Unicode()\n    Debug.Print "?"\nEnd Sub'
    )
    with pytest.raises(VBAError, match="encoding loss"):
        session.startup()
    assert not source(session).exists()


def test_ansi_unicode_code_round_trips(session):
    session.startup()
    change_file(session, 'Sub Unicode()\n    Debug.Print "Gr\u00fcsse"\nEnd Sub')
    settle(session)
    assert "Gr\u00fcsse" in session.baseline("Module1").body
    assert not session.state["paused"]


@pytest.mark.parametrize("name,kind", [("ThisWorkbook", 100), ("Sheet1", 100), ("ThisDocument", 100)])
def test_document_modules_are_updated_in_place(session, name, kind):
    h = session.adapter.handler
    if name == "ThisDocument":
        h.app_name = "Word"
        session = SyncSession(OfficeAdapter(h))
    h.get_vba_project().VBComponents[:] = [Component(name, kind)]
    session.startup()
    (h.vba_dir / f"{name}.cls").write_text(f'Attribute VB_Name = "{name}"\nUpdated\n')
    settle(session)
    assert h._update_document_module.call_count == 1
    assert len(h.get_vba_project().VBComponents) == 1


def test_new_standard_modules_on_either_side(session):
    session.startup()
    change_file(session, "New source", "NewSource")
    session.adapter.handler.get_vba_project().VBComponents.append(Component("NewOffice", body="New Office"))
    settle(session)
    assert "New Office" in source(session, "NewOffice").read_text()
    assert session.adapter.handler.get_vba_project().VBComponents("NewSource").CodeModule.body == "New source"


def test_separate_headers_and_rubberduck_recovery_exclusion(session):
    h = session.adapter.handler
    h.in_file_headers = False
    h.save_headers = True
    h.use_rubberduck_folders = True
    session = SyncSession(OfficeAdapter(h))
    session.startup()
    source(session).with_suffix(".header").write_text('Attribute VB_Name = "Another"')
    settle(session)
    assert "Module1" in session.state["paused"]
    assert session.adapter.handler._update_module_content.call_count == 0
    assert not list(session.directory.rglob("*.bas"))
    assert not list(session.directory.rglob("*.cls"))
    assert set(session.adapter.files()) == {"Module1"}


def test_source_move_is_paused_and_empty_body_is_valid(session):
    h = session.adapter.handler
    h.use_rubberduck_folders = True
    session = SyncSession(OfficeAdapter(h))
    session.startup()
    change_file(session, "")
    settle(session)
    assert h.get_vba_project().VBComponents("Module1").CodeModule.CountOfLines == 0
    folder = h.vba_dir / "Moved"
    folder.mkdir()
    source(session).rename(folder / "Module1.bas")
    settle(session)
    assert "Module1" in session.state["paused"]


def test_identity_mismatch_and_interrupted_operation_require_review(session):
    session.startup()
    session.state["pending"]["Module1"] = {"phase": "apply", "direction": "import"}
    session.persist()
    restarted = SyncSession(session.adapter)
    restarted.startup()
    assert "Module1" in restarted.state["paused"]
    session.adapter.handler.encoding = "utf-8"
    with pytest.raises(VBAError, match="mismatch"):
        SyncSession(session.adapter).startup()


def test_stale_input_cancels_mutation(session):
    session.startup()
    baseline = session.baseline("Module1")
    change_file(session, "First change")
    first = session.adapter.files()["Module1"]
    change_file(session, "Second change")
    session.transfer("import", "Module1", baseline, baseline, first)
    assert session.adapter.handler._update_module_content.call_count == 0
    assert not session.state["pending"]


def test_partial_import_failure_keeps_baseline_and_recovery(session):
    session.startup()
    baseline = session.baseline("Module1")
    change_file(session, "Requested source")
    component = session.adapter.handler.get_vba_project().VBComponents("Module1")

    def fail_after_clear(*_):
        component.CodeModule.body = ""
        raise OSError("COM insert failed")

    session.adapter.handler._update_module_content.side_effect = fail_after_clear
    settle(session)
    assert session.baseline("Module1") == baseline
    assert "Module1" in session.state["paused"]
    assert session.state["pending"]["Module1"]["phase"] == "apply"
    assert "Requested source" in source(session).read_text()


def test_unresolved_session_exits_nonzero_on_ctrl_c(session):
    change_file(session, "Divergent initial source")

    def interrupted(*_, **__):
        raise KeyboardInterrupt

    with patch("vba_edit.bidirectional.watch", interrupted), pytest.raises(VBAError, match="unresolved"):
        session.run()


def test_ctrl_c_during_startup_with_pending_work_exits_nonzero(session):
    def interrupted():
        session.state["pending"]["Module1"] = {"phase": "apply"}
        session.persist()
        raise KeyboardInterrupt

    with patch.object(session, "startup", interrupted), pytest.raises(VBAError, match="unresolved"):
        session.run()


def test_lock_excludes_another_session(tmp_path):
    with SessionLock(tmp_path / "session.bin"):
        with pytest.raises(VBAError, match="Another edit session"):
            with SessionLock(tmp_path / "session.bin"):
                pass
    with SessionLock(tmp_path / "session.bin"):
        pass


def test_access_saves_existing_module_and_refuses_new_modules(session):
    h = session.adapter.handler
    h.app_name = "Access"
    h.app = SimpleNamespace(
        CurrentDb=lambda: SimpleNamespace(Name=str(h.doc_path)),
        CurrentProject=SimpleNamespace(AllModules=SimpleNamespace(Item=Mock())),
        DoCmd=SimpleNamespace(Save=Mock()),
    )
    session = SyncSession(OfficeAdapter(h))
    session.startup()
    change_file(session, "Saved Access module")
    settle(session)
    h.app.DoCmd.Save.assert_called_once_with(5, "Module1")
    assert h.doc.Save.call_count == 0
    change_file(session, "New module", "NewSource")
    settle(session)
    assert "NewSource" in session.state["paused"]
    assert len(h.get_vba_project().VBComponents) == 1


def test_excluded_type_becomes_supported_without_overwriting_files(session):
    h = session.adapter.handler
    cs = h.get_vba_project().VBComponents
    cs.append(Component("ChangedType", kind=2))
    change_file(session, "Existing source", "ChangedType")
    session.startup()
    cs("ChangedType").Type = 1
    settle(session)
    assert "ChangedType" in session.state["paused"]
    assert "Existing source" in source(session, "ChangedType").read_text()


def test_duplicate_names_and_case_mismatch_suspend_reads(session):
    h = session.adapter.handler
    h.use_rubberduck_folders = True
    session = SyncSession(OfficeAdapter(h))
    session.startup()
    folder = h.vba_dir / "Nested"
    folder.mkdir()
    (folder / "Module1.bas").write_text("Duplicate")
    settle(session)
    assert session.read_failed
    assert not session.state["paused"]
    (folder / "Module1.bas").unlink()
    source(session).rename(h.vba_dir / "module1.bas")
    settle(session)
    assert session.read_failed
    assert h._update_module_content.call_count == 0


@pytest.mark.parametrize("host", ["excel", "word", "powerpoint", "access"])
def test_cli_registration_and_config(host):
    parser = OfficeVBACLI(host).create_cli_parser()
    args = parser.parse_args(["edit", "--bidirectional", "--sync-resolve", "files", "--sync-component", "M"])
    validate_bidirectional_options(args)
    assert args.bidirectional
    defaults = parser.parse_args(["edit"])
    merged = merge_config_with_args(defaults, {"general": {"bidirectional": True}}, cli_explicit=set())
    assert merged.bidirectional
    with pytest.raises(SystemExit):
        parser.parse_args(["export", "--bidirectional"])


@pytest.mark.parametrize(
    "options",
    [
        {"bidirectional": True, "with_references": True},
        {"bidirectional": True, "xlwings": True},
        {"bidirectional": True, "detect_encoding": True},
        {"sync_resolve": "files", "sync_component": "M"},
        {"bidirectional": True, "sync_resolve": "office"},
    ],
)
def test_unsupported_options(options):
    with pytest.raises(ValueError):
        validate_bidirectional_options(SimpleNamespace(**options))


def test_cli_skips_one_way_export_and_deletion_note(session, capsys):
    cli = OfficeVBACLI("excel")
    args = cli.create_cli_parser().parse_args(["edit", "--bidirectional"])
    with (
        patch("vba_edit.bidirectional.run_bidirectional") as run,
        patch.object(cli, "_call_handle_export_with_warnings") as export,
    ):
        cli._run_vba_command(session.adapter.handler, args)
        run.assert_called_once()
        export.assert_not_called()
    assert "File deletion pauses" in capsys.readouterr().out


def test_run_handles_ctrl_c_and_releases_locks(session):
    def watcher(*args, **kwargs):
        assert kwargs["yield_on_timeout"]
        assert not kwargs["watch_filter"](1, str(session.directory / "state.json"))
        yield set()
        raise KeyboardInterrupt

    with patch("vba_edit.bidirectional.watch", watcher):
        run_bidirectional(session.adapter.handler)
    with SessionLock(session.directory / "session.bin"):
        pass

"""--include / --exclude (issue #77) applied inside bidirectional edit mode."""

from vba_edit.bidirectional import Snapshot
from vba_edit.module_filter import ModuleFilter

# Reuse the deterministic fake-Office fixtures from the main bidirectional suite.
from tests.test_bidirectional import Component, change_file, session, settle, source  # noqa: F401


def _with_filter(session, **kwargs):
    session.adapter.handler.module_filter = ModuleFilter(**kwargs)
    return session


def test_excluded_component_is_never_exported_or_imported(session):
    _with_filter(session, exclude=["Vendor_*"])
    cs = session.adapter.handler.get_vba_project().VBComponents
    cs.append(Component("Vendor_Lib", body="Sub Vendor()\nEnd Sub"))

    session.startup()
    settle(session)

    assert source(session).exists()  # Module1 synced as usual
    assert not source(session, "Vendor_Lib").exists()
    assert "Vendor_Lib" not in session.state["baseline"]

    # A later edit to the excluded component on the Office side stays in Office only.
    cs("Vendor_Lib").CodeModule.body = "Sub Vendor()\n  ' changed\nEnd Sub"
    settle(session)
    assert not source(session, "Vendor_Lib").exists()
    assert "Vendor_Lib" not in session.state["paused"]


def test_excluded_file_on_disk_is_not_imported(session):
    _with_filter(session, exclude=["Vendor_*"])
    session.startup()
    change_file(session, "Sub New()\nEnd Sub", "Vendor_New")

    settle(session)
    settle(session)

    names = [c.Name for c in session.adapter.handler.get_vba_project().VBComponents]
    assert "Vendor_New" not in names
    assert "Vendor_New" not in session.state["paused"]


def test_include_limits_sync_to_matching_components(session):
    _with_filter(session, include=["Module1"])
    cs = session.adapter.handler.get_vba_project().VBComponents
    cs.append(Component("Other", body="Sub Other()\nEnd Sub"))

    session.startup()
    settle(session)

    assert source(session).exists()
    assert not source(session, "Other").exists()
    assert set(session.state["baseline"]) == {"Module1"}


def test_stale_baseline_of_excluded_component_does_not_pause(session):
    """A baseline left by an earlier unfiltered session must not be reported as a missing component."""
    cs = session.adapter.handler.get_vba_project().VBComponents
    cs.append(Component("Vendor_Lib", body="Sub Vendor()\nEnd Sub"))
    session.startup()
    settle(session)
    assert "Vendor_Lib" in session.state["baseline"]

    _with_filter(session, exclude=["Vendor_*"])
    session.startup()
    settle(session)

    assert "Vendor_Lib" not in session.state["paused"]
    assert "Vendor_Lib" in session.state["baseline"]  # left untouched, not deleted


def test_filtered_component_edits_do_not_feed_back(session):
    _with_filter(session, exclude=["Module1"])
    session.startup()
    settle(session)
    assert not source(session).exists()

    change_file(session, "Sub FromDisk()\nEnd Sub")
    settle(session)
    settle(session)

    office_body = session.adapter.handler.get_vba_project().VBComponents("Module1").CodeModule.body
    assert office_body == "Sub Example()\nEnd Sub"
    assert isinstance(Snapshot.parse("Module1", 1, "Module1.bas", source(session).read_text()), Snapshot)

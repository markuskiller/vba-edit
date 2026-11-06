"""Checks for the Office focus-loss detection. No Office application is needed."""

import os
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pytest

from vba_edit.focus_monitor import EVENT_SYSTEM_FOREGROUND, WindowFocusMonitor, office_window_handle

pytestmark = pytest.mark.skipif(os.name != "nt", reason="WinEvent hooks exist on Windows only")

OFFICE, OFFICE_OTHER_WINDOW, EDITOR, BROWSER = 101, 102, 201, 301
NAMES = {OFFICE: "excel.exe", OFFICE_OTHER_WINDOW: "excel.exe", EDITOR: "code.exe", BROWSER: "browser.exe"}


def start(office_hwnd=OFFICE):
    handler = SimpleNamespace(app_name="Excel", app=SimpleNamespace(Hwnd=office_hwnd))
    with patch.object(WindowFocusMonitor, "_executable_name", lambda self, hwnd: NAMES.get(hwnd)):
        return WindowFocusMonitor(handler)


@pytest.fixture
def monitor():
    """A monitor that believes Office has the foreground and hears only the events sent by the test."""
    try:
        instance = start()
    except OSError as exc:
        pytest.skip(f"WinEvent hook unavailable in this session: {exc}")
    instance._foreground_hwnd, instance._foreground_exe = OFFICE, NAMES[OFFICE]
    with (
        patch.object(WindowFocusMonitor, "_executable_name", lambda self, hwnd: NAMES.get(hwnd)),
        patch.object(instance, "pump_messages"),
    ):
        yield instance
    instance.stop()


def foreground(monitor, hwnd):
    monitor._on_foreground_change(None, EVENT_SYSTEM_FOREGROUND, hwnd, 0, 0, 0, 0)


def test_leaving_office_is_reported_once(monitor):
    foreground(monitor, EDITOR)
    assert monitor.office_lost_focus() is True
    assert monitor.office_lost_focus() is False


def test_switching_between_office_windows_is_not_a_loss(monitor):
    foreground(monitor, OFFICE_OTHER_WINDOW)
    assert monitor.office_lost_focus() is False


def test_returning_to_office_before_the_check_cancels_the_report(monitor):
    foreground(monitor, EDITOR)
    foreground(monitor, OFFICE)
    assert monitor.office_lost_focus() is False


def test_switching_between_other_programs_is_ignored(monitor):
    foreground(monitor, EDITOR)
    assert monitor.office_lost_focus() is True
    foreground(monitor, BROWSER)
    assert monitor.office_lost_focus() is False


def test_unknown_foreground_process_counts_as_leaving_office(monitor):
    foreground(monitor, 999)
    assert monitor.office_lost_focus() is True


def test_real_hook_installs_pumps_and_stops():
    try:
        instance = start()
    except OSError as exc:
        pytest.skip(f"WinEvent hook unavailable in this session: {exc}")
    try:
        assert instance._hook
        instance.pump_messages()
    finally:
        instance.stop()
        instance.stop()
    assert instance._hook is None


def test_unidentified_office_process_fails_without_installing_a_hook():
    with patch("vba_edit.focus_monitor._load_win32") as load:
        user32 = Mock()
        load.return_value = (user32, Mock(), Mock())
        with pytest.raises(OSError, match="Cannot identify"):
            start(office_hwnd=999)
        user32.SetWinEventHook.assert_not_called()


def test_office_window_handle_per_host():
    assert office_window_handle(SimpleNamespace(app_name="Excel", app=SimpleNamespace(Hwnd=1))) == 1
    word = SimpleNamespace(app_name="Word", app=SimpleNamespace(ActiveWindow=SimpleNamespace(Hwnd=2)))
    assert office_window_handle(word) == 2
    access = SimpleNamespace(app_name="Access", app=SimpleNamespace(hWndAccessApp=lambda: 3))
    assert office_window_handle(access) == 3
    with patch("win32gui.FindWindow", return_value=4) as find:
        assert office_window_handle(SimpleNamespace(app_name="PowerPoint")) == 4
        find.assert_called_once_with("PPTFrameClass", None)
    with pytest.raises(ValueError):
        office_window_handle(SimpleNamespace(app_name="Outlook"))

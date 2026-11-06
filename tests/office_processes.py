"""Find and close the Office applications that a test session started.

Office applications that were already running when the session began belong to
the person at the keyboard. They are recorded once at startup and never closed.
"""

import os
import subprocess

import win32ts

OFFICE_PROCESS_NAMES = frozenset({"EXCEL.EXE", "WINWORD.EXE", "POWERPNT.EXE", "MSACCESS.EXE"})


def running_office_pids() -> dict[int, str]:
    """Map the process ID of each Office application in this Windows session to its image name."""
    own_session = win32ts.ProcessIdToSessionId(os.getpid())
    processes = win32ts.WTSEnumerateProcesses(win32ts.WTS_CURRENT_SERVER_HANDLE)
    return {
        pid: name.upper()
        for session, pid, name, _sid in processes
        if session == own_session and name.upper() in OFFICE_PROCESS_NAMES
    }


def office_pids_at_session_start() -> dict[int, str] | None:
    """Return the Office processes running now, or None if they cannot be listed."""
    try:
        return running_office_pids()
    except Exception:
        return None


def close_office_apps_started_since(pids_before) -> list[str]:
    """Force-close Office processes that are not in ``pids_before`` and return what was closed.

    Nothing is closed when ``pids_before`` is None, because it is then unknown
    which applications the tests started.
    """
    if pids_before is None:
        return []
    try:
        running = running_office_pids()
    except Exception:
        return []

    closed = []
    for pid, name in running.items():
        if pid in pids_before:
            continue
        try:
            # /T also ends child processes of the application.
            result = subprocess.run(
                ["taskkill", "/F", "/PID", str(pid), "/T"], capture_output=True, text=True, timeout=10
            )
        except (OSError, subprocess.SubprocessError):
            continue
        if result.returncode == 0:
            closed.append(f"{name} (PID {pid})")
    return closed

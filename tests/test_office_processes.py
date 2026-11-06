"""Tests for the cleanup that closes only the Office applications the tests started."""

import subprocess

from tests import office_processes


def test_lists_only_office_processes_of_the_own_windows_session(monkeypatch):
    monkeypatch.setattr(office_processes.win32ts, "ProcessIdToSessionId", lambda _pid: 2)
    monkeypatch.setattr(
        office_processes.win32ts,
        "WTSEnumerateProcesses",
        lambda _server: [
            (2, 100, "EXCEL.EXE", None),
            (2, 101, "winword.exe", None),
            (2, 102, "OUTLOOK.EXE", None),
            (2, 103, "python.exe", None),
            (3, 104, "EXCEL.EXE", None),  # Another user's session
        ],
    )

    assert office_processes.running_office_pids() == {100: "EXCEL.EXE", 101: "WINWORD.EXE"}


def test_closes_only_applications_started_after_the_session_began(monkeypatch):
    commands = []

    def fake_run(command, **_kwargs):
        commands.append(command)
        return subprocess.CompletedProcess(command, 0, stdout="SUCCESS", stderr="")

    monkeypatch.setattr(office_processes, "running_office_pids", lambda: {100: "EXCEL.EXE", 200: "EXCEL.EXE"})
    monkeypatch.setattr(office_processes.subprocess, "run", fake_run)

    closed = office_processes.close_office_apps_started_since({100: "EXCEL.EXE"})

    assert commands == [["taskkill", "/F", "/PID", "200", "/T"]]
    assert closed == ["EXCEL.EXE (PID 200)"]


def test_closes_nothing_when_the_starting_state_is_unknown(monkeypatch):
    monkeypatch.setattr(office_processes, "running_office_pids", lambda: {200: "EXCEL.EXE"})
    monkeypatch.setattr(office_processes.subprocess, "run", lambda *_a, **_k: 1 / 0)

    assert office_processes.close_office_apps_started_since(None) == []


def test_closes_nothing_when_processes_cannot_be_listed(monkeypatch):
    def broken():
        raise OSError("enumeration failed")

    monkeypatch.setattr(office_processes, "running_office_pids", broken)
    monkeypatch.setattr(office_processes.subprocess, "run", lambda *_a, **_k: 1 / 0)

    assert office_processes.office_pids_at_session_start() is None
    assert office_processes.close_office_apps_started_since({}) == []


def test_reports_only_processes_that_were_actually_closed(monkeypatch):
    def fake_run(command, **_kwargs):
        return subprocess.CompletedProcess(command, 128, stdout="", stderr="not found")

    monkeypatch.setattr(office_processes, "running_office_pids", lambda: {200: "WINWORD.EXE"})
    monkeypatch.setattr(office_processes.subprocess, "run", fake_run)

    assert office_processes.close_office_apps_started_since({}) == []

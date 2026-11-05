"""Detect when the Office application gives up the foreground (Windows only).

A WinEvent hook reports foreground-window changes. The callback only records
that Office lost focus; it never calls into Office, so all COM work stays with
the caller. Events are delivered while this thread's message queue is pumped,
so ``office_lost_focus()`` must be called regularly from the thread that
created the monitor.

Based on the WindowFocusMonitor contributed by russ-who in PR #123.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wintypes
import logging
import os

logger = logging.getLogger(__name__)

EVENT_SYSTEM_FOREGROUND = 0x0003
WINEVENT_OUTOFCONTEXT = 0x0000
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PM_REMOVE = 0x0001


def office_window_handle(handler) -> int:
    """Return a top-level window handle of the handler's Office application."""
    app_name = handler.app_name
    if app_name == "Access":
        return handler.app.hWndAccessApp()
    if app_name == "Excel":
        return handler.app.Hwnd
    if app_name == "Word":
        return handler.app.ActiveWindow.Hwnd
    if app_name == "PowerPoint":
        import win32gui

        return win32gui.FindWindow("PPTFrameClass", None)
    raise ValueError(f"Unexpected Office application ({app_name})")


def _load_win32():
    """Load private user32/kernel32 handles with pointer-sized signatures declared."""
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(
        None,
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.HWND,
        wintypes.LONG,
        wintypes.LONG,
        wintypes.DWORD,
        wintypes.DWORD,
    )
    message = ctypes.POINTER(wintypes.MSG)

    user32.GetForegroundWindow.argtypes = []
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.SetWinEventHook.argtypes = [
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HMODULE,
        callback_type,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.DWORD,
    ]
    user32.SetWinEventHook.restype = wintypes.HANDLE
    user32.UnhookWinEvent.argtypes = [wintypes.HANDLE]
    user32.UnhookWinEvent.restype = wintypes.BOOL
    user32.PeekMessageW.argtypes = [message, wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.UINT]
    user32.PeekMessageW.restype = wintypes.BOOL
    user32.TranslateMessage.argtypes = [message]
    user32.TranslateMessage.restype = wintypes.BOOL
    user32.DispatchMessageW.argtypes = [message]
    user32.DispatchMessageW.restype = ctypes.c_ssize_t

    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    return user32, kernel32, callback_type


class WindowFocusMonitor:
    """Report when the foreground moves from the Office application to another program.

    Switching between windows of the same executable (for example from the VBE
    to the workbook) is not a loss of focus.

    Args:
        handler: The OfficeVBAHandler whose application should be observed.

    Raises:
        Exception: If the Office process cannot be identified or the hook
            cannot be installed. Nothing stays registered in that case.
    """

    def __init__(self, handler):
        self._user32, self._kernel32, callback_type = _load_win32()
        self._hook = None
        self._lost_focus = False

        self.office_exe = self._executable_name(office_window_handle(handler))
        if self.office_exe is None:
            raise OSError("Cannot identify the Office process")

        # Remember what has the foreground now so the first change is judged correctly.
        self._foreground_hwnd = self._user32.GetForegroundWindow()
        self._foreground_exe = self._executable_name(self._foreground_hwnd)
        logger.debug("Initial foreground application: %s", self._foreground_exe)

        # The callback object must stay referenced for as long as the hook exists.
        self._callback = callback_type(self._on_foreground_change)
        # Installed last: nothing after this point can fail and leave the hook behind.
        self._hook = self._user32.SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, None, self._callback, 0, 0, WINEVENT_OUTOFCONTEXT
        )
        if not self._hook:
            raise ctypes.WinError(ctypes.get_last_error())

    def office_lost_focus(self) -> bool:
        """Pump pending events; True once after Office gave up the foreground and stayed away."""
        self.pump_messages()
        lost, self._lost_focus = self._lost_focus, False
        return lost

    def pump_messages(self) -> None:
        """Dispatch this thread's queued window messages, which delivers the hook callbacks."""
        msg = wintypes.MSG()
        while self._user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, PM_REMOVE):
            self._user32.TranslateMessage(ctypes.byref(msg))
            self._user32.DispatchMessageW(ctypes.byref(msg))

    def stop(self) -> None:
        """Remove the hook. Safe to call more than once."""
        if self._hook:
            self._user32.UnhookWinEvent(self._hook)
            self._hook = None

    def _on_foreground_change(self, _hook, event, hwnd, _id_object, _id_child, _thread, _time) -> None:
        """Called on a foreground change. Record whether Office lost or regained focus."""
        try:
            if event != EVENT_SYSTEM_FOREGROUND or not hwnd or hwnd == self._foreground_hwnd:
                return
            exe_name = self._executable_name(hwnd)
            if exe_name != self._foreground_exe:
                logger.debug("New foreground application: %s", exe_name)
                if self._foreground_exe == self.office_exe:
                    self._lost_focus = True
                elif exe_name == self.office_exe:
                    # Back in Office before anyone asked: its code is being edited again.
                    self._lost_focus = False
            self._foreground_hwnd = hwnd
            self._foreground_exe = exe_name
        except Exception as exc:  # An exception must never escape into the Win32 dispatcher.
            logger.debug("Ignoring focus event: %s", exc)

    def _executable_name(self, hwnd) -> str | None:
        """Return the case-folded executable file name owning a window, or None if unknown."""
        if not hwnd:
            return None
        pid = wintypes.DWORD()
        self._user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        process = self._kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
        if not process:
            return None  # System or protected process.
        try:
            size = wintypes.DWORD(1024)
            buffer = ctypes.create_unicode_buffer(size.value)
            if self._kernel32.QueryFullProcessImageNameW(process, 0, buffer, ctypes.byref(size)):
                return os.path.basename(buffer.value).casefold()
            return None
        finally:
            self._kernel32.CloseHandle(process)

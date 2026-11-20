"""Take the screenshots used in the docs. Each one is a capture of a single window this script opened.

- config-gen.png : the Config Generator (`excel-vba config-gen`) with a sample configuration loaded
- vbe-synced.png : the VBA editor of the sample workbook after an edit was synced from a source file

The windows appear on the desktop for a few seconds; nothing else on the screen is captured
(PrintWindow renders the one window, not the screen). Requires Excel and the dev environment.

    python capture_screenshots.py <new work dir> [config-gen] [vbe]
"""

from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import win32gui
import win32ui
from PIL import Image

HERE = Path(__file__).resolve().parent
IMAGES = HERE.parent / "public" / "assets" / "img"
PW_RENDERFULLCONTENT = 2


def capture(hwnd: int, target: Path, max_width: int | None = None) -> None:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    window_dc = win32gui.GetWindowDC(hwnd)
    source = win32ui.CreateDCFromHandle(window_dc)
    memory = source.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source, width, height)
    memory.SelectObject(bitmap)
    ctypes.windll.user32.PrintWindow(hwnd, memory.GetSafeHdc(), PW_RENDERFULLCONTENT)
    info = bitmap.GetInfo()
    image = Image.frombuffer(
        "RGB", (info["bmWidth"], info["bmHeight"]), bitmap.GetBitmapBits(True), "raw", "BGRX", 0, 1
    )
    win32gui.DeleteObject(bitmap.GetHandle())
    memory.DeleteDC()
    source.DeleteDC()
    win32gui.ReleaseDC(hwnd, window_dc)
    if max_width and image.width > max_width:
        image = image.resize((max_width, round(image.height * max_width / image.width)), Image.LANCZOS)
    target.parent.mkdir(parents=True, exist_ok=True)
    image.save(target, optimize=True)
    print("saved", target, image.size)


def config_gen(work: Path) -> None:
    from tkinter import filedialog

    from vba_edit.config_gen import ConfigGenApp

    sample = work / "vba-config.toml"
    sample.write_text(
        "[general]\n"
        "file = 'Budget.xlsm'\n"
        "vba_directory = '{file.path}\\src'\n"
        "in_file_headers = true\n"
        "rubberduck_folders = true\n"
        "save_metadata = true\n"
        'exclude = ["Vendor_*", "Sheet*"]\n',
        encoding="utf-8",
    )
    window = ConfigGenApp("excel")
    try:
        original = filedialog.askopenfilename
        filedialog.askopenfilename = lambda **_kwargs: str(sample)
        try:
            window._load_config()
        finally:
            filedialog.askopenfilename = original
        window.geometry("1700x1580")  # a little taller than the default so the button row is not cut off
        for _ in range(20):
            window.update_idletasks()
            window.update()
            time.sleep(0.05)
        capture(int(window.wm_frame(), 16), IMAGES / "config-gen.png", max_width=1200)
    finally:
        window.destroy()


def vbe(work: Path) -> None:
    import pythoncom
    import win32com.client

    import demo_project

    path = demo_project.build(work / "vbe")
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        excel.Visible = True
        excel.WindowState = -4140  # xlMinimized: only the editor window matters
        book = excel.Workbooks.Open(str(path))
        component = book.VBProject.VBComponents("Reporting")
        code = component.CodeModule
        text = code.Lines(1, code.CountOfLines).replace('"Total: "', '"Total spend: "')
        code.DeleteLines(1, code.CountOfLines)
        code.AddFromString(text)
        editor = excel.VBE
        editor.MainWindow.Visible = True
        editor.MainWindow.WindowState = 0  # vbext_ws_Normal
        hwnd = editor.MainWindow.HWnd
        # Sizes are physical pixels (the process is DPI aware); the capture is scaled down afterwards.
        win32gui.MoveWindow(hwnd, 80, 80, 2400, 1440, True)
        for tool_window in editor.Windows:
            # Watch, Locals, Immediate and Properties windows only take room away from the code.
            if int(tool_window.Type) in (3, 4, 5, 7):
                tool_window.Visible = False
        code.CodePane.Show()
        code.CodePane.Window.WindowState = 2  # maximised inside the editor
        code.CodePane.SetSelection(8, 5, 8, 5)
        time.sleep(2.5)
        capture(hwnd, IMAGES / "vbe-synced.png", max_width=1200)
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    work = Path(sys.argv[1]).resolve()
    work.mkdir(parents=True)
    wanted = sys.argv[2:] or ["config-gen", "vbe"]
    if "config-gen" in wanted:
        config_gen(work)
    if "vbe" in wanted:
        vbe(work)

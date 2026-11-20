"""Run one vba-edit command for the cast recorder.

Two things differ from a real terminal session, neither changes what the tool prints:
- rich is told to write ANSI colours although stdout is a pipe (it would otherwise use the
  Windows console API, which leaves nothing to capture);
- Ctrl+C is delivered when the file named in CAST_STOP appears, so `edit` can be ended on cue.

    python cli_runner.py excel export -f Budget.xlsm
"""

import _thread
import importlib
import os
import sys
import threading
import time
from pathlib import Path

import rich.console

_original_init = rich.console.Console.__init__


def _ansi_console(self, *args, **kwargs):
    kwargs.update(
        force_terminal=True,
        legacy_windows=False,
        color_system="standard",
        width=int(os.environ.get("CAST_WIDTH", "100")),
    )
    _original_init(self, *args, **kwargs)


rich.console.Console.__init__ = _ansi_console


def _interrupt_when_asked(stop_file: Path) -> None:
    while not stop_file.exists():
        time.sleep(0.2)
    _thread.interrupt_main()


if __name__ == "__main__":
    tool, arguments = sys.argv[1], sys.argv[2:]
    if os.environ.get("CAST_STOP"):
        threading.Thread(target=_interrupt_when_asked, args=(Path(os.environ["CAST_STOP"]),), daemon=True).start()
    sys.argv = [f"{tool}-vba", *arguments]
    importlib.import_module(f"vba_edit.{tool}_vba").main()

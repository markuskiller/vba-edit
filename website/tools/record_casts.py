"""Record the terminal sequences shown in the docs by running the real commands.

Every command below is executed against a disposable copy of the sample project (see
demo_project.py) and its output is captured as printed. Only two things are edited afterwards:
the working directory is shown as C:\\Projects\\budget, and long idle waits are shortened.
Requires Windows, Excel with "Trust access to the VBA project object model", git and PowerShell.
Excel windows open and close on the desktop while this runs.

    python record_casts.py <new work dir> [scene ...]

Writes ../public/assets/casts.js (scenes not named on the command line are kept as they are).
"""

from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

import demo_project

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
OUTPUT = HERE.parent / "public" / "assets" / "casts.js"
PYTHON = sys.executable
SHOWN_ROOT = "C:\\Projects\\budget"
SGR = re.compile(r"\x1b\[([0-9;]*)m")
OTHER_ESCAPES = re.compile(r"\x1b\[[0-9;?]*[A-Za-ln-z]")
COLOURS = {30: "k", 31: "r", 32: "g", 33: "y", 34: "b", 35: "m", 36: "c", 37: "w"}


def ansi_to_html(text: str) -> str:
    """Translate SGR colour codes into <span class="..."> runs; everything else is escaped."""
    text = OTHER_ESCAPES.sub("", text.replace("\r\n", "\n").replace("\r", ""))
    out, classes, position = [], set(), 0

    def emit(chunk: str) -> None:
        # A span never crosses a line break, so the player can reveal the output line by line.
        for index, piece in enumerate(chunk.split("\n")):
            if index:
                out.append("\n")
            if piece:
                escaped = html.escape(piece, quote=False)
                out.append(f'<span class="{" ".join(sorted(classes))}">{escaped}</span>' if classes else escaped)

    for match in SGR.finditer(text):
        emit(text[position : match.start()])
        position = match.end()
        for code in [int(c) for c in match.group(1).split(";") if c] or [0]:
            if code == 0:
                classes.clear()
            elif code == 1:
                classes.add("t-bold")
            elif code == 2:
                classes.add("t-dim")
            elif code in (22,):
                classes.discard("t-bold")
                classes.discard("t-dim")
            elif code == 39:
                classes -= {c for c in classes if c.startswith("t-fg-")}
            elif code in COLOURS or code - 60 in COLOURS:
                classes -= {c for c in classes if c.startswith("t-fg-")}
                classes.add("t-fg-" + COLOURS.get(code, COLOURS.get(code - 60, "w")) + ("b" if code >= 90 else ""))
    emit(text[position:])
    return "".join(out)


class Recorder:
    def __init__(self, work: Path):
        self.work = work
        self.steps: list[dict] = []
        self.environment = {
            **os.environ,
            "PYTHONPATH": str(REPO / "src"),
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUTF8": "1",
            "CAST_WIDTH": "400",
            "GIT_AUTHOR_NAME": "Sample",
            "GIT_AUTHOR_EMAIL": "sample@example.com",
            "GIT_COMMITTER_NAME": "Sample",
            "GIT_COMMITTER_EMAIL": "sample@example.com",
            "GIT_CONFIG_GLOBAL": os.devnull,
            "GIT_CONFIG_NOSYSTEM": "1",
        }
        self.environment.pop("NO_COLOR", None)

    # -- helpers -------------------------------------------------------------------------------
    def clean(self, text: str) -> str:
        """Show the working directory as SHOWN_ROOT (colour codes may sit anywhere inside the path)."""
        gap = r"(?:\x1b\[[0-9;]*m)*"
        for root in {str(self.work), self.work.as_posix()}:
            pattern = gap.join(re.escape(character) for character in root)
            text = re.sub(pattern, lambda _m: SHOWN_ROOT, text, flags=re.IGNORECASE)
        return text

    def out(self, text: str) -> None:
        text = self.clean(text.replace("\r\n", "\n").replace("\r", ""))
        text = re.sub(r"[ \t]+((?:\x1b\[[0-9;]*m)*)$", r"\1", text, flags=re.MULTILINE).rstrip("\n")
        if text.strip():
            self.steps.append({"out": ansi_to_html(text)})

    def note(self, text: str) -> None:
        self.steps.append({"note": text})

    def command(self, shown: str) -> None:
        self.steps.append({"cmd": shown})

    def vba(self, shown: str, *arguments: str, tool: str = "excel", stdin: str | None = None) -> int:
        """Run `<tool>-vba <arguments>` to completion."""
        self.command(shown)
        done = subprocess.run(
            [PYTHON, "-X", "utf8", str(HERE / "cli_runner.py"), tool, *arguments],
            cwd=self.work,
            env=self.environment,
            input=stdin,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=300,
        )
        output = done.stdout + done.stderr
        prompt = re.search(r"\[[yY]/[nN]\]: ?", SGR.sub("", output)) if stdin else None
        if prompt:
            # The answer is not echoed into a pipe: split the output where it was typed.
            marker = output.index("]: ", output.index("/")) + 3
            self.out(output[:marker])
            self.steps[-1]["prompt"] = True
            self.steps.append({"typed": stdin.strip()})
            self.out(output[marker:])
        else:
            self.out(output)
        return done.returncode

    def shell(self, shown: str, argv: list[str]) -> None:
        self.command(shown)
        done = subprocess.run(
            argv,
            cwd=self.work,
            env=self.environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        self.out(done.stdout + done.stderr)

    def powershell(self, shown: str) -> None:
        self.shell(shown, ["powershell", "-NoProfile", "-Command", "$ProgressPreference='SilentlyContinue'; " + shown])

    def git(self, shown: str) -> None:
        arguments = shown.split(" ", 1)[1]
        self.shell(
            shown,
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "git -c color.ui=always -c core.autocrlf=false -c advice.defaultBranchName=false " + arguments,
            ],
        )

    def watch(self, shown: str, *arguments: str, tool: str = "excel") -> "Watch":
        self.command(shown)
        return Watch(self, tool, arguments)


class Watch:
    """A running `edit` session whose output is collected while the script acts on files and Office."""

    def __init__(self, recorder: Recorder, tool: str, arguments: tuple[str, ...]):
        self.recorder = recorder
        self.stop_file = recorder.work / f".stop-{time.time_ns()}"
        self.buffer: list[str] = []
        self.taken = 0
        self.last_output = time.monotonic()
        self.process = subprocess.Popen(
            [PYTHON, "-X", "utf8", "-u", str(HERE / "cli_runner.py"), tool, *arguments],
            cwd=recorder.work,
            env={**recorder.environment, "CAST_STOP": str(self.stop_file), "PYTHONUNBUFFERED": "1"},
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        self.reader = threading.Thread(target=self._read, daemon=True)
        self.reader.start()

    def _read(self) -> None:
        for line in self.process.stdout:
            self.buffer.append(line)
            self.last_output = time.monotonic()

    def settle(self, quiet: float = 4.0, limit: float = 90.0, expect: str | None = None) -> None:
        """Wait until nothing was printed for `quiet` seconds (and `expect` was seen), then record the output."""
        begin = time.monotonic()
        end = begin + limit

        def seen() -> str:
            return SGR.sub("", "".join(self.buffer[self.taken :]))

        while time.monotonic() < end and self.process.poll() is None:
            idle = time.monotonic() - max(self.last_output, begin)
            if idle >= quiet and (expect is None or expect in seen()):
                break
            time.sleep(0.2)
        else:
            if expect is not None and expect not in seen():
                self.recorder.out("".join(self.buffer[self.taken :]))
                raise AssertionError(f"expected output not seen: {expect!r}")
        text, self.taken = "".join(self.buffer[self.taken :]), len(self.buffer)
        self.recorder.out(text)

    def stop(self) -> int:
        self.stop_file.touch()
        try:
            self.process.wait(timeout=60)
        except subprocess.TimeoutExpired:
            self.process.kill()
        self.reader.join(timeout=10)
        self.recorder.steps.append({"key": "Ctrl+C"})
        text, self.taken = "".join(self.buffer[self.taken :]), len(self.buffer)
        self.recorder.out(text)
        self.stop_file.unlink(missing_ok=True)
        return self.process.returncode


# -- Office side helpers (the "you, in the VBA editor" part of a scene) ---------------------------
def open_workbook(path: Path):
    import win32com.client

    return win32com.client.GetObject(str(path))


def close_excel(path: Path) -> None:
    """Close the sample workbook and quit its Excel if nothing else is open in it."""
    import pythoncom
    import win32com.client

    try:
        book = win32com.client.GetObject(str(path))
        application = book.Application
        application.DisplayAlerts = False
        book.Close(SaveChanges=False)
        if application.Workbooks.Count == 0:
            application.Quit()
    except pythoncom.com_error:
        pass
    time.sleep(2)


def replace_in_module(path: Path, module: str, old: str, new: str) -> None:
    component = open_workbook(path).VBProject.VBComponents(module)
    code = component.CodeModule
    text = code.Lines(1, code.CountOfLines)
    assert old in text, (module, old)
    code.DeleteLines(1, code.CountOfLines)
    code.AddFromString(text.replace(old, new))


def replace_in_file(path: Path, old: str, new: str, encoding: str = "cp1252") -> None:
    data = path.read_bytes()
    assert old.encode(encoding) in data, (path, old)
    path.write_bytes(data.replace(old.encode(encoding), new.encode(encoding)))


# -- scenes ---------------------------------------------------------------------------------------
RECORDERS: list[Recorder] = []


def fresh(work: Path, name: str, template: Path) -> Recorder:
    directory = work / name
    directory.mkdir(parents=True)
    shutil.copy2(template, directory / "Budget.xlsm")
    recorder = Recorder(directory)
    RECORDERS.append(recorder)
    return recorder


def scene_check(work, plain, forms):
    r = fresh(work, "check", plain)
    r.vba("excel-vba check", "check")
    return r.steps


def scene_export(work, plain, forms):
    r = fresh(work, "export", plain)
    r.vba(
        "excel-vba export -f Budget.xlsm --vba-directory src", "export", "-f", "Budget.xlsm", "--vba-directory", "src"
    )
    r.powershell("Get-ChildItem src -Name")
    r.powershell("Get-Content src\\Reporting.bas")
    return r.steps


def scene_edit(work, plain, forms):
    r = fresh(work, "edit", plain)
    book = r.work / "Budget.xlsm"
    session = r.watch(
        "excel-vba edit -f Budget.xlsm --vba-directory src", "edit", "-f", "Budget.xlsm", "--vba-directory", "src"
    )
    try:
        session.settle(quiet=3, expect="Watching for changes")
        r.note("In your editor: change the message in src\\Reporting.bas and press Ctrl+S")
        replace_in_file(r.work / "src" / "Reporting.bas", '"Total: "', '"Total spend: "')
        session.settle(quiet=3, expect="Successfully processed")
        r.note("Add a new file src\\Forecast.bas and save it")
        (r.work / "src" / "Forecast.bas").write_bytes(
            b"Option Explicit\r\n\r\nPublic Function NextMonth(ByVal last As Currency) As Currency\r\n"
            b"    NextMonth = last * 1.03\r\nEnd Function\r\n"
        )
        session.settle(quiet=3, expect="Successfully processed")
        session.stop()
        code = open_workbook(book).VBProject.VBComponents("Reporting").CodeModule
        assert "Total spend" in code.Lines(1, code.CountOfLines)
    finally:
        if session.process.poll() is None:
            session.process.kill()
        close_excel(book)
    return r.steps


def scene_git(work, plain, forms):
    r = fresh(work, "git", plain)
    r.vba(
        "excel-vba export -f Budget.xlsm --vba-directory src", "export", "-f", "Budget.xlsm", "--vba-directory", "src"
    )
    r.git("git init -q -b main")
    r.git("git add src")
    r.git('git commit -q -m "Baseline export of Budget.xlsm"')
    r.note("... a morning of editing with `excel-vba edit` running ...")
    replace_in_file(r.work / "src" / "Helpers.bas", '"#,##0.00"', '"#,##0.00 \\C\\H\\F"')
    replace_in_file(r.work / "src" / "Reporting.bas", '"Total: "', '"Total spend: "')
    r.git("git status --short")
    r.git("git diff src/Helpers.bas")
    r.git('git commit -q -am "Show totals in CHF"')
    r.git("git log --oneline")
    return r.steps


def scene_headers(work, plain, forms):
    r = fresh(work, "headers", forms)
    r.vba(
        "excel-vba export -f Budget.xlsm --vba-directory src --in-file-headers --rubberduck-folders",
        "export",
        "-f",
        "Budget.xlsm",
        "--vba-directory",
        "src",
        "--in-file-headers",
        "--rubberduck-folders",
    )
    r.powershell("Get-ChildItem src -Recurse -Name -File")
    r.powershell("Get-Content src\\Domain\\Invoice.cls -TotalCount 12")
    return r.steps


def scene_safety(work, plain, forms):
    r = fresh(work, "safety", forms)
    base = ("export", "-f", "Budget.xlsm", "--vba-directory", "src")
    shown = "excel-vba export -f Budget.xlsm --vba-directory src"
    r.vba(shown + " --save-headers", *base, "--save-headers")
    r.note("Later: the same export again, files already exist. Answer n to keep them")
    r.vba(shown + " --save-headers", *base, "--save-headers", stdin="n\n")
    r.note("Switch the header mode and answer y: files are rewritten, old .header files removed")
    r.vba(shown + " --in-file-headers", *base, "--in-file-headers", stdin="y\n")
    r.powershell("Get-ChildItem src -Name")
    return r.steps


def scene_filter(work, plain, forms):
    r = fresh(work, "filter", plain)
    r.vba(
        "excel-vba export -f Budget.xlsm --vba-directory src", "export", "-f", "Budget.xlsm", "--vba-directory", "src"
    )
    r.note("Edit Reporting.bas and Helpers.bas offline, then push only those two back")
    replace_in_file(r.work / "src" / "Reporting.bas", '"Total: "', '"Total spend: "')
    r.vba(
        'excel-vba import -f Budget.xlsm --vba-directory src --include "Rep*" Helpers.bas',
        "import",
        "-f",
        "Budget.xlsm",
        "--vba-directory",
        "src",
        "--include",
        "Rep*",
        "Helpers.bas",
    )
    r.vba(
        'excel-vba export -f Budget.xlsm --vba-directory modules-only --exclude "Sheet*" ThisWorkbook',
        "export",
        "-f",
        "Budget.xlsm",
        "--vba-directory",
        "modules-only",
        "--exclude",
        "Sheet*",
        "ThisWorkbook",
    )
    r.powershell("Get-ChildItem modules-only -Name")
    return r.steps


def scene_references(work, plain, forms):
    r = fresh(work, "references", plain)
    r.vba("excel-vba references list -f Budget.xlsm", "references", "list", "-f", "Budget.xlsm")
    r.vba(
        "excel-vba references list -f Budget.xlsm --no-default",
        "references",
        "list",
        "-f",
        "Budget.xlsm",
        "--no-default",
    )
    r.vba(
        "excel-vba references export -f Budget.xlsm -r budget-refs.toml --no-default",
        "references",
        "export",
        "-f",
        "Budget.xlsm",
        "-r",
        "budget-refs.toml",
        "--no-default",
    )
    r.powershell("Get-Content budget-refs.toml")
    r.vba("excel-vba references validate -f Budget.xlsm", "references", "validate", "-f", "Budget.xlsm")
    return r.steps


def scene_config(work, plain, forms):
    r = fresh(work, "config", plain)
    (r.work / "vba-config.toml").write_text(
        '[general]\nfile = "Budget.xlsm"\nvba_directory = "{file.name}-vba"\n'
        'in_file_headers = true\nrubberduck_folders = true\nskip_empty = true\nexclude = ["Sheet*"]\n',
        encoding="utf-8",
    )
    r.powershell("Get-Content vba-config.toml")
    r.vba("excel-vba export --conf vba-config.toml", "export", "--conf", "vba-config.toml")
    r.powershell("Get-ChildItem Budget-vba -Recurse -Name -File")
    return r.steps


def scene_bidirectional(work, plain, forms):
    r = fresh(work, "bidirectional", plain)
    book = r.work / "Budget.xlsm"
    shown = "excel-vba edit -f Budget.xlsm --vba-directory src --bidirectional --in-file-headers"
    arguments = ("edit", "-f", "Budget.xlsm", "--vba-directory", "src", "--bidirectional", "--in-file-headers")
    session = r.watch(shown, *arguments)
    try:
        session.settle(quiet=6, expect="Bidirectional edit")
        r.note("In the VBA editor: change the message in Reporting, then switch to your editor")
        replace_in_module(book, "Reporting", '"Total: "', '"Total spend: "')
        session.settle(quiet=4, expect="Exported Reporting")
        r.note("In your editor: change the number format in src\\Helpers.bas and press Ctrl+S")
        replace_in_file(r.work / "src" / "Helpers.bas", '"#,##0.00"', '"#,##0"')
        session.settle(quiet=4, expect="Imported Helpers")
        r.note("Now change Reporting on BOTH sides before the next sync")
        replace_in_module(book, "Reporting", '"Total spend: "', '"Total (VBE): "')
        replace_in_file(r.work / "src" / "Reporting.bas", '"Total spend: "', '"Total (editor): "')
        session.settle(quiet=4, expect="PAUSED")
        code = session.stop()
        r.steps.append({"exit": code})
        session = r.watch(
            shown + " --sync-resolve files --sync-component Reporting",
            *arguments,
            "--sync-resolve",
            "files",
            "--sync-component",
            "Reporting",
        )
        session.settle(quiet=8, expect="Imported Reporting")
        code = session.stop()
        r.steps.append({"exit": code})
    finally:
        if session.process.poll() is None:
            session.process.kill()
        close_excel(book)
    return r.steps


SCENES = {
    "check": scene_check,
    "export": scene_export,
    "edit": scene_edit,
    "git": scene_git,
    "headers": scene_headers,
    "safety": scene_safety,
    "filter": scene_filter,
    "references": scene_references,
    "config": scene_config,
    "bidirectional": scene_bidirectional,
}


def load_existing() -> dict:
    if not OUTPUT.exists():
        return {}
    text = OUTPUT.read_text(encoding="utf-8")
    return json.loads(text[text.index("{") : text.rindex("}") + 1])


def close_sample_workbooks(work: Path) -> None:
    """Close every workbook below the work directory that a scene left open; quit Excel if it is then empty."""
    import pythoncom
    import win32com.client

    try:
        application = win32com.client.GetActiveObject("Excel.Application")
        application.DisplayAlerts = False
        for book in list(application.Workbooks):
            if str(work).lower() in str(book.FullName).lower():
                book.Close(SaveChanges=False)
        if application.Workbooks.Count == 0:
            application.Quit()
    except pythoncom.com_error:
        pass
    time.sleep(3)


def write(casts: dict) -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(
        "/* Generated by website/tools/record_casts.py from real command runs. Do not edit by hand. */\n"
        "window.VBA_EDIT_CASTS = " + json.dumps(casts, indent=1, ensure_ascii=False) + ";\n",
        encoding="utf-8",
        newline="\n",
    )


def main() -> None:
    work = Path(sys.argv[1]).resolve()
    names = sys.argv[2:] or list(SCENES)
    work.mkdir(parents=True)
    plain = demo_project.build(work / "_template")
    forms = demo_project.build(work / "_template_forms", with_form=True)
    casts = load_existing()
    for name in names:
        print("recording", name, flush=True)
        try:
            casts[name] = SCENES[name](work, plain, forms)
        except Exception:
            # Show what was captured so far; a failed scene is never written.
            for step in RECORDERS[-1].steps if RECORDERS else []:
                print("   ", re.sub(r"<[^>]+>", "", json.dumps(step, ensure_ascii=False))[:600], flush=True)
            raise
        finally:
            close_sample_workbooks(work)
        write(casts)
    print("wrote", OUTPUT)


if __name__ == "__main__":
    main()

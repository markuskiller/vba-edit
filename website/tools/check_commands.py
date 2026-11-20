"""Check the commands shown on the site against the real command-line parsers.

python check_commands.py             (exit code 1 when something is wrong)
python check_commands.py --browser   (also check the rendered pages; read the warning below first)

check_site.py looks at links, anchors and recordings. This script looks at what the pages
claim about the tool itself:

1. Synopsis blocks. Every option in the synopsis of ``edit``, ``export`` and ``import`` exists
   in that command's parser, and every option of the parser is in the synopsis.
2. Config keys. Every key in the config tables of the guide is a key the tool reads.
3. Copyable commands. Every ``excel-vba ...`` line in a command block is parsed with the real
   parser. With ``--browser`` this is done on the rendered page for all four applications and
   all three ways of running the tool ("Your setup"), which also shows whether a block that
   should follow the choice still names Excel, and whether the .exe form starts with ``.\\``.

It needs vba-edit importable (``pip install -e .`` in the repository) and, for the rendered
checks, Microsoft Edge or Google Chrome.

WARNING about ``--browser``: on Windows, a headless Edge that starts with a new browser profile
registers one failed sign-in for the Windows account it runs under (observed with a local
account on Windows 11 on 2026-11-20; the reason is not known). Without a profile folder every
launch gets a new profile, and 36 launches in a row locked the account out of Remote Desktop
for ten minutes. This script therefore keeps ONE profile folder and reuses it for every launch
and every run, and it watches the account's failed sign-in counter: if the counter rises more
than once, it stops launching the browser and says so. The rendered checks are opt-in for the
same reason.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import re
import getpass
import shutil
import subprocess
import sys
import tempfile
from html.parser import HTMLParser
from pathlib import Path

from vba_edit import cli_common
from vba_edit.office_cli import OfficeVBACLI

PUBLIC = Path(__file__).resolve().parent.parent / "public"
PAGES = ["index.html", "quickstart.html", "guide.html"]

# Mirrors TOOLS in assets/site.js; test_site_js_agrees() below fails when the two drift apart.
APPS = {
    "excel": {"tool": "excel-vba", "doc": "Budget.xlsm", "noun": "workbook"},
    "word": {"tool": "word-vba", "doc": "Letter.docm", "noun": "document"},
    "powerpoint": {"tool": "powerpoint-vba", "doc": "Deck.pptm", "noun": "presentation"},
    "access": {"tool": "access-vba", "doc": "Orders.accdb", "noun": "database"},
}
RUNNERS = ["installed", "uvx", "exe"]
TOOL = re.compile(r"^(?:\.\\)?(excel|word|powerpoint|access)-vba(?:\.exe)?(?:@[\w.]+)?$")
VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}

# Options a synopsis may leave out, with the reason.
NOT_IN_SYNOPSIS = {
    "--help": "every command has it",
    "--config": "alias of --conf, named in the config chapter",
    "--no-colour": "alias of --no-color, named in the logs chapter",
    "--xlwings": "Excel only, described in the Excel chapter",
    "--pq-directory": "not implemented yet; hidden from the documentation on purpose",
}

problems: list[str] = []


# ── parsers ──────────────────────────────────────────────────────────────────────────────


def parser_for(app: str) -> argparse.ArgumentParser:
    return OfficeVBACLI(app).create_cli_parser()


def subparser(parser: argparse.ArgumentParser, command: str) -> argparse.ArgumentParser:
    for action in parser._actions:
        if isinstance(action, argparse._SubParsersAction):
            return action.choices[command]
    raise KeyError(command)


def parses(app: str, arguments: list[str]) -> str | None:
    """None if the real parser accepts the arguments, otherwise its error message."""
    errors = io.StringIO()
    try:
        with contextlib.redirect_stderr(errors), contextlib.redirect_stdout(io.StringIO()):
            parser_for(app).parse_args(arguments)
    except SystemExit as stop:
        if stop.code not in (0, None):
            lines = errors.getvalue().strip().splitlines()
            return lines[-1] if lines else f"exit code {stop.code}"
    return None


# ── reading pages ─────────────────────────────────────────────────────────────────────────


class Blocks(HTMLParser):
    """Collect the code blocks of a page, with what matters about where they sit."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack: list[tuple[str, dict]] = []
        self.blocks: list[dict] = []
        self.current: dict | None = None
        self.depth = 0

    @staticmethod
    def _classes(attributes: dict) -> set[str]:
        return set((attributes.get("class") or "").split())

    def _is_block(self, tag: str, attributes: dict) -> bool:
        classes = self._classes(attributes)
        return tag in ("pre", "code") or "card-cmd" in classes or "synopsis" in classes

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if tag in VOID:
            return
        if self.current is None and self._is_block(tag, attributes):
            fixed = "data-fixed" in attributes or any(
                "data-fixed" in a or self._classes(a) & {"term", "editor"} for _t, a in self.stack
            )
            hidden = "hidden" in attributes or any("hidden" in a for _t, a in self.stack)
            self.current = {
                "tag": tag,
                "classes": self._classes(attributes),
                "fixed": fixed or bool(self._classes(attributes) & {"term-body"}),
                "hidden": hidden,
                "text": "",
                "code": "",  # the text without prompts and comments
                "line": self.getpos()[0],
            }
            self.depth = len(self.stack)
        self.stack.append((tag, attributes))

    def handle_endtag(self, tag):
        if tag in VOID:
            return
        while self.stack:
            closed, _attributes = self.stack.pop()
            if closed == tag:
                break
        if self.current is not None and len(self.stack) <= self.depth:
            self.blocks.append(self.current)
            self.current = None

    def handle_data(self, data):
        if self.current is None:
            return
        self.current["text"] += data
        inside = [self._classes(a) for _t, a in self.stack[self.depth + 1 :]]
        if not any(classes & {"p", "c", "copy"} for classes in inside):
            self.current["code"] += data


def blocks_of(html: str) -> list[dict]:
    reader = Blocks()
    reader.feed(html)
    return reader.blocks


def split_arguments(text: str) -> list[str]:
    """Split a PowerShell command line the way the tool would receive it (double quotes only)."""
    arguments, current, quoted, started = [], "", False, False
    for character in text:
        if character == '"':
            quoted = not quoted
            started = True
        elif character.isspace() and not quoted:
            if started:
                arguments.append(current)
            current, started = "", False
        else:
            current += character
            started = True
    if started:
        arguments.append(current)
    return arguments


def command_lines(block: dict) -> list[list[str]]:
    """The vba-edit command lines of a block, each as a list of words starting with the runner."""
    lines = []
    for line in block["code"].splitlines():
        line = line.split(";")[0].strip()  # "command; if ($LASTEXITCODE ...) { ... }"
        if not line or "<" in line or "…" in line:
            continue  # a pattern such as "excel-vba <command> --help", not a command
        words = split_arguments(line)
        if any(TOOL.match(word) for word in words[:1]) or (words[0] == "uvx" and any(TOOL.match(w) for w in words)):
            lines.append(words)
    return lines


# ── checks on the source ─────────────────────────────────────────────────────────────────────


def check_synopsis(guide: str) -> None:
    excel = parser_for("excel")
    for block in blocks_of(guide):
        if "synopsis" not in block["classes"]:
            continue
        words = block["text"].split()
        command = words[1]
        documented = set(re.findall(r"(?<![\w-])--?[A-Za-z][\w-]*", block["text"]))
        actions = [a for a in subparser(excel, command)._actions if a.help is not argparse.SUPPRESS]
        real = {option for action in actions for option in action.option_strings}
        for option in sorted(documented - real):
            problems.append(f"guide.html: synopsis of '{command}' names {option}, which the command does not have")
        for action in actions:
            names = set(action.option_strings)
            if names and not names & documented and not names & set(NOT_IN_SYNOPSIS):
                problems.append(f"guide.html: synopsis of '{command}' leaves out {'/'.join(action.option_strings)}")


def check_config_keys(guide: str) -> None:
    known = {
        value for name, value in vars(cli_common).items() if name.startswith("CONFIG_KEY_") and isinstance(value, str)
    }
    section = guide[guide.index('<section id="config">') : guide.index('<h3 id="placeholders">')]
    for row in re.findall(r"<tr><td>(.*?)</td>", section):
        for key in re.findall(r"<code data-fixed>([a-z_]+)</code>", row):
            if key not in known:
                problems.append(f"guide.html: config key '{key}' is in the table but not read by the tool")


def check_site_js() -> None:
    script = (PUBLIC / "assets" / "site.js").read_text(encoding="utf-8")
    for app, values in APPS.items():
        expected = f'{app}: {{ tool: "{values["tool"]}", doc: "{values["doc"]}", noun: "{values["noun"]}"'
        if expected not in script:
            problems.append(f"site.js: TOOLS entry for {app} differs from APPS in check_commands.py")


def check_commands(name: str, html: str, app: str, run: str, rendered: bool) -> int:
    """Parse every command line of a page. ``rendered``: the page as the browser shows it for app/run."""
    checked = 0
    for block in blocks_of(html):
        if block["hidden"] or "term-body" in block["classes"]:
            continue
        where = f"{name} line {block['line']}" + (f" [{app}, {run}]" if rendered else "")
        swappable = not block["fixed"]
        if rendered and swappable and app != "excel":
            for leftover in ("excel-vba", "Budget", ".xlsm", "workbook"):
                if leftover in block["text"]:
                    problems.append(f"{where}: still says '{leftover}' after choosing {app}: {block['text'][:70]!r}")
        if "synopsis" in block["classes"]:
            continue  # a pattern with [optional] parts, compared with the parser in check_synopsis()
        for words in command_lines(block):
            position = next(index for index, word in enumerate(words) if TOOL.match(word))
            tool_word, arguments = words[position], words[position + 1 :]
            tool_app = TOOL.match(tool_word).group(1)
            if not arguments and block["tag"] == "code":
                continue  # the tool is only mentioned by name in running text
            if rendered and swappable:
                expected = {"installed": [APPS[app]["tool"]], "uvx": ["uvx", APPS[app]["tool"]]}.get(
                    run, [f".\\{APPS[app]['tool']}.exe"]
                )
                if words[: position + 1] != expected:
                    problems.append(f"{where}: expected the command to start with {' '.join(expected)!r}: {words[:3]}")
                    continue
            if error := parses(tool_app, arguments):
                problems.append(f"{where}: {tool_app}-vba does not accept {' '.join(arguments)!r}: {error}")
            checked += 1
    return checked


# ── the rendered page ──────────────────────────────────────────────────────────────────────────


def find_browser() -> str | None:
    candidates = [
        shutil.which("msedge"),
        shutil.which("chrome"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    return next((c for c in candidates if c and Path(c).exists()), None)


# One browser profile for every launch and every run: a new profile costs a failed sign-in (see the top).
BROWSER_PROFILE = Path(tempfile.gettempdir()) / "vba-edit-site-check-browser-profile"


def failed_signins() -> int | None:
    """The failed sign-in counter of the current Windows account, or None if it cannot be read."""
    try:
        import win32com.client

        return int(win32com.client.GetObject(f"WinNT://./{getpass.getuser()},user").BadPasswordAttempts)
    except Exception:
        return None


def render(browser: str, page: str, app: str, run: str) -> str:
    url = f"{(PUBLIC / page).as_uri()}?app={app}&run={run}"
    result = subprocess.run(
        [
            browser,
            "--headless=new",
            "--disable-gpu",
            f"--user-data-dir={BROWSER_PROFILE}",
            "--no-first-run",
            "--no-default-browser-check",
            "--virtual-time-budget=3000",
            "--dump-dom",
            url,
        ],
        capture_output=True,
        timeout=120,
    )
    return result.stdout.decode("utf-8", errors="replace")


def main() -> int:
    options = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    options.add_argument(
        "--browser",
        action="store_true",
        help="also check the rendered pages in headless Edge or Chrome (read the warning in this script first)",
    )
    options.add_argument("--max-launches", type=int, default=36, help=argparse.SUPPRESS)
    arguments = options.parse_args()

    sources = {page: (PUBLIC / page).read_text(encoding="utf-8") for page in PAGES}
    check_synopsis(sources["guide.html"])
    check_config_keys(sources["guide.html"])
    check_site_js()
    checked = sum(check_commands(page, html, "excel", "installed", rendered=False) for page, html in sources.items())
    print(f"{checked} command lines parsed as written")

    browser = find_browser() if arguments.browser else None
    if browser:
        combinations = 0
        signins_before = failed_signins()
        variants = [(page, app, run) for page in PAGES for app in APPS for run in RUNNERS]
        for launches, (page, app, run) in enumerate(variants[: arguments.max_launches], start=1):
            html = render(browser, page, app, run)
            signins_now = failed_signins()
            if signins_before is not None and signins_now is not None and signins_now - signins_before > 1:
                problems.append(
                    f"stopped after {launches} browser launches: the failed sign-in counter of this Windows "
                    f"account rose from {signins_before} to {signins_now}. More launches could lock the account; "
                    f"see the warning at the top of check_commands.py."
                )
                break
            if 'class="tk"' not in html and page != "index.html":
                problems.append(f"{page} [{app}, {run}]: the page was not rendered (no swapped commands found)")
                continue
            checked += check_commands(page, html, app, run, rendered=True)
            combinations += 1
        print(f"{combinations} of {len(variants)} rendered page variants checked (pages x applications x ways to run)")
        if signins_before is not None:
            print(f"failed sign-in counter of this account: {signins_before} before, {failed_signins()} after")
    elif arguments.browser:
        print("NOTE: rendered pages not checked (no Edge or Chrome found)")
    else:
        print("NOTE: rendered pages not checked (opt-in: --browser, after reading the warning in this script)")

    for problem in dict.fromkeys(problems):
        print("PROBLEM:", problem)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())

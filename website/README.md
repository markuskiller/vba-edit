# vba-edit docs site

A static site with no build step: three pages (`index.html`, `quickstart.html`, `guide.html`), one stylesheet,
one script, no external requests. Everything that gets served is in `public/`.

```
website/
├── public/                 what Cloudflare Pages serves
│   ├── index.html          landing page
│   ├── quickstart.html     five-step quickstart
│   ├── guide.html          the complete guide (one page, filterable table of contents)
│   ├── 404.html
│   ├── _headers            security headers, Content-Security-Policy, cache rules
│   ├── _redirects          /docs -> /guide and a few shortcuts
│   └── assets/
│       ├── site.css        design tokens, light and dark theme, all components
│       ├── site.js         theme, "your setup" switcher, terminal player, tabs, contents filter
│       ├── theme.js        applies a stored theme before first paint
│       ├── casts.js        GENERATED: the terminal recordings
│       └── img/            GENERATED: the screenshots
└── tools/                  maintainer scripts, not served
    ├── demo_project.py         builds the sample Budget.xlsm
    ├── record_casts.py         runs the real commands and writes assets/casts.js
    ├── cli_runner.py           helper for the recorder
    ├── capture_screenshots.py  takes the screenshots
    ├── check_site.py           checks links, anchors, recordings and the CSP rules
    └── check_commands.py       checks the commands and options on the pages against the real parsers
```

## Preview locally

```powershell
python -m http.server 8765 --directory website/public
# http://127.0.0.1:8765/
```

Links use the `.html` form so they work with any static server. Cloudflare Pages redirects them to the
extensionless URLs (`/guide`).

## Deploy to Cloudflare Pages

Connect the repository in the Cloudflare dashboard (Workers & Pages, Create, Pages, Connect to Git) with:

| Setting | Value |
|---|---|
| Framework preset | None |
| Build command | *(empty)* |
| Build output directory | `website/public` |
| Production branch | the branch the docs should track |

Or upload directly, without connecting the repository:

```powershell
npx wrangler pages deploy website/public --project-name vba-edit-docs
```

`_headers` sets a strict Content-Security-Policy (`script-src 'self'; style-src 'self'`). Inline `<script>`,
`<style>` and `style=""` attributes are therefore blocked; `check_site.py` fails if one slips in.

## Before publishing a change

```powershell
python website/tools/check_site.py
python website/tools/check_commands.py
```

`check_commands.py` needs vba-edit importable (`pip install -e .`). It compares the synopsis blocks and the
config key tables of the guide with the tool's own parsers, and parses every copyable command. With
`--browser` (and Edge or Chrome installed) it does that on the rendered pages for all four applications and
all three ways of running the tool (`?app=word&run=exe` selects a setup), and reports a block that still
names Excel after another application was chosen.

> **`--browser` is opt-in for a reason.** A headless Edge that starts with a new browser profile registers one
> failed sign-in for the Windows account it runs under. Launched 36 times without a profile folder, it locked
> the account out of Remote Desktop for ten minutes (2026-11-20). The script now reuses one profile folder
> (`%TEMP%\vba-edit-site-check-browser-profile`), which costs at most one failed sign-in, the first time the
> folder is created, and it stops launching the browser if the account's failed sign-in counter rises further.
> Do not script headless Edge or Chrome in a loop without a fixed `--user-data-dir`.

A config file or note that belongs to a command must follow the choice too: leave
`data-fixed` off it, and mark derived names with `data-doc-stem` or `data-app-noun`. Use `data-fixed` only
for blocks that are about one particular application.

What the pages say about behaviour (rollback, exit codes, encodings, recovery from a paused sync) is pinned
by the unit tests named after it: `tests/test_component_replacement.py`, `tests/test_failure_reporting.py`,
`tests/test_source_encoding.py`, `tests/test_check_command.py`, `tests/test_bidirectional_recovery.py`.

## Regenerating recordings and screenshots

Both scripts drive a real Excel on this machine, with "Trust access to the VBA project object model" enabled.
They work on a throwaway copy of the sample workbook in a new directory you name, and Excel windows open and
close on the desktop while they run. Use the project's development environment.

```powershell
cd website/tools
python record_casts.py C:\Temp\casts-run-1              # all scenes, about five minutes
python record_casts.py C:\Temp\casts-run-2 edit git     # only these scenes; the others are kept
python capture_screenshots.py C:\Temp\shots-run-1
```

- The recordings are the real output of the commands. The recorder changes two things only: the working
  directory is shown as `C:\Projects\budget`, and idle time is not recorded. The yellow "you" lines are
  narration added by the recorder.
- A page shows a recording with `<div class="term" data-cast="NAME">`. `data-steps="2-6"` shows a slice,
  and `data-running` marks a slice that ends while the command is still running.
- The screenshots capture one window each (the Config Generator, the VBA editor), not the screen. The menu
  language in the VBA editor screenshot is that of the Office installation used; the captions say so.

## When a release ships

Search `public/` for the version strings and adjust them: `0.6.0a4`, `0.5.1`, `November 2026`, and the
`badge new` marks (features that are in the pre-release but not in the stable release).

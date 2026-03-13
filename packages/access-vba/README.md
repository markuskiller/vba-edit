# access-vba belongs to [vba-edit](https://github.com/markuskiller/vba-edit) CLI Tools

**Edit VBA code in VS Code, PyCharm, Wing IDE, or any editor you love.** Real-time sync with **MS ACCESS**. Git-friendly. No more VBA editor pain.

[![CI](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg)](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml)
[![PyPI - Version](https://img.shields.io/pypi/v/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![Platform](https://img.shields.io/badge/platform-windows-blue.svg)](https://pypi.org/search/?q=vba-edit&o=&c=Operating+System+%3A%3A+Microsoft+%3A%3A+Windows)
[![vba-edit - Downloads](https://img.shields.io/pypi/dm/vba-edit)](https://www.pypiplus.com/project/vba-edit/)
[![access-vba - Downloads](https://img.shields.io/pypi/dm/access-vba)](https://www.pypiplus.com/project/access-vba/)
[![License](https://img.shields.io/badge/License-BSD_3--Clause-blue.svg)](https://opensource.org/licenses/BSD-3-Clause)

This is a thin entry-point package for [vba-edit](https://pypi.org/project/vba-edit/). Installing `access-vba` gives you the `access-vba` command and automatically installs the `vba-edit` core package.

## Quick Start

### RECOMMENDED: Via [``uvx``](https://docs.astral.sh/uv/guides/tools/)

```bash
uvx access-vba edit -f myfile.accdb --vba-directory .\src\VBA
```

> **Note:** The ``uvx`` command invokes a tool without installing it — same as: ``uv tool run access-vba``

### 30-Second Demo
```bash
# Start editing (uses active Access database)
uvx access-vba edit

# That's it! Edit the .bas/.cls files in your editor. Save = Sync.
```

## How It Works
<pre>
                        <--- vba-edit --->

    MS ACCESS                COMMANDS              Your favourite
                                v                       Editor

+------------------+                            +------------------+
|                  |                            |                  |
|   VBA Project    |   <---   EDIT*   (once ->) |  (e.g. VS CODE)  | 
|                  |                            |                  |     latest
|  (Office VBA-    |          EXPORT      --->  |   .bas           |  <- AI coding-  
|    Editor)       |                            |   .cls           |     assistants
|                  |   <---   IMPORT            |                  |   
|                  |                            |                  | 
|                  |                            |                  | 
+------------------+                            +------------------+
                                                         v
                                                +------------------+
                                                |                  |
 * watches & syncs                              |    (e.g. Git)    |
   back to Office                               |  version control |
   VBA-Editor live                              |                  |
   on save [CTRL+S]                             |                  |
                                                +------------------+
</pre>


## Install

```bash
pip install access-vba
```

> **Note:** Installing `access-vba` also installs the `vba-edit` core package, which provides entry points for all supported Office apps (`excel-vba`, `word-vba`, `powerpoint-vba`, `access-vba`).

## Usage

> All commands below work with both `access-vba <command>` (installed) and `uvx access-vba <command>` (no install required).

| Command | What it does |
|---------|-------------|
| `access-vba edit` | Start live editing |
| `access-vba export` | One-time export |
| `access-vba import` | One-time import |
| `access-vba export --open-folder --keep-open` | Export and open folder in explorer, keep document open for inspection |
| `access-vba export --force-overwrite` | Export without confirmation prompts |
| `access-vba check` | Verify status of *Trust access* to the VBA project object model |

> 💡 **Complete Option Matrix**: available **[here](https://langui.ch/current-projects/vba-edit/#OptionMatrix)**

> **Note:** Access VBA support covers standard modules and class modules only. UserForms are not supported.

For full documentation see [github.com/markuskiller/vba-edit](https://github.com/markuskiller/vba-edit).

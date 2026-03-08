# powerpoint-vba belongs to [vba-edit](https://github.com/markuskiller/vba-edit) CLI Tools

**Edit VBA code in VS Code, PyCharm, Wing IDE, or any editor you love.** Real-time sync with **MS POWERPOINT**. Git-friendly. No more VBA editor pain.

[![CI](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg)](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml)
[![PyPI - Version](https://img.shields.io/pypi/v/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![Platform](https://img.shields.io/badge/platform-windows-blue.svg)](https://pypi.org/search/?q=vba-edit&o=&c=Operating+System+%3A%3A+Microsoft+%3A%3A+Windows)
[![vba-edit - Downloads](https://img.shields.io/pypi/dm/vba-edit)](https://www.pypiplus.com/project/vba-edit/)
[![powerpoint-vba - Downloads](https://img.shields.io/pypi/dm/powerpoint-vba)](https://www.pypiplus.com/project/powerpoint-vba/)
[![License](https://img.shields.io/badge/License-BSD_3--Clause-blue.svg)](https://opensource.org/licenses/BSD-3-Clause)

This is a thin entry-point package for [vba-edit](https://pypi.org/project/vba-edit/). Installing `powerpoint-vba` gives you the `powerpoint-vba` command and automatically installs the `vba-edit` core package.

## Quick Start

### RECOMMENDED: Via [``uvx``](https://docs.astral.sh/uv/guides/tools/)

```bash
uvx powerpoint-vba edit -f myfile.pptm --vba-directory .\src\VBA
```

> **Note:** The ``uvx`` command invokes a tool without installing it — same as: ``uv tool run powerpoint-vba``

### 30-Second Demo
```bash
# Start editing (uses active PowerPoint presentation)
powerpoint-vba edit

# That's it! Edit the .bas/.cls files in your editor. Save = Sync.
```

## How It Works
<pre>
                        <--- vba-edit --->

  MS POWERPOINT              COMMANDS              Your favourite
                                v                       Editor

+------------------+                            +------------------+
|                  |                            |                  |
|   VBA Project    |   <---   EDIT*   (once ->) |  (e.g. VS CODE)  | 
|                  |                            |                  |     latest
|  (Office VBA-    |          EXPORT      --->  |   .bas           |  <- AI coding-  
|    Editor)       |                            |   .cls           |     assistants
|                  |   <---   IMPORT            |   .frm           |   
|                  |                            |  (.frx binary)   | 
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
pip install powerpoint-vba
```

> **Note:** Installing `powerpoint-vba` also installs the `vba-edit` core package, which provides entry points for all supported Office apps (`excel-vba`, `word-vba`, `powerpoint-vba`, `access-vba`).

## Usage

| Command | What it does |
|---------|-------------|
| `powerpoint-vba edit` | Start live editing |
| `powerpoint-vba export` | One-time export |
| `powerpoint-vba import` | One-time import |
| `powerpoint-vba export --open-folder --keep-open` | Export and open folder in explorer, keep document open for inspection |
| `powerpoint-vba export --force-overwrite` | Export without confirmation prompts |
| `powerpoint-vba check` | Verify status of *Trust access* to the VBA project object model |

> 💡 **Complete Option Matrix**: available **[here](https://langui.ch/current-projects/vba-edit/#OptionMatrix)**

For full documentation see [github.com/markuskiller/vba-edit](https://github.com/markuskiller/vba-edit).

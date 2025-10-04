# [vba-edit](https://github.com/markuskiller/vba-edit) 

**Edit VBA code in VS Code, Sublime, or any editor you love.** Real-time sync with MS Office apps (support for **Excel**, **Word**, **PowerPoint** & **Access**). Git-friendly. No more VBA editor pain.


[![CI](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg)](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml)
[![PyPI - Version](https://img.shields.io/pypi/v/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Downloads](https://img.shields.io/pypi/dm/vba-edit)](https://pypi.org/project/vba-edit)
[![License](https://img.shields.io/badge/License-BSD_3--Clause-blue.svg)](https://opensource.org/licenses/BSD-3-Clause)


## 30-Second Demo
```bash
# Install
pip install -U vba-edit

# Start editing (uses active Excel/Word document)
excel-vba edit    # or word-vba edit

# That's it! Edit the .bas/.cls files in your editor. Save = Sync.
```

## How It Works

```text

                         ←── vba-edit ──→

Excel / Word                 COMMANDS            Your favourite
PowerPoint / Access              ↓                   Editor

┌──────────────────┐                          ┌──────────────────┐
│                  │                          │                  │
│   VBA Project    │    ←──    EDIT* (once →) │  (e.g. VS CODE)  │ 
│                  │                          │                  │     latest
│  (Office VBA-    │           EXPORT ──→     │   .bas           │   ← AI coding-  
│    Editor)       │                          │   .cls           │     assistants
│                  │       ←── IMPORT         │   .frm           │   
│                  │                          │  (.frx binary)   │ 
│                  │                          │                  │ 
└──────────────────┘                          └──────────────────┘
                                                       ↓
                                              ┌──────────────────┐
                                              │                  │
 * watches & syncs                            │    (e.g. Git)    │
   back to Office                             │  version control │
   VBA-Editor live                            │                  │
   on save [CTRL+S]                           │                  │
                                              └──────────────────┘
```

## Why vba-edit?

- **Use YOUR editor** - VS Code, PyCharm, Wing IDE, Vim, etc. whatever you love 
- **AI-ready** - Use Copilot, ChatGPT, or any coding assistant 
- **Team-friendly** - Share code via Git, no COM add-ins needed 
- **Real version control** - Diff, merge, and track changes properly 
- **Well-organized** - Keep your VBA structured, clean, and consistent

## Setup (One-Time)

**Windows Only** | **MS Office**

Enable VBA access in Office:

`File → Options → Trust Center → Trust Center Settings → Macro Settings`

✅ **Trust access to the VBA project object model**

> 💡 Can't find it? Run `excel-vba check` to verify settings


## Common Workflows

### Start Fresh
```bash
excel-vba edit                    # Start with active workbook
```
### Team Project with Git

```bash
excel-vba export --vba-directory ./src/vba
git add . && git commit -m "Updated reports module"
``` 

### Support for RubberduckVBA Style (big thank you to @onderhold!)

```bash
excel-vba edit --rubberduck-folders --in-file-headers
``` 

## Quick Reference

| Command | What it does |
|---------|-------------|
| `excel-vba edit` | Start live editing |
| `excel-vba export` | One-time export |
| `excel-vba export --force-overwrite` | Export without confirmation prompts |
| `excel-vba check` | Verify status of *Trust access* to the VBA project object model |
| `--vba-directory ./src` | Custom folder |
| `--rubberduck-folders` | Organize by @Folder |
| `--in-file-headers` | Embed headers in code files |
| `--force-overwrite` | Skip safety prompts (automation) |

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Trust access" error | Run `excel-vba check` for diagnostics |
| Changes not syncing | Save the file in your editor |
| Forms not working | Add `--in-file-headers` flag |


## Safety Features

**🛡️ Data Loss Prevention** (v0.4.0+)

vba-edit now protects your work with smart safety checks:

- **Overwrite Protection**: Warns before overwriting existing VBA files
- **Header Mode Detection**: Alerts when switching between header storage modes
- **Orphaned File Cleanup**: Automatically removes stale `.header` files on mode change
- **UserForm Validation**: Prevents exports without proper header handling

**Bypass for Automation**: Use `--force-overwrite` flag to skip prompts in CI/CD pipelines:
```bash
excel-vba export --vba-directory ./src --force-overwrite
```

> ⚠️ **Note**: `--force-overwrite` suppresses all safety prompts. Use with caution!


## Features

**🚀 Core**
- Live sync between Office and your editor
- Full Git/version control support
- All Office apps: Excel, Word, Access & **NEW v0.4.0+** PowerPoint

**�️ Safety** (NEW v0.4.0+)
- Automatic detection of existing files before overwrite
- Header mode change warnings with cleanup
- UserForm header validation
- Optional `--force-overwrite` for automation

**�📁 Organization** 
- **NEW v0.4.0+** RubberduckVBA folder structure support
- **NEW v0.4.0+** Smart file organization with `@Folder` annotations
- **NEW v0.4.0+** TOML config files for team standards

**🔧 Advanced**
- Unicode & encoding support
- **IMPROVED v0.4.0+** UserForms with layout preservation  
- Class modules with custom attributes


## Command Line Tools

### App-specific tools

- `word-vba`
- `excel-vba`
- `access-vba`
- `powerpoint-vba`

### Commands

- `edit`: Live sync between editor and Office
- `export`: Export VBA modules to files
- `import`: Import VBA modules from files
- `check {all}`: Check if 'Trust Access to the VBA project object model' is enabled

### Options

```text
--file, -f                   Path to Office document
--conf, -c                   Supply config file
--vba-directory              Directory for VBA files
--rubberduck-folders         Use RubberduckVBA folder annotations
--save-headers               Save module headers separately
--in-file-headers            Include VBA headers directly in code files
--encoding, -e               Specify character encoding
--detect-encoding, -d        Auto-detect encoding
--verbose, -v                Enable detailed logging
--logfile, -l                Enable file logging
--version, -V                Show program's version number and exit
```

### Example of `--in-file-headers --rubberduck-folders` (v0.4.0+)

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Business.Domain")
Public Sub DoSomething()
    ' Your code here
End Sub
```

> [!CAUTION]
> **1.** Always **backup your Office files** before using `vba-edit` **2.** Use **version control (git)** to track your VBA code **3.** Run `export` after changing **form layouts** or module properties


### Known Limitations

- UserForms require `--save-headers` option (`edit` process is aborted if this is not the case)
- If `*.header` files are modified on their own, the corresponding `*.cls`, `*.bas` or `*.frm` file needs to be saved in order to sync the complete module back into the VBA project model

## Links

- [Homepage](https://langui.ch/current-projects/vba-edit/)
- [Documentation](https://github.com/markuskiller/vba-edit/blob/main/README.md)
- [Source Code](https://github.com/markuskiller/vba-edit)
- [Changelog](https://github.com/markuskiller/vba-edit/blob/main/CHANGELOG.md)
- [Changelog of latest dev version](https://github.com/markuskiller/vba-edit/blob/dev/CHANGELOG.md)
- [Video Tutorial](https://www.youtube.com/watch?v=xoO-Fx0fTpM) (xlwings walkthrough, with similar functionality)

## License

BSD 3-Clause License

## Credits

**vba-edit** builds on an excellent idea first implemented for Excel in [xlwings](https://www.xlwings.org/) (BSD-3).

Special thanks to **@onderhold** for improved header handling, RubberduckVBA folder and config file support in v0.4.0.
# [vba-edit](https://github.com/markuskiller/vba-edit) 

**Edit VBA code in VS Code, PyCharm, Wing IDE, or any editor you love.** Real-time sync with MS Office apps (support for **Excel**, **Word**, **PowerPoint** & **Access**). Git-friendly. No more VBA editor pain.

[![CI](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg)](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml)
[![PyPI - Version](https://img.shields.io/pypi/v/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![Platform](https://img.shields.io/badge/platform-windows-blue.svg)](https://github.com/markuskiller/vba-edit)
[![PyPI - Downloads](https://img.shields.io/pypi/dm/vba-edit)](https://pypi.org/project/vba-edit)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
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
<pre>
                        <--- vba-edit --->

Excel / Word                 COMMANDS              Your favourite
PowerPoint / Access             v                       Editor

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

## Why vba-edit?

- **Use YOUR editor** - VS Code, PyCharm, Wing IDE, Sublime, Vim, etc. whatever you love 
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

### Quick Export with Folder View
```bash
excel-vba export --open-folder    # Export and open in File Explorer
excel-vba export --keep-open      # Export but keep document open for inspection
excel-vba export --no-color       # Export without colorized output
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
| `excel-vba export --open-folder` | Export and open folder in explorer |
| `excel-vba export --force-overwrite` | Export without confirmation prompts |
| `excel-vba check` | Verify status of *Trust access* to the VBA project object model |
| `--vba-directory ./src` | Custom folder |
| `--rubberduck-folders` | Organize by @Folder |
| `--in-file-headers` | Embed headers in code files |
| `--conf myconfig.toml` | Use config file |
| `--force-overwrite` | Skip safety prompts (automation) |
| `--keep-open` | Keep document open after export |
| `--no-color` | Disable colorized terminal output |

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

> ⚠️ **CAUTION**: `--force-overwrite` suppresses all safety prompts. Use with caution!


## Features

**🚀 Core**
- Live sync between Office and your editor
- Full Git/version control support
- All Office apps: Excel, Word, Access & **NEW v0.4.0+** PowerPoint

**📁 Organization** 
- **NEW v0.4.0+** RubberduckVBA folder structure support
- **NEW v0.4.0+** Smart file organization with `@Folder` annotations
- **NEW v0.4.0+** TOML config files for team standards

**🔧 Advanced**
- Unicode & encoding support
- **IMPROVED v0.4.0+** UserForms with layout preservation  
- Class modules with custom attributes


## Roadmap

### 🎯 Current Focus (v0.4.x)
- 🔄 Stable VBA editing for Excel, Word, PowerPoint, Access
- 🔄 Enhanced CLI with organized help and grouped options
- 🔄 Refining RubberduckVBA integration (`@Folder` annotations)
- 🔄 Optimizing configuration file support (TOML)
- 🔄 Safety features and data loss prevention
- 🔄 Colorized output for better readability (coming in v0.4.1)
- 🔄 Windows binaries with security verification (SHA256, SBOM, Attestations)

### 🚀 Next release (v0.5.0)

**Manual Bi-Directional Sync**
- Export Office VBA changes back to files during `edit` mode
- Keyboard shortcut in terminal (e.g., press `E` to export)
- Safety warnings for unsaved files
- Automatic backups before overwrite
- _No automatic polling - explicit user control_

**PowerQuery Support** (Excel)
- Export/Import PowerQuery queries (M language)
- Similar workflow to VBA editing
- `.pq` or `.m` file format support

**CLI Enhancements**
- Progress indicators
- Improved error messages

### 🔮 Future (v0.6.0+)

**Advanced Features**
- Multi-document support (watch multiple Office files)
- Dependency tracking and visualization
- Git workflow automation

**Code Signing** (if feasible)
- Signed Windows executables via SignPath.io
- No more SmartScreen warnings
- Enhanced trust and security

### 💡 Feedback & Contributions

Found a bug? Have a feature idea? Questions about usage? Open an [Issue](https://github.com/markuskiller/vba-edit/issues) - we use labels to organize different types of feedback.

---

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
--conf, --config, -c         Supply config file (TOML format)
--vba-directory              Directory for VBA files
--rubberduck-folders         Use RubberduckVBA folder annotations
--save-headers               Save module headers separately
--in-file-headers            Include VBA headers directly in code files
--encoding, -e               Specify character encoding
--detect-encoding, -d        Auto-detect encoding
--verbose, -v                Enable detailed logging
--logfile, -l                Enable file logging
--open-folder                Open export directory in file explorer after export
--save-metadata, -m          Save metadata file with encoding information
--force-overwrite            Skip all confirmation prompts (for automation)
--keep-open                  Keep document open after export (default: closes automatically)
--no-color                   Disable colorized terminal output
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

## Colorized Output

**NEW v0.4.1** Terminal output features professional color-coded messages for better readability:

- ✓ **Success** messages in green
- ✗ **Error** messages in red
- ⚠ **Warning** messages in yellow
- **Technical terms** (VBA, TOML, JSON) highlighted in cyan
- **Code examples** shown in dim gray

**Automatic Behavior:**
- Colors automatically disabled when output is piped or redirected
- Disabled in non-TTY environments (CI/CD pipelines)
- Respects `NO_COLOR` environment variable

**Manual Control:**
```bash
excel-vba export --no-color              # Disable colors
export NO_COLOR=1; excel-vba export      # Via environment variable
```

> 💡 **Tip**: Use `--no-color` when saving output to log files or when terminal colors cause issues.

## Configuration Files

**NEW v0.4.0+** Use TOML configuration files to standardize team workflows and avoid repetitive command-line arguments.

### Basic Configuration

Create a `vba-config.toml` file in your project:

```toml
[general]
file = "MyWorkbook.xlsm"
vba_directory = "src/vba"
verbose = true
rubberduck_folders = true
in_file_headers = true
```

Then use it:
```bash
excel-vba export --conf vba-config.toml
```

### Available Configuration Keys

**[general] section:**
- `file` - Path to Office document
- `vba_directory` - Directory for VBA files
- `encoding` - Character encoding (e.g., "utf-8", "cp1252")
- `verbose` - Enable verbose logging (true/false)
- `logfile` - Path to log file
- `rubberduck_folders` - Use RubberduckVBA @Folder annotations (true/false)
- `save_headers` - Save headers to separate .header files (true/false)
- `in_file_headers` - Embed headers in code files (true/false)
- `open_folder` - Open export directory after export (true/false)
- `keep_open` - Keep document open after export (true/false)
- `no_color` - Disable colorized terminal output (true/false)

**Other sections (reserved for future use):**
- `[office]` - Office-wide settings
- `[excel]` - Excel-specific settings
- `[word]` - Word-specific settings
- `[access]` - Access-specific settings
- `[powerpoint]` - PowerPoint-specific settings

### Configuration Placeholders

Configuration values support dynamic placeholders for flexible path management.

**Available placeholders (v0.4.1+):**
- `{config.path}` - Directory containing the config file
- `{file.name}` - Document filename without extension
- `{file.fullname}` - Document filename with extension
- `{file.path}` - Directory containing the document
- `{file.vbaproject}` - VBA project name (resolved at runtime)

**Legacy placeholders (deprecated in v0.4.1, removed in v0.5.0):**
- `{general.file.name}` → use `{file.name}`
- `{general.file.fullname}` → use `{file.fullname}`
- `{general.file.path}` → use `{file.path}`
- `{vbaproject}` → use `{file.vbaproject}`

**Example with placeholders:**

```toml
[general]
file = "C:/Projects/MyApp/MyWorkbook.xlsm"
vba_directory = "{file.path}/{file.name}-vba"
# This resolves to: C:/Projects/MyApp/MyWorkbook-vba
```

**Relative paths example:**

```toml
[general]
file = "../documents/report.xlsm"
vba_directory = "{config.path}/vba-modules"
# vba_directory is relative to config file location
```

### Command-Line Override

Command-line arguments always override config file settings:

```bash
# Config says vba_directory = "src/vba"
# This overrides it to "build/vba"
excel-vba export --conf vba-config.toml --vba-directory build/vba
```

> ⚠️ **CAUTION**: **1.** Always **backup your Office files** before using `vba-edit` **2.** Use **version control (git)** to track your VBA code **3.** Run `export` after changing **form layouts** or module properties


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
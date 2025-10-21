# [vba-edit](https://github.com/markuskiller/vba-edit) 

**Edit VBA code in VS Code, PyCharm, Wing IDE, or any editor you love.** Real-time sync with MS Office apps (support for **Excel**, **Word**, **PowerPoint** & **Access**). Git-friendly. No more VBA editor pain.

[![CI](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg)](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml)
[![PyPI - Version](https://img.shields.io/pypi/v/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![Platform](https://img.shields.io/badge/platform-windows-blue.svg)](https://pypi.org/search/?q=vba-edit&o=&c=Operating+System+%3A%3A+Microsoft+%3A%3A+Windows)
[![PyPI - Downloads](https://img.shields.io/pypi/dm/vba-edit)](https://www.pypiplus.com/project/vba-edit/)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![License](https://img.shields.io/badge/License-BSD_3--Clause-blue.svg)](https://opensource.org/licenses/BSD-3-Clause)

## Installation

### Python Package (Recommended)
```bash
pip install -U vba-edit
```

### Windows Binaries (No Python Required)

**NEW v0.4.1+** Standalone executables available - no Python installation needed!

📦 **Download**: [Latest Release](https://github.com/markuskiller/vba-edit/releases/latest) (scroll all the way down to **Assets** section)

Available binaries:
- `excel-vba.exe` - For Excel workbooks
- `word-vba.exe` - For Word documents
- `access-vba.exe` - For Access databases
- `powerpoint-vba.exe` - For PowerPoint presentations

🔒 **Security**: See [Security Verification Guide](SECURITY_VERIFICATION.md) for SHA256 checksums and attestation validation.

> ⚠️ **Windows SmartScreen & VirusTotal Warnings**: Binaries are currently unsigned. You may see warning messages - this is expected. See [Issue #24](https://github.com/markuskiller/vba-edit/issues/24) and [Security Info](SECURITY.md) for details.

## 30-Second Demo
```bash
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

### App-specific tools

| CLI Tool | Description |
|---------|-------------|
| `excel-vba.exe` | For Excel Workbooks |
| `word-vba.exe` | For Word documents |
| `access-vba.exe` | For Access databases |
| `powerpoint-vba.exe` | For PowerPoint presentations |

### Commands

| Command overview | Description |
|---------|-------------|
| `edit` | Edit VBA content in Office document |
| `import` | Import VBA content into Office document |
| `export` | Export VBA content from Office document |
| `references` | Manage VBA library references (list, export, import) |
| `check` | Check if 'Trust Access to the Office VBA project object model' is enabled |

> 💡 Use **`*-vba.exe <command> --help`** in your terminal for a detailed option overview.

### Examples of options
| Command | What it does |
|---------|-------------|
| `excel-vba edit` | Start live editing |
| `excel-vba export` | One-time export |
| `excel-vba import` | One-time import |
| `word-vba references list -f template.dotm` | List all VBA references in document |
| `word-vba references export -f doc.docx -r refs.toml` | Export references to TOML config |
| `word-vba references import -f doc.docx -r refs.toml` | Import references from TOML config |
| `excel-vba export --open-folder --keep-open` | Export and open folder in explorer, keep document open for inspection |
| `excel-vba export --force-overwrite` | Export without confirmation prompts |
| `excel-vba check` | Verify status of *Trust access* to the VBA project object model |

> 💡 **Complete Option Matrix**: available **[here](https://langui.ch/current-projects/vba-edit/#OptionMatrix)**

## VBA Reference Management

**NEW in v0.5.0+** Manage VBA library references across your Office documents with version control support.

### Why Manage References?

- **Team consistency**: Share library configurations via Git
- **Document templates**: Export reference configs from master templates
- **Dependency tracking**: Know which libraries your VBA code needs
- **Path updates**: Update references when shared libraries move or are renamed

### Reference Commands

```bash
# List all references in a document
word-vba references list -f template.dotm

# Export references to TOML config
word-vba references export -f template.dotm -r template-refs.toml

# Import references into a new document
word-vba references import -f new-doc.docx -r template-refs.toml
```

### What Gets Exported?

Reference exports create TOML files with:
- Library name, GUID, and version (for COM references)
- File paths (critical for document-based references like `.dotm` templates)
- Built-in references (VBA, Office) are excluded automatically

**Example TOML:**
```toml
[[references]]
name = "ADODB"
guid = "{B691E011-1797-432E-907A-4D8C69339129}"
major = 6
minor = 1
description = "Microsoft ActiveX Data Objects 6.1 Library"
path = "C:\\Program Files\\Common Files\\System\\ado\\msado15.dll"

[[references]]
name = "SharedTemplates"
guid = ""
path = "\\\\server\\shared\\templates\\SharedCode.dotm"
```

### Supported Applications

Reference management works with:
- ✅ **Word** - Full support (modules, classes, forms, document references)
- ✅ **Excel** - Full support (modules, classes, forms, workbook references)
- ✅ **PowerPoint** - Full support (modules, classes, forms)
- ✅ **Access** - Modules and classes only (Access forms work differently)

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Trust access" error | Run `excel-vba check` for diagnostics |
| Changes not syncing | Save the file in your editor |
| Forms not working | Add `--in-file-headers` flag |
| Reference import fails | Check that library files exist at specified paths |


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

**� Reference Management** *(NEW v0.5.0+)*
- List, export, and import VBA library references
- Track COM libraries (ADODB, MSForms, etc.)
- Manage document-based references (.dotm, .xlam templates)
- Path tracking for shared template libraries
- Version control friendly TOML format

**�🔧 Advanced**
- Unicode & encoding support
- **IMPROVED v0.4.0+** UserForms with layout preservation  
- Class modules with custom attributes


## Roadmap

Development priorities evolve based on user feedback and real-world needs. 

+👀 **See active planning**: [GitHub Milestones](https://github.com/markuskiller/vba-edit/milestones)  
+💡 **Request features**: [Open an Issue](https://github.com/markuskiller/vba-edit/issues)  
+📝 **Current focus**: v0.5.0 - Batch operations for processing multiple Office documents


### 💡 Feedback & Contributions

Found a bug? Have a feature idea? Questions about usage? Open an [Issue](https://github.com/markuskiller/vba-edit/issues) - we use labels to organize different types of feedback.

---

## Command Line Tools

> 💡 **Complete CLI Overview**: available **[here](https://langui.ch/current-projects/vba-edit/#CLI)**

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

> 💡 **Tip**: Use `--no-color` when terminal colors cause issues.

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

- UserForms require `--save-headers` or newer `--in-file-headers` option (`edit` process is aborted if this is not the case)
- If separate `*.header` files are modified on their own, the corresponding `*.cls`, `*.bas` or `*.frm` file needs to be saved in order to sync the complete module back into the VBA project model

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
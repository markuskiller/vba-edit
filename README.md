# vba-editRD
This is a fork of [vba-edit](https://github.com/markuskiller/vba-edit), extended to optionally use RubberduckVBA folder annotations when importing/exporting from/to folder hierarchies.

[vba-edit](https://github.com/markuskiller/vba-edit) enables seamless Microsoft Office VBA code editing in your preferred editor or IDE, facilitating the use of coding assistants and version control workflows.

## Features

- Edit VBA code in your favorite code editor or IDE
- Automatically sync changes between your editor and Office applications
- Support for Word, Excel, Access, and PowerPoint
- Preserve form layouts and module properties
- Handle different character encodings
- Integration with version control systems
- Support for UserForms and class modules
- Optionally uses RubberduckVBA folder annotations when importing/exporting from/to folder hierarchies

## Quick Start

### Prerequisites

Enable "Trust access to the VBA project object model" in your Office application's Trust Center settings:

1. Open your Office application
2. Go to File > Options > Trust Center > Trust Center Settings
3. Select "Macro Settings"
4. Check "Trust access to the VBA project object model"

### Basic Usage

#### Excel Example

1. Open your Excel workbook with VBA code
2. In your terminal, run:

    ```bash
    excel-vba edit
    ```

3. Edit the exported .bas, .cls, or .frm files in your preferred editor
4. Changes are automatically synced back to Excel when you save

#### Word Example

```bash
# Export VBA modules from active document
word-vba export --vba-directory ./VBA

# Edit and sync changes automatically
word-vba edit --vba-directory ./VBA

# Import changes back to document
word-vba import --vba-directory ./VBA
```

## Detailed Features

### Supported File Types

- Standard Modules (.bas)
- Class Modules (.cls)
- UserForms (.frm)
- Document/Workbook Modules

### Command Line Tools

The package provides separate command-line tools for each Office application:

- `word-vba`
- `excel-vba`
- `access-vba`
- `powerpoint-vba`

Each tool supports three main commands (plus `check {all}` for troubleshooting):

- `edit`: Live sync between editor and Office (Word/Excel only)
- `export`: Export VBA modules to files
- `import`: Import VBA modules from files
- `check`: Check if 'Trust Access to the VBA project object model' is enabled

> [!NOTE]
> The command `check all` can be used to troubleshoot Trust Access to VBA project object model,
> scanning and giving feedback on **all supported MS Office apps**

### Common Options

```text
--file, -f                   Path to Office document (optional)
--vba-directory              Directory for VBA files
--encoding, -e               Specify character encoding
--detect-encoding, -d        Auto-detect encoding
--save-headers               Save module headers separately
--verbose, -v                Enable detailed logging
--logfile, -l                Enable file logging
--rubberduck-folders         Use RubberduckVBA folder annotations
```

### Excel-Specific Features

For Excel users who also have xlwings installed:

```bash
excel-vba edit -x  # Use xlwings wrapper
```

## New Features (v0.4.0+)

### In-File Headers (Default: Enabled)
VBA headers are now embedded directly in code files by default:

**MyClass.cls:**
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

## Best Practices

1. **New Projects and Workflows**: Use default settings (in-file headers + Rubberduck folders)
2. **Workflows with version < v0.4.0 **: Add `--save-headers --no-in-file-headers` for compatibility
3. Always backup your Office files before using vba-edit
4. Use version control (git) to track your VBA code
5. Run `export` after changing form layouts or module properties
6. Consider using `--detect-encoding` for non-English VBA code

## Known Limitations

- UserForms require `--save-headers` option (`edit` process is aborted if this is not the case)
- If `*.header` files are modified on their own, the corresponding `*.cls`, `*.bas` or `*.frm` file needs to be saved in order to sync the complete module back into the VBA project model

## Development Setup

see ./docs/development/DEVELOPMENT.md

## License

BSD 3-Clause License
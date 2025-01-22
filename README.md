# vba-edit

Enable seamless Microsoft Office VBA code editing in your preferred editor or IDE, facilitating the use of coding assistants and version control workflows.

[![CI](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg)](https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml)
[![PyPI - Version](https://img.shields.io/pypi/v/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Downloads](https://img.shields.io/pypi/dm/vba-edit)](https://pypi.org/project/vba-edit)
[![License](https://img.shields.io/badge/License-BSD_3--Clause-blue.svg)](https://opensource.org/licenses/BSD-3-Clause)

## Features

- Edit VBA code in your favorite code editor or IDE
- Automatically sync changes between your editor and Office applications
- Support for Word, Excel, Access, and PowerPoint
- Preserve form layouts and module properties
- Handle different character encodings
- Integration with version control systems
- Support for UserForms and class modules

> [!NOTE]
> Inspired by code from ``xlwings vba edit`` ([xlwings-Project](https://www.xlwings.org/)) under the BSD 3-Clause License.

## Quick Start

### Installation

```bash
pip install vba-edit
```

### Prerequisites

Enable "Trust access to the VBA project object model" in your Office application's Trust Center Settings:

1. Open your Office application
2. Go to File > Options > Trust Center > Trust Center Settings
3. Select "Macro Settings"
4. Check "Trust access to the VBA project object model"

> [!NOTE]
> In MS Access, Trust Access to VBA project object model is always enabled if database is stored in trusted location.

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

#### Access Example

```bash
# Export VBA modules
access-vba export --vba-directory ./VBA

# Import changes
access-vba import --vba-directory ./VBA
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
--file, -f             Path to Office document (optional)
--vba-directory        Directory for VBA files
--encoding, -e         Specify character encoding
--detect-encoding, -d  Auto-detect encoding
--save-headers         Save module headers separately
--verbose, -v          Enable detailed logging
--logfile, -l          Enable file logging
--version              Show program's version number and exit
```

### Excel-Specific Features

For Excel users who also have xlwings installed:

```bash
excel-vba edit -x  # Use xlwings wrapper
```

## Best Practices

1. Always backup your Office files before using vba-edit
2. Use version control (git) to track your VBA code
3. Run `export` after changing form layouts or module properties
4. Use `--save-headers` when working with UserForms
5. Consider using `--detect-encoding` for non-English VBA code

## Known Limitations

- UserForms require `--save-headers` option (`edit` process is aborted if this is not the case)
- If `*.header` files are modified on their own, the corresponding `*.cls`, `*.bas` or `*.frm` file needs to be saved in order to sync the complete module back into the VBA project model
- `--in-file-headers` option coming soon (thanks to @cargocultprogramming for suggestion and use case context)

## Links

- [Homepage](https://langui.ch/current-projects/vba-edit/)
- [Documentation](https://github.com/markuskiller/vba-edit/blob/main/README.md)
- [Source Code](https://github.com/markuskiller/vba-edit)
- [Changelog](https://github.com/markuskiller/vba-edit/blob/main/CHANGELOG.md)
- [Changelog of latest dev version](https://github.com/markuskiller/vba-edit/blob/dev/CHANGELOG.md)
- [Video Tutorial](https://www.youtube.com/watch?v=xoO-Fx0fTpM) (xlwings walkthrough, with similar functionality)

## License

BSD 3-Clause License

## Acknowledgments

This project is heavily inspired by code from `xlwings vba edit`, maintained by the [xlwings-Project](https://www.xlwings.org/) under the BSD 3-Clause License.

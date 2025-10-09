# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.1a1] - unreleased

### Added

- **Enhanced CLI Help System**: Complete refactoring of all CLI entry points with improved help formatting
  - Streamlined main help with concise descriptions and 4 key examples per entry point
  - Grouped argument options (File Options, Configuration, Encoding Options, Header Options, Command-Specific Options, Common Options) (Specs provided by @oderhold)
  - Consistent formatting using `EnhancedHelpFormatter` (title case headings, single colons)
  - Professional "Commands" section with clear `<command>` metavar
  - Enhanced "check" command with proper "Subcommands" section
- **--save-metadata Option**: Added `-m/--save-metadata` flag to both `edit` and `export` commands
  - Allows saving metadata during initial export in edit mode
  - Previously only available in export command
- `add_folder_organization_arguments()` function in `cli_common.py` for command-specific folder options
- **Simplified Placeholders**: New streamlined placeholder format for configuration files
  - `{file.name}` - replaces `{general.file.name}` 
  - `{file.fullname}` - replaces `{general.file.fullname}`
  - `{file.path}` - replaces `{general.file.path}`
  - `{file.vbaproject}` - replaces `{vbaproject}`
  - Legacy placeholders still supported for backward compatibility (will be removed in v0.5.0)
- **Enhanced Help Formatter**: New help output system for better CLI documentation
  - `EnhancedHelpFormatter` class for improved help text formatting
  - `GroupedHelpFormatter` class for organizing options into logical groups
  - Helper functions for creating consistent help output across all entry points
  - Predefined example templates for all commands (edit, import, export, check)

### Changed

- **CLI Architecture**: All four entry points (excel-vba, word-vba, access-vba, powerpoint-vba) now use inline argument definitions
  - Removed dependency on old helper functions (`add_common_arguments`, `add_encoding_arguments`, etc.)
  - Improved code visibility and maintainability with explicit argument declarations
  - All entry points now have identical structure and formatting
  - Main help shows overview with basic examples
  - Command-specific help shows detailed usage with all options
  - Better organization with mutually exclusive groups properly labeled
  - Folder organization options (`--rubberduck-folders`, `--open-folder`) are now command-specific (edit/import/export only)
  - Removed `add_common_arguments()` from main parser level to prevent global option leakage
- **Option Availability**: `--open-folder` now available on all manipulation commands (edit, import, export) for improved workflow continuity
- Updated all handler initializations to use `getattr()` with defaults for optional folder organization arguments
- **Placeholder System**: Simplified configuration placeholder naming for better clarity and consistency

### Deprecated

- Legacy placeholder format `{general.file.*}` and `{vbaproject}` (use new `{file.*}` format instead)

### Fixed

- **[Issue #20](https://github.com/markuskiller/vba-edit/issues/20)**: Fixed AttributeError in `check` command for all entry points (word-vba, excel-vba, access-vba, powerpoint-vba)
  - Modified `validate_paths()` to skip validation for `check` command
  - Added `hasattr()` guards for safer attribute access
- **[Issue #21](https://github.com/markuskiller/vba-edit/issues/21)**: CLI options are now properly scoped as command-specific vs global
  - `--rubberduck-folders` and `--open-folder` only appear on commands where they're applicable
  - `--help` and `--version` remain as global options
  - Prevents user confusion from seeing irrelevant options on incompatible commands

## [0.4.0] - 2025-10-06

### Added
- new option `--in-file-headers`: embedding VBA headers in code files ([@onderhold](https://github.com/onderhold))
- new option `--rubberduck-folders`: support for RubberduckVBA-style `@Folder` annotations and folder organization ([Issue #9](https://github.com/markuskiller/vba-edit/issues/9)) ([@onderhold](https://github.com/onderhold))
- new option `--conf`: support for config files  ([@onderhold](https://github.com/onderhold))
- new option `--force-overwrite`: skip safety prompts for automated exports (CI/CD friendly)
- Automatic warnings before overwriting existing VBA files
- Detection and warning when header storage mode changes between exports
- Automatic cleanup of orphaned `.header` files when switching modes
- Enhanced UserForm validation - prevents export without proper header handling (`--in-file-headers` or `--save-headers`)
- Better support for VBA class modules with custom attributes ([@onderhold](https://github.com/onderhold))
- Enhanced `VB_PredeclaredId` handling for class modules ([@onderhold](https://github.com/onderhold))
- Support for macro-enabled MS PowerPoint documents
- `check all` subcommand for cli entry points, which processes all suported MS Office apps in a single call (replaces calling `python -m vba_edit.utils`)
- Option to show program's version number and exit added to all cli interfaces (`--version`)
- Metadata tracking: exports now save `header_mode` to detect configuration changes between runs
- **Build**: Added version file generation support to `create_binaries.py` for Windows executables

### Changed
- Improved version control compatibility with embedded headers ([@onderhold](https://github.com/onderhold))
- Streamlined project setup by extending pyproject.toml and .gitignore, while reducing requirements.txt to the bare minimum that VS Code needs. Thus setup.cfg became superfluous. ([@onderhold](https://github.com/onderhold))
- Some refactoring: handling common cli options now in separate module cli_common.py ([@onderhold](https://github.com/onderhold))
- Refactored warning handling logic into centralized helper function (`handle_export_with_warnings()` in `cli_common.py`)
- Core logic (`office_vba.py`) now raises `VBAExportWarning` exceptions instead of handling user interaction
- CLI layer handles all user prompts via shared helper, eliminating code duplication across entry points
- **[Issue #14](https://github.com/markuskiller/vba-edit/issues/14)**: The watchgod library is no longer actively developed and has been superseded by watchfiles. The dependency has been replaced with watchfiles ([@onderhold](https://github.com/onderhold))

### Fixed
- **[Issue #16](https://github.com/markuskiller/vba-edit/issues/16)**: Hidden member attributes (VB_VarHelpID, VB_VarDescription, VB_UserMemId) no longer appear in VBA editor after import. These attributes are legal in exported VBA files but cause syntax errors when written directly into modules. Solution: filter all member-level Attribute lines from code sections before calling AddFromString(). (Reported by [@takutta](https://github.com/takutta), [@loehnertj](https://github.com/loehnertj))
- **[Issue #11](https://github.com/markuskiller/vba-edit/issues/11)**: Verified `--in-file-headers` feature works correctly for LSP tool compatibility (VBA Pro extension support)
- **UserForm parsing**: Fixed case-sensitive BEGIN block matching - now handles both "BEGIN" (class modules) and "Begin {GUID} FormName" (UserForms) correctly
- fix check for form safety on `export` (if edit command is run without `--save-headers` option, forms cannot be processed correctly -> check for forms and abort if `--save-headers` is not enabled)
- fix header file handling (`--save-headers`) in already populated `--vba-directory` (only 1 header file was created rather than one per *.cls, *.bas or *.frm file) - calling it on empty `--vba-directory` worked as expected
- `--in-file-headers` validation now correctly allows UserForm exports (was incorrectly rejected even when flag was set)
- Boolean logic error in `_check_form_safety()` that caused false rejections

### Data loss prevention
- Export operations now require explicit user confirmation when:
  - Overwriting existing VBA files (prevents accidental data loss)
  - Changing header storage modes (prevents configuration drift and orphaned files)
- Confirmation prompts can be bypassed with `--force-overwrite` flag for automation or for backward compatibility scenarios with `vba-edit<0.4.0` or `xlwings vba edit`

## [0.3.0] - 2025-01-19

### Added

- new command ``check``: e.g. ``word-vba check`` detects if 'Trust Access to Word VBA project object model' is enabled (available for all entry points)
- ACCESS: basic support for MS Access databases (standard modules ``*.bas`` and class modules ``*.cls`` are supported)
- ``--save-headers`` command-line option, supporting a comprehensive handling of headers ([Issue #11](https://github.com/markuskiller/vba-edit/issues/11))
- new files created in ``--vba-directory`` are now automatically synced back to MS Office VBA editor in ``*-vba edit`` mode (previously, only file deletions were monitored and synced)
- better tests for utils.py and office_vba.py

### Fixed

- Bug Fix for [Issue #10](https://github.com/markuskiller/vba-edit/issues/10): Different VBA types are now properly recognised and minimal header sections generated before importing code back into the MS Office VBA editor ensures correct placement of 'Document modules' and 'Class modules' which are all exported as ``.cls`` files. (thanks to [@takutta](https://github.com/takutta) for testing and reporting)
- In previous versions ``.frm`` were exported and code could be edited, however, when imported back to MS Office VBA editor, forms were not processed correctly due to a lack of header handling

## [0.2.1] - 2024-12-09

### Fixed

- WORD: fixed removing of headers before saving exported file to disk and making sure that there is no attempt at removing headers when changed file is reimported`
- While ``excel-vba`` edit does not overwrite files in ``edit`` mode (>=v0.2.0), ``word-vba edit`` still did. That's fixed now.
- WORD & EXCEL: when files are deleted from ``--vba-directory`` in ``edit`` mode those files are now also deleted in VBA editor of the respective office application (which aligns with ``xlwings vba edit`` implementation)
- improved credit to original ``xlwings`` project by adding an inline comment closer to OfficeVBAHandler, which contains the core VBA interaction logic that was inspired by ``xlwings``

## [0.2.0] - 2024-12-08

### Added

- basic tests for entry points ``word-vba`` & ``excel-vba``
- introducing binary distribution cli tools on Windows
- ``excel-vba`` no longer relies on xlwings being installed
- ``xlwings`` is now an optional import
- ``--xlwings|-x`` command-line option to be able to use ``xlwings`` if installed

### Fixed

- on ``*-vba edit`` files are no longer overwritten if ``vba-directory`` is
  already populated (trying to achieve similar behaviour compared to ``xlwings``;
  TODO: files which are deleted from  ``vba-directory`` are not
  automatically deleted in Office VBA editor just yet.)

## [0.1.2] - 2024-12-05

### Added

- Initial release (``word-vba edit|export|import`` & ``excel-vba edit|export|import``: wrapper for ``xlwings vba edit|export|import``)
- Providing semantically consistent entry points for future integration of other MS Office applications (e.g. ``access-vba``, ``powerpoint-vba``, ...)
- (0.1.1) -> (0.1.2) Fix badges and 1 dependency (``chardet``)

# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.1-rc1] - unreleased

### Added

- **Support for Python 3.14**
- **Windows Binaries**: Automated build workflow for Windows executables
  - Includes SHA256 checksums for verification
  - GitHub Attestations provide cryptographic proof of authenticity
  - Software Bill of Materials (SBOM) included with each release
  - See **[Issue #24](https://github.com/markuskiller/vba-edit/issues/24)** for more information on expected 'false positive' categorization of our unsigned binaries by some anti-malware tools 
- **Automatic Header Detection**: Import command now works without manual header flags
  - Automatically finds VBA headers in your code files (VERSION/BEGIN/Attribute lines)
  - Falls back to separate `.header` files if you use those
  - Creates minimal headers if neither exists
  - Import command now warns when both inline and separate header formats exist
  - Clearly states that inline headers take precedence
  - Helps identify conflicting header configurations
- **Colorized Terminal Output**: Modern, professional CLI appearance with automatic color support
  - Success messages appear in green with checkmarks (✓)
  - Error messages appear in red with X marks (✗)
  - Warning messages appear in yellow with warning symbols (⚠)
  - Help text uses syntax highlighting with different colors for options, file names, and section headings
  - Technical terms (TOML, JSON, VBA, Module, Class) highlighted in cyan
  - Example command lines shown in dim gray for easy scanning

### Fixed

- **CI/CD Reliability**: Improved test stability on GitHub Actions runners
  - Added subprocess timeouts to prevent intermittent hangs
  - Increased test timeouts for better reliability on slower CI runners
  - Prevents rare deadlocks when running help commands under load
  - IMPORTANT warnings appear in yellow so you don't miss them
  - Colors match modern tools like uv and ruff for a familiar experience
  - Automatically disabled when output is piped or redirected
  - Use `--no-color` flag to disable if needed
- **Improved Help Messages**: Completely redesigned help output for all commands
  - Options organized into clear groups (File Options, Configuration, Encoding, etc.)
  - Main help shows concise overview with practical examples
  - Command-specific help displays all available options with detailed descriptions
  - Design by [@onderhold](https://github.com/onderhold)
- **Metadata Saving**: New `-m/--save-metadata` option now available in both `edit` and `export` commands
  - Preserve metadata when exporting in edit mode
  - Makes workflows more consistent across commands
- **Simpler Configuration Placeholders**: Streamlined placeholder names for config files
  - Use `{file.name}`, `{file.fullname}`, `{file.path}` instead of `{general.file.*}`
  - Use `{file.vbaproject}` instead of `{vbaproject}`
  - Old placeholders still work (will be removed in v0.5.0)
- **Security Documentation**: Comprehensive security policy and verification guide
  - Clear vulnerability reporting process
  - Multiple verification methods for binaries
  - Best practices for secure usage
- **Automated Dependency Updates**: Dependabot monitors for security updates
  - Weekly checks for outdated dependencies
  - Automatic pull requests for updates
  - Reduces security risk from vulnerable dependencies

### Changed

- **Option Availability**: `--open-folder` now works with `edit`, `import`, and `export` commands
  - Previously only available on some commands
  - Improves workflow by opening results folder after operations
- **Help Display**: Options now only appear where they're relevant
  - Folder options (`--rubberduck-folders`, `--open-folder`) only show on commands that use them
  - Less clutter in help output
  - Easier to find the options you need

### Deprecated

- **Old placeholder format**: `{general.file.*}` and `{vbaproject}` 
  - Still supported for backward compatibility
  - Use new `{file.*}` format instead
  - Will be removed in v0.5.0

### Fixed

- **[Issue #20](https://github.com/markuskiller/vba-edit/issues/20)**: Fixed crash when running `check` command
  - All entry points (excel-vba, word-vba, access-vba, powerpoint-vba) now handle `check` correctly
- **[Issue #21](https://github.com/markuskiller/vba-edit/issues/21)**: Commands no longer show irrelevant options
  - Each command only displays options that actually work with it
  - Prevents confusion about which options to use

### Security

- **GitHub Actions permissions**: Workflows now use minimal required permissions
  - Follows security best practice of least privilege
  - Limits potential damage if workflow is compromised
- **Binary verification**: Multiple methods to verify download authenticity
  - SHA256 checksums included with all releases
  - GitHub Attestations prove binaries were built by official workflow
  - Build provenance cryptographically signed by GitHub
- **Security policy**: Clear process for reporting vulnerabilities
  - Response within 48 hours
  - Private disclosure via GitHub Security Advisories
  - Documented timeline for fixes

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

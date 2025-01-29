# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Support for macro-enabled MS PowerPoint documents
- Option to show program's version number and exit added to all CLI interfaces (`--version`)
- `check all` subcommand for CLI entry points, which processes all suported MS Office apps in a single call (replaces calling `python -m vba_edit.utils`)

### Changed

- Updated build script `create_binaries.py` to create Windows stand-alone binaries for all entry points (``access-vba.exe``, ``excel-vba.exe``, ``powerpoint-vba.exe``, ``word-vba.exe``)
- Handling of ``VBADocumentNames.is_document_module`` as suggested by sourcery.ai
<!-- -  -->

### Fixed

- fix: check for form safety on `export` (if edit command is run without `--save-headers` option, forms cannot be processed correctly -> check for forms and abort if `--save-headers` is not enabled)
- fix: header file handling (`--save-headers`) in already populated `--vba-directory` (only 1 header file was created rather than one per ``*.cls``, ``*.bas`` or ``*.frm`` file) - calling it on empty `--vba-directory` worked as expected

<!-- -  -->
<!-- ### Removed -->
<!-- -  -->
<!-- ### Security -->
<!-- -  -->

## [0.3.0] - 2025-01-19

### Added

- new command ``check``: e.g. ``word-vba check`` detects if 'Trust Access to Word VBA project object model' is enabled (available for all entry points)
- ACCESS: basic support for MS Access databases (standard modules ``*.bas`` and class modules ``*.cls`` are supported)
- ``--save-headers`` command-line option, supporting a comprehensive handling of headers (Issue #11)
- new files created in ``--vba-directory`` are now automatically synced back to MS Office VBA editor in ``*-vba edit`` mode (previously, only file deletions were monitored and synced)
- better tests for utils.py and office_vba.py

### Fixed

- Bug Fix for Issue #10: Different VBA types are now properly recognised and minimal header sections generated before importing code back into the MS Office VBA editor ensures correct placement of 'Document modules' and 'Class modules' which are all exported as ``.cls`` files. (thanks to @takutta for testing and reporting)
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
- introducing script to create binary distribution CLI tools on Windows
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

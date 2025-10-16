import datetime
import json
import logging
import os
import re
import shutil
import sys
import time
from abc import ABC, abstractmethod
from enum import Enum, auto
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

# Third-party imports
import win32com.client
from watchfiles import Change, watch

from vba_edit.exceptions import (
    DocumentClosedError,
    DocumentNotFoundError,
    PathError,
    RPCError,
    VBAAccessError,
    VBAError,
    VBAExportWarning,
    check_rpc_error,
)

# Updated local imports
from vba_edit.path_utils import (
    get_document_paths,
    resolve_path,
)
from vba_edit.utils import (
    get_vba_error_details,
    is_vba_access_error,
)
from vba_edit.console import console


def _filter_attributes(code: str) -> str:
    """Filter out hidden member-level Attribute lines from the given code.

    Attribute directives are exported for WithEvents fields and other hidden members,
    but are illegal when put in Module code verbatim (via AddFromString).

    These are "hidden member attributes" that:
    - Are legal in exported VBA files (.cls, .bas, .frm)
    - Are illegal when written directly into a VBA module
    - Cause syntax errors if they appear in the VBE after import

    Examples of hidden member attributes that need filtering (note the dot before VB_*):
    - Attribute MyCtrl.VB_VarHelpID = -1 (WithEvents controls)
    - Attribute mField.VB_VarDescription = "..." (member descriptions)
    - Attribute Prop.VB_UserMemId = 0 (default member at field level)

    Module-level attributes (without dots) are preserved:
    - Attribute VB_Name = "MyModule" (module name)
    - Attribute VB_Exposed = True (class exposure)
    - Attribute VB_GlobalNameSpace = False (module global namespace)

    The key difference: hidden member attributes contain a dot (.) before the VB_* name,
    while module-level attributes do not.

    Reference: https://vbaplanet.com/attributes.php
    Issue: https://github.com/markuskiller/vba-edit/issues/16

    Args:
        code: VBA code that may contain illegal hidden member Attribute lines

    Returns:
        Code with hidden member Attribute lines filtered out, module-level attributes preserved
    """
    if not code:
        return code

    lines = code.split("\n")
    filtered_lines = []
    for line in lines:
        stripped = line.strip().lower()
        # Filter lines that start with "attribute " and contain a dot before "="
        # This identifies hidden member attributes like "Attribute MyCtrl.VB_VarHelpID = -1"
        # while preserving module-level attributes like "Attribute VB_Exposed = False"
        if stripped.startswith("attribute ") and "." in stripped.split("=")[0]:
            continue  # Skip this line (it's a hidden member attribute)
        filtered_lines.append(line)
    return "\n".join(filtered_lines)


"""
The VBA import/export/edit functionality is based on the excellent work done by the xlwings project
(https://github.com/xlwings/xlwings) which is distributed under the BSD 3-Clause License:

Copyright (c) 2014-present, Zoomer Analytics GmbH.
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this
  list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice,
  this list of conditions and the following disclaimer in the documentation
  and/or other materials provided with the distribution.

* Neither the name of the copyright holder nor the names of its
  contributors may be used to endorse or promote products derived from
  this software without specific prior written permission.

This module extends the original xlwings VBA interaction concept to provide a consistent 
interface for interacting with VBA code across different Microsoft Office applications.
"""

# Configure module logger
logger = logging.getLogger(__name__)


# Office app configuration
# Primary macro-enabled extensions (used in examples/docs)
OFFICE_MACRO_EXTENSIONS: Dict[str, str] = {
    "word": ".docm",
    "excel": ".xlsm",
    "access": ".accdb",
    "powerpoint": ".pptm",
    # Potential future support
    # "outlook": ".otm",
    # "project": ".mpp",
    # "visio": ".vsdm",
}

# All file extensions that can contain VBA macros (for file type detection)
# Maps file extension to Office application type
ALL_VBA_EXTENSIONS: Dict[str, str] = {
    # Word formats
    ".docm": "word",  # Word macro-enabled document (2007+)
    ".doc": "word",  # Word 97-2003 document (legacy, supports macros)
    ".dotm": "word",  # Word macro-enabled template (2007+)
    ".dot": "word",  # Word 97-2003 template (legacy, supports macros)
    # Excel formats
    ".xlsm": "excel",  # Excel macro-enabled workbook (2007+)
    ".xlsb": "excel",  # Excel binary workbook (2007+, supports macros)
    ".xls": "excel",  # Excel 97-2003 workbook (legacy, supports macros)
    ".xltm": "excel",  # Excel macro-enabled template (2007+)
    ".xlt": "excel",  # Excel 97-2003 template (legacy, supports macros)
    ".xlam": "excel",  # Excel add-in (macro-enabled)
    ".xla": "excel",  # Excel 97-2003 add-in (legacy, supports macros)
    # PowerPoint formats
    ".pptm": "powerpoint",  # PowerPoint macro-enabled presentation (2007+)
    ".ppt": "powerpoint",  # PowerPoint 97-2003 presentation (legacy, supports macros)
    ".potm": "powerpoint",  # PowerPoint macro-enabled template (2007+)
    ".pot": "powerpoint",  # PowerPoint 97-2003 template (legacy, supports macros)
    ".ppam": "powerpoint",  # PowerPoint add-in (macro-enabled)
    ".ppa": "powerpoint",  # PowerPoint 97-2003 add-in (legacy, supports macros)
    # Access formats
    ".accdb": "access",  # Access database (2007+)
    ".mdb": "access",  # Access 97-2003 database (legacy)
    ".accde": "access",  # Access execute-only database (compiled)
    ".mde": "access",  # Access 97-2003 execute-only database (legacy, compiled)
}

# Command-line entry points for different Office applications
OFFICE_CLI_NAMES = {app: f"{app}-vba" for app in OFFICE_MACRO_EXTENSIONS.keys()}

# Regex pattern for Rubberduck @Folder annotations
RUBBERDUCK_FOLDER_PATTERN = re.compile(r"'\s*@folder\s*(?:\(\s*)?[\"']([^\"']+)[\"']\s*(?:\))?\s*$", re.IGNORECASE)

# Currently supported apps in vba-edit
# "access" is only partially supported at this stage and will be included
# in list as soon as tests are adapted to handle it
SUPPORTED_APPS = [
    "word",
    "excel",
    "powerpoint",
    "access",
]


class VBADocumentNames:
    """Document module names across different languages."""

    # Excel document module names
    EXCEL_WORKBOOK_NAMES = {
        "ThisWorkbook",  # English
        "DieseArbeitsmappe",  # German
        "CeClasseur",  # French
        "EstaLista",  # Spanish
        "QuestoFoglio",  # Italian
        "EstaLista",  # Portuguese
        "このブック",  # Japanese
        "本工作簿",  # Chinese Simplified
        "本活頁簿",  # Chinese Traditional
        "이통합문서",  # Korean
        "ЭтаКнига",  # Russian
    }

    # Excel worksheet module prefixes
    EXCEL_SHEET_PREFIXES = {
        "Sheet",  # English
        "Tabelle",  # German
        "Feuil",  # French
        "Hoja",  # Spanish
        "Foglio",  # Italian
        "Planilha",  # Portuguese
        "シート",  # Japanese
        "工作表",  # Chinese Simplified/Traditional
        "시트",  # Korean
        "Лист",  # Russian
    }

    # Word document module names
    WORD_DOCUMENT_NAMES = {
        "ThisDocument",  # English
        "DiesesDokument",  # German
        "CeDocument",  # French
        "EsteDocumento",  # Spanish/Portuguese
        "QuestoDocumento",  # Italian
        "この文書",  # Japanese
        "本文檔",  # Chinese Traditional
        "本文档",  # Chinese Simplified
        "이문서",  # Korean
        "ЭтотДокумент",  # Russian
    }

    # PowerPoint slide module prefixes
    POWERPOINT_SLIDE_PREFIXES = {
        "Slide",  # English
        "Folie",  # German
        "Diapo",  # French
        "Diapositiva",  # Spanish/Italian
        "Slide",  # Portuguese
        "スライド",  # Japanese
        "投影片",  # Chinese Traditional
        "幻灯片",  # Chinese Simplified
        "슬라이드",  # Korean
        "Слайд",  # Russian
    }

    @classmethod
    def is_document_module(cls, name: str) -> bool:
        """Check if a name matches any known document module name."""
        # Handle standard document modules (Excel/Word)
        if name in cls.EXCEL_WORKBOOK_NAMES or name in cls.WORD_DOCUMENT_NAMES:
            return True

        # Handle Excel sheets
        if any(name.startswith(prefix) and name[len(prefix) :].isdigit() for prefix in cls.EXCEL_SHEET_PREFIXES):
            return True

        # Handle PowerPoint slides
        if any(name.startswith(prefix) and name[len(prefix) :].isdigit() for prefix in cls.POWERPOINT_SLIDE_PREFIXES):
            return True

        return False


# VBA type definitions and constants
class VBAModuleType(Enum):
    """VBA module types"""

    DOCUMENT = auto()  # ThisWorkbook/ThisDocument modules
    CLASS = auto()  # Regular class modules
    STANDARD = auto()  # Standard modules (.bas)
    FORM = auto()  # UserForm modules


class VBATypes:
    """Constants for VBA component types"""

    VBEXT_CT_DOCUMENT = 100  # Document module type
    VBEXT_CT_MSFORM = 3  # UserForm type
    VBEXT_CT_STDMODULE = 1  # Standard module type
    VBEXT_CT_CLASSMODULE = 2  # Class module type

    # Application specific constants
    XL_WORKSHEET = -4167  # xlWorksheet type for Excel

    # Map module types to file extensions and metadata
    TYPE_TO_EXT = {
        VBEXT_CT_STDMODULE: ".bas",  # Standard Module
        VBEXT_CT_CLASSMODULE: ".cls",  # Class Module
        VBEXT_CT_MSFORM: ".frm",  # MSForm
        VBEXT_CT_DOCUMENT: ".cls",  # Document Module
    }

    TYPE_INFO = {
        VBEXT_CT_STDMODULE: {
            "type_name": "Standard Module",
            "extension": ".bas",
            "cls_header": False,
        },
        VBEXT_CT_CLASSMODULE: {
            "type_name": "Class Module",
            "extension": ".cls",
            "cls_header": True,
        },
        VBEXT_CT_MSFORM: {
            "type_name": "UserForm",
            "extension": ".frm",
            "cls_header": True,
        },
        VBEXT_CT_DOCUMENT: {
            "type_name": "Document Module",
            "extension": ".cls",
            "cls_header": True,
        },
    }


class VBAComponentHandler:
    """Handles VBA component operations independent of Office application type.

    This class provides core functionality for managing VBA components, including
    analyzing module types, handling headers, and preparing content for import/export
    operations. It serves as a utility class for the main Office-specific handlers.
    """

    def __init__(self, use_rubberduck_folders: bool = False):
        """Initialize the component handler.

        Args:
            use_rubberduck_folders: Whether to process Rubberduck folder annotations
        """
        self.use_rubberduck_folders = use_rubberduck_folders

    def get_component_info(self, component: Any) -> Dict[str, Any]:
        """Get detailed information about a VBA component.

        Analyzes a VBA component and returns metadata including its type,
        line count, and appropriate file extension.

        Args:
            component: A VBA component object from any Office application

        Returns:
            Dict containing component metadata with the following keys:
                - name: Component name
                - type: VBA type code
                - type_name: Human-readable type name
                - extension: Appropriate file extension
                - code_lines: Number of lines of code
                - has_cls_header: Whether component requires a class header

        Raises:
            VBAError: If component information cannot be retrieved
        """
        try:
            # Get code line count safely
            code_lines = component.CodeModule.CountOfLines if hasattr(component, "CodeModule") else 0

            # Get type info or use defaults for unknown types
            type_data = VBATypes.TYPE_INFO.get(
                component.Type, {"type_name": "Unknown", "extension": ".txt", "cls_header": False}
            )

            return {
                "name": component.Name,
                "type": component.Type,
                "type_name": type_data["type_name"],
                "extension": type_data["extension"],
                "code_lines": code_lines,
                "has_cls_header": type_data["cls_header"],
            }
        except Exception as e:
            logger.error(f"Failed to get component info for {component.Name}: {str(e)}")
            raise VBAError(f"Failed to analyze component {component.Name}") from e

    def determine_cls_type(self, header: str) -> VBAModuleType:
        """Determine if a .cls file is a document module or regular class module.

        Analyzes the VBA component header to determine its exact type based on
        the presence and values of specific attributes.

        Args:
            header: Content of the VBA component header

        Returns:
            VBAModuleType.DOCUMENT or VBAModuleType.CLASS based on header analysis
        """
        # Extract key attributes
        predeclared = re.search(r"Attribute VB_PredeclaredId = (\w+)", header)
        exposed = re.search(r"Attribute VB_Exposed = (\w+)", header)

        # Document modules have both attributes set to True
        if predeclared and exposed and predeclared.group(1).lower() == "true" and exposed.group(1).lower() == "true":
            return VBAModuleType.DOCUMENT

        return VBAModuleType.CLASS

    def get_module_type(self, file_path: Path, in_file_headers: bool = False, encoding: str = "utf-8") -> VBAModuleType:
        """Determine VBA module type from file extension and content.

        Args:
            file_path: Path to the VBA module file
            in_file_headers: Whether headers are embedded in files
            encoding: Character encoding for reading files

        Returns:
            Appropriate VBAModuleType

        Raises:
            ValueError: If file extension is unknown
        """
        suffix = file_path.suffix.lower()
        name = file_path.stem

        # Check if it's a known document module name in any language
        if VBADocumentNames.is_document_module(name):
            return VBAModuleType.DOCUMENT

        if suffix == ".bas":
            return VBAModuleType.STANDARD
        elif suffix == ".frm":
            return VBAModuleType.FORM
        elif suffix == ".cls":
            # For .cls files, check the header if available
            if in_file_headers:
                # When using in-file headers, check the file content directly
                try:
                    with open(file_path, "r", encoding=encoding) as f:
                        content = f.read()
                    header, _ = self.split_vba_content(content)
                    if header:
                        return self.determine_cls_type(header)
                except Exception:
                    logger.debug(f"Could not read content from {file_path}, treating as regular class module")
            else:
                # Use separate header file
                header_file = file_path.with_suffix(".header")
                if header_file.exists():
                    with open(header_file, "r", encoding=encoding) as f:
                        return self.determine_cls_type(f.read())

            logger.debug(f"No header file found for {file_path}, treating as regular class module")
            return VBAModuleType.CLASS

        raise ValueError(f"Unknown file extension: {suffix}")

    def split_vba_content(self, content: str) -> Tuple[str, str]:
        """Split VBA content into header and code sections.

        Args:
            content: Complete VBA component content

        Returns:
            Tuple of (header, code)

        Note:
            Headers include VERSION, BEGIN/END blocks, and module-level Attribute VB_ lines.
            Only module-level attributes (VB_Name, VB_GlobalNameSpace, VB_Creatable,
            VB_PredeclaredId, VB_Exposed) are considered part of the header.
            Procedure-level attributes are considered part of the code.
        """
        if not content.strip():
            return "", ""

        lines = content.splitlines()
        last_header_idx = -1
        in_begin_block = False

        for i, line in enumerate(lines):
            stripped = line.strip()

            # Check for header components
            if stripped.startswith("VERSION"):
                last_header_idx = i
            elif stripped.upper() == "BEGIN" or stripped.upper().startswith("BEGIN "):
                # Matches both "BEGIN" (class modules) and "Begin {GUID} UserFormName" (UserForms)
                in_begin_block = True
                last_header_idx = i
            elif stripped.upper() == "END" and in_begin_block:
                in_begin_block = False
                last_header_idx = i
            elif in_begin_block:
                # Lines inside BEGIN/END block (like MultiUse or form properties)
                last_header_idx = i
            elif stripped.startswith("Attribute VB_"):
                last_header_idx = i
            elif last_header_idx >= 0 and not stripped.startswith("Attribute VB_"):
                # First non-header line after we've seen headers
                break

        if last_header_idx == -1:
            return "", content

        header = "\n".join(lines[: last_header_idx + 1])
        code = "\n".join(lines[last_header_idx + 1 :])

        return header.strip(), code.strip()

    def has_inline_headers(self, file_path: Path, encoding: str = "utf-8") -> bool:
        """Detect if a code file contains embedded VBA headers.

        Checks if the file starts with header markers like VERSION, BEGIN, or Attribute.
        This allows automatic detection of header format during import.

        The detection is strict to avoid false positives from comments:
        - VERSION and BEGIN must be at start of non-comment, non-blank lines
        - Attribute VB_ must be at line start (not in comments)
        - Only checks first 10 lines for performance

        Args:
            file_path: Path to the code file to check
            encoding: File encoding to use

        Returns:
            True if file contains inline headers, False otherwise
        """
        try:
            with open(file_path, "r", encoding=encoding) as f:
                # Read first few lines to check for header markers
                first_lines = []
                for i, line in enumerate(f):
                    if i >= 10:  # Headers are always at the top
                        break
                    first_lines.append(line)  # Keep original line with whitespace

            # Check for header markers - must be actual code, not comments
            for line in first_lines:
                stripped = line.strip()

                # Skip blank lines
                if not stripped:
                    continue

                # Skip comment lines (VBA comments start with ' or REM)
                if stripped.startswith("'") or stripped.upper().startswith("REM "):
                    continue

                # Check for VERSION marker (case-insensitive, must have space or be exact)
                if stripped.upper().startswith("VERSION "):
                    logger.debug(f"Detected inline headers in {file_path.name} (found VERSION marker)")
                    return True

                # Check for BEGIN marker (case-insensitive)
                # Matches: "BEGIN" or "Begin {GUID} FormName"
                if stripped.upper() == "BEGIN" or stripped.upper().startswith("BEGIN "):
                    logger.debug(f"Detected inline headers in {file_path.name} (found BEGIN marker)")
                    return True

                # Check for Attribute VB_ (case-sensitive for VB_, as per VBA convention)
                if stripped.startswith("Attribute VB_"):
                    logger.debug(f"Detected inline headers in {file_path.name} (found VB attribute)")
                    return True

            logger.debug(f"No inline headers detected in {file_path.name}")
            return False

        except Exception as e:
            logger.debug(f"Error checking for inline headers in {file_path}: {e}", exc_info=True)
            return False

    def create_minimal_header(self, name: str, module_type: VBAModuleType) -> str:
        """Create a minimal header for a VBA component.

        Args:
            name: Name of the VBA component
            module_type: Type of the VBA module

        Returns:
            Minimal valid header for the component type
        """
        if module_type == VBAModuleType.CLASS:
            # Class modules need the class declaration and standard attributes
            header = [
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                f'Attribute VB_Name = "{name}"',
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = False",
                "Attribute VB_Exposed = False",
            ]
        elif module_type == VBAModuleType.FORM:
            # UserForm requires specific form structure and GUID
            # {C62A69F0-16DC-11CE-9E98-00AA00574A4F} is the standard UserForm GUID
            header = [
                "VERSION 5.00",
                "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " + name,
                f'   Caption         =   "{name}"',
                "   ClientHeight    =   3000",
                "   ClientLeft      =   100",
                "   ClientTop       =   400",
                "   ClientWidth     =   4000",
                '   OleObjectBlob   =   "' + name + '.frx":0000',
                "   StartUpPosition =   1  'CenterOwner",
                "End",
                f'Attribute VB_Name = "{name}"',
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = True",
                "Attribute VB_Exposed = False",
            ]
            logger.info(
                f"Created minimal header for UserForm: {name} \n"
                "Consider using the command-line option --save-headers "
                "in order not to lose previously specified form structure and GUID."
            )
        else:
            # Standard modules only need the name
            header = [f'Attribute VB_Name = "{name}"']

        return "\n".join(header)

    def prepare_import_content(
        self, name: str, module_type: VBAModuleType, header: str, code: str, in_file_headers: bool = False
    ) -> str:
        """Prepare content for VBA component import.

        Args:
            name: Name of the VBA component
            module_type: Type of the VBA module
            header: Header content (may be empty)
            code: Code content
            in_file_headers: Whether headers are embedded in files

        Returns:
            Properly formatted content for import
        """
        if in_file_headers:
            # When in_file_headers is True, return only the code part for AddFromString
            # The header handling will be done differently (via COM properties, not AddFromString)
            return code
        else:
            if not header and module_type == VBAModuleType.STANDARD:
                header = self.create_minimal_header(name, module_type)

        return f"{header}\n{code}\n" if header else f"{code}\n"

    def validate_component_header(self, header: str, expected_type: VBAModuleType) -> bool:
        """Validate that a component's header matches its expected type.

        Args:
            header: Header content to validate
            expected_type: Expected module type

        Returns:
            True if header is valid for the expected type
        """
        if not header:
            return expected_type == VBAModuleType.STANDARD

        actual_type = self.determine_cls_type(header)

        if expected_type == VBAModuleType.DOCUMENT:
            return actual_type == VBAModuleType.DOCUMENT

        return True  # Other types are less strict about headers

    def _update_module_content(self, component: Any, content: str) -> None:
        """Update the content of an existing module.

        When updating content directly in the VBA editor (without full import),
        we must not include header information as it can't be processed by
        the VBA project.

        Args:
            component: VBA component to update
            content: New content to set
        """
        try:
            # For direct updates, we want just the code without any header
            # manipulation - the existing module already has its header
            if component.CodeModule.CountOfLines > 0:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)

            # Filter out hidden member attributes (Issue #16)
            if content.strip():
                component.CodeModule.AddFromString(_filter_attributes(content))

            logger.debug(f"Updated content for: {component.Name}")
        except Exception as e:
            logger.error(f"Failed to update content for {component.Name}: {str(e)}")
            raise VBAError("Failed to update module content") from e

    def get_rubberduck_folder(self, code: str) -> Tuple[str, str]:
        """Find Rubberduck @Folder in VBA code.

        Supports various Rubberduck annotation syntaxes:
        - '@Folder "MyFolder"
        - '@Folder("MyFolder")
        - '@folder "My.Nested.Folder"
        - Case insensitive matching

        Only scans leading comment lines, stopping at the first non-comment/non-whitespace line,
        or after finding the first folder annotation, just as RubberDuckVBA does.
        Returns the folder path and the original code (unmodified).

        Args:
            code: VBA code content

        Returns:
            Tuple of (folder_path, code)
        """
        if not self.use_rubberduck_folders:
            return "", code

        lines = code.splitlines()
        folder_path = ""

        for line in lines:
            stripped = line.strip()
            if not stripped or stripped.startswith("'"):
                # Look for @Folder annotation
                match = RUBBERDUCK_FOLDER_PATTERN.match(stripped)
                if match:
                    # Extract folder path and convert dot notation to filesystem path
                    folder_path = match.group(1).replace(".", os.sep)
                    # Found the folder annotation, no need to continue
                    break
                continue
            # Stop at first non-comment/non-whitespace line
            break

        # Return the folder path and the original code
        return folder_path, code

    def add_rubberduck_folder(self, code: str, folder_path: str) -> str:
        """Add Rubberduck @Folder annotation to VBA code.

        If a @Folder annotation already exists, it will be updated with the new path.
        If no annotation exists, a new one will be added.

        Args:
            code: VBA code content
            folder_path: Folder path to add

        Returns:
            Code with @Folder annotation added or updated
        """
        if not self.use_rubberduck_folders or not folder_path:
            return code

        # Convert filesystem path to Rubberduck notation
        rubberduck_path = folder_path.replace(os.sep, ".")
        folder_annotation = f'\'@Folder("{rubberduck_path}")'

        lines = code.splitlines()

        # Find insertion point (after attributes, before actual code) and check for existing @Folder annotations
        insert_index = 0
        existing_folder_line = -1

        for i, line in enumerate(lines):
            stripped = line.strip()

            # Skip empty lines at the beginning
            if not stripped:
                continue

            # Check if we're still in attributes section
            if stripped.startswith("Attribute ") or stripped.startswith("VERSION ") or stripped.startswith("BEGIN"):
                insert_index = i + 1
                continue

            # Check for existing @Folder annotation
            if RUBBERDUCK_FOLDER_PATTERN.match(stripped):
                existing_folder_line = i
                continue

            # This is where we want to insert if no existing annotation
            if existing_folder_line == -1:
                insert_index = i
            break

        # Update existing annotation or add new one
        if existing_folder_line >= 0:
            # Replace existing annotation
            lines[existing_folder_line] = folder_annotation
        else:
            # Insert new annotation
            lines.insert(insert_index, folder_annotation)

        return "\n".join(lines)

    def get_folder_from_file_path(self, file_path: Path, vba_base_dir: Path) -> str:
        """Extract folder path from file system location.

        Args:
            file_path: Path to the VBA file
            vba_base_dir: Base VBA directory

        Returns:
            Relative folder path
        """
        if not self.use_rubberduck_folders:
            return ""

        try:
            relative_path = file_path.relative_to(vba_base_dir)
            folder_path = str(relative_path.parent)

            # Return empty string for root directory
            if folder_path == ".":
                return ""

            return folder_path
        except ValueError:
            # File is not under vba_base_dir
            return ""


class OfficeVBAHandler(ABC):
    """Abstract base class for handling VBA operations across Office applications.

    This class provides the foundation for application-specific VBA handlers,
    implementing common functionality while requiring specific implementations
    for application-dependent operations.

    Args:
        doc_path: Path to the Office document
        vba_dir: Directory for VBA files (defaults to current directory)
        encoding: Character encoding for VBA files (default: cp1252)
        verbose: Enable verbose logging (default: False)
        save_headers: Save VBA headers to separate .header files (default: False)
        use_rubberduck_folders: Organize by RubberduckVBA @Folder annotations (default: False)
        in_file_headers: Embed headers in code files instead of separate files (default: False)
        open_folder: Open VBA directory after export (default: False)

    Attributes:
        app: Office application COM object
        doc: Office document COM object
        component_handler: Utility handler for VBA components
    """

    def __init__(
        self,
        doc_path: str,
        vba_dir: Optional[str] = None,
        encoding: str = "cp1252",
        verbose: bool = False,
        save_headers: bool = False,
        use_rubberduck_folders: bool = False,
        open_folder: bool = False,
        in_file_headers: bool = False,
    ):
        """Initialize the VBA handler."""
        try:
            # Let DocumentNotFoundError propagate as is - it's more fundamental than VBA errors
            self.doc_path, self.vba_dir = get_document_paths(doc_path, None, vba_dir)
            self.encoding = encoding
            self.verbose = verbose
            self.save_headers = save_headers
            self.use_rubberduck_folders = use_rubberduck_folders
            self.open_folder = open_folder
            self.in_file_headers = in_file_headers
            self.app = None
            self.doc = None
            self.component_handler = VBAComponentHandler(use_rubberduck_folders)

            # Configure logging
            log_level = logging.DEBUG if verbose else logging.INFO
            logger.setLevel(log_level)

            logger.debug(f"Initialized {self.__class__.__name__} with document: {doc_path}")
            logger.debug(f"VBA directory: {self.vba_dir}")
            logger.debug(f"Using encoding: {encoding}")
            logger.debug(f"Save headers: {save_headers}")
            logger.debug(f"In-file headers: {in_file_headers}")
            logger.debug(f"Rubberduck folders: {use_rubberduck_folders}")
            logger.debug(f"Open folder after export: {open_folder}")

        except DocumentNotFoundError:
            raise  # Let it propagate
        except Exception as e:
            raise VBAError(f"Failed to initialize VBA handler: {str(e)}") from e

    @property
    @abstractmethod
    def app_name(self) -> str:
        """Name of the Office application."""
        pass

    @property
    @abstractmethod
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        pass

    @property
    def document_type(self) -> str:
        """Get the document type string for error messages."""
        return "workbook" if self.app_name == "Excel" else "document"

    def get_vba_project(self) -> Any:
        """Get VBA project based on application type."""
        logger.debug("Getting VBA project...")
        try:
            if self.app_name == "Access":
                vba_project = self.app.VBE.ActiveVBProject
            else:
                try:
                    vba_project = self.doc.VBProject
                except Exception as e:
                    if is_vba_access_error(e):
                        details = get_vba_error_details(e)

                        # Available details:
                        # details['hresult']
                        # details['message']
                        # details['source']
                        # details['description']
                        # details['scode']

                        raise VBAAccessError(
                            f"Cannot access VBA project in {details['source']}. "
                            f"Error: {details['description']}\n"
                            f"Please ensure 'Trust access to the VBA project object model' "
                            f"is enabled in Trust Center Settings."
                        ) from e
                    raise

            if vba_project is None:
                raise VBAAccessError(
                    f"Cannot access VBA project in {self.app_name}. "
                    "Please ensure 'Trust access to the VBA project object model' "
                    "is enabled in Trust Center Settings."
                )

            logger.debug("VBA project accessed successfully")
            return vba_project

        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(self.app_name)

            if isinstance(e, VBAAccessError):
                raise

            logger.error(f"Failed to access VBA project: {str(e)}")
            raise VBAError(f"Failed to access VBA project in {self.app_name}: {str(e)}") from e

    @abstractmethod
    def get_document_module_name(self) -> str:
        """Get the name of the document module (e.g., ThisDocument, ThisWorkbook)."""
        pass

    def is_document_open(self) -> bool:
        """Check if the document is still open and accessible."""
        try:
            if self.doc is None:
                return False

            # Try to access document name
            name = self.doc.Name
            if callable(name):  # Handle Mock case in tests
                name = name()

            # Check if document is still active
            return self.doc.FullName == str(self.doc_path)

        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            raise DocumentClosedError(self.document_type)

    def initialize_app(self) -> None:
        """Initialize the Office application."""
        try:
            if self.app is None:
                logger.debug(f"Initializing {self.app_name} application")
                self.app = win32com.client.Dispatch(self.app_progid)
                if self.app_name != "Access":
                    self.app.Visible = True
        except Exception as e:
            error_msg = f"Failed to initialize {self.app_name} application"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def _check_form_safety(self, vba_dir: Path) -> None:
        """Check if there are .frm files when headers are disabled.

        Args:
            vba_dir: Directory to check for .frm files

        Raises:
            VBAError: If .frm files are found and neither save_headers nor in_file_headers is enabled
        """
        if not self.save_headers and not self.in_file_headers:
            form_files = list(vba_dir.glob("*.frm"))
            if form_files:
                form_names = ", ".join(f.stem for f in form_files)
                error_msg = (
                    f"\nERROR: Found UserForm files ({form_names}) but preferred header option is not enabled!\n"
                    f"UserForms require their full header information to maintain form specifications.\n"
                    f"Please re-run the command with --in-file-headers or --save-headers flag to preserve form settings."
                )
                logger.error(error_msg)
                sys.exit(1)

    def _check_file_extension_mismatch(self) -> None:
        """Check if the file extension matches the Office application type.

        Raises:
            VBAError: If the file extension doesn't match the expected type for this application.
        """
        file_ext = self.doc_path.suffix.lower()

        # Check if this extension belongs to a different Office app
        # Use ALL_VBA_EXTENSIONS for comprehensive coverage (includes legacy formats)
        if file_ext in ALL_VBA_EXTENSIONS:
            correct_app = ALL_VBA_EXTENSIONS[file_ext]

            # Only raise error if it's a mismatch (not the current app)
            if correct_app != self.app_name.lower():
                correct_cli = OFFICE_CLI_NAMES[correct_app]
                current_cli = OFFICE_CLI_NAMES[self.app_name.lower()]

                # Use proper article (a/an) based on first letter
                article = "an" if correct_app[0].lower() in "aeiou" else "a"

                # Build colorized error message using Rich markup
                error_lines = [
                    "",
                    f"File extension mismatch: '[option]{file_ext}[/option]' files cannot be opened with {self.app_name}.",
                    f"The file '[path]{self.doc_path.name}[/path]' appears to be {article} [tech]{correct_app.capitalize()}[/tech] file.",
                    "",
                    "Please use the correct entry point:",
                    f"  [error]✗[/error] Current:  [dim]{current_cli}.exe[/dim]",
                    f"  [success]✓[/success] Correct:  [command]{correct_cli}.exe[/command]",
                    "",
                    "Example:",
                    f'  [command]{correct_cli}.exe[/command] export -f [path]"{self.doc_path.name}"[/path] --vba-directory src',
                ]

                # Render the colorized message to a string
                with console.capture() as capture:
                    for line in error_lines:
                        console.print(line, highlight=False)

                raise VBAError(capture.get())

    def open_document(self) -> None:
        """Open the Office document."""
        try:
            if self.doc is None:
                # Check for file extension mismatch before attempting to open
                self._check_file_extension_mismatch()

                self.initialize_app()
                logger.debug(f"Opening document: {self.doc_path}")
                self.doc = self._open_document_impl()
        except Exception as e:
            error_msg = f"Failed to open document: {self.doc_path}"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    @abstractmethod
    def _open_document_impl(self) -> Any:
        """Implementation-specific document opening logic."""
        pass

    def save_document(self) -> None:
        """Save the document if it's open."""
        if self.doc is not None:
            try:
                self.doc.Save()
                logger.info("Document has been saved and left open for further editing")
            except Exception as e:
                raise VBAError("Failed to save document") from e

    def close_document(self) -> None:
        """Close the currently open document and application if no other documents are open.

        This method closes the document without saving and sets self.doc to None.
        If there are no other open documents (besides the one being closed),
        it also closes the application.
        It's primarily used after export operations to clean up resources.
        """
        if self.doc is not None:
            try:
                # Check document count BEFORE closing to decide if we should close app
                # We need to check if count will be 0 after closing this document
                should_close_app = False
                if self.app is not None:
                    try:
                        if self.app_name == "Access":
                            # Access: will close app after closing database
                            should_close_app = True  # Always close Access after closing database
                        elif self.app_name == "Word":
                            # Word: close app if this is the only document (count == 1)
                            should_close_app = self.app.Documents.Count <= 1
                        elif self.app_name == "Excel":
                            # Excel: close app if this is the only workbook (count == 1)
                            should_close_app = self.app.Workbooks.Count <= 1
                        elif self.app_name == "PowerPoint":
                            # PowerPoint: close app if this is the only presentation (count == 1)
                            should_close_app = self.app.Presentations.Count <= 1

                        # Log the document count
                        if self.app_name == "Word":
                            logger.debug(f"Document count before closing: {self.app.Documents.Count}")
                        elif self.app_name == "Excel":
                            logger.debug(f"Workbook count before closing: {self.app.Workbooks.Count}")
                        elif self.app_name == "PowerPoint":
                            logger.debug(f"Presentation count before closing: {self.app.Presentations.Count}")
                    except Exception as e:
                        logger.debug(f"Could not check document count: {str(e)}")
                        # Default to not closing app if we can't determine count
                        should_close_app = False

                logger.debug(f"Closing {self.document_type}: {self.doc_path}")
                # Close without saving (SaveChanges=False / wdDoNotSaveChanges=0)
                if self.app_name == "Access":
                    # Access handles close differently
                    self.app.CloseCurrentDatabase()
                else:
                    # Excel, Word, PowerPoint use Close method
                    self.doc.Close(SaveChanges=False)
                self.doc = None
                logger.info(f"{self.document_type.capitalize()} closed successfully")

                # Now close the application if appropriate
                if should_close_app and self.app is not None:
                    try:
                        logger.debug(f"No other documents open, closing {self.app_name} application")
                        self.app.Quit()
                        self.app = None
                        logger.info(f"{self.app_name} application closed successfully")
                    except Exception as e:
                        logger.warning(f"Failed to close {self.app_name} application: {str(e)}")
                        # Don't raise - closing app is not critical
                elif not should_close_app:
                    logger.debug(f"Other documents are open, keeping {self.app_name} application running")

            except Exception as e:
                logger.warning(f"Failed to close {self.document_type}: {str(e)}")
                # Don't raise exception - closing is not critical
                # Set to None anyway to avoid stale references
                self.doc = None

    def _check_header_mode_change(self) -> bool:
        """Check if the header storage mode has changed since last export.

        Returns:
            bool: True if header mode has changed or no metadata exists
        """
        metadata_path = self.vba_dir / "vba_metadata.json"
        if not metadata_path.exists():
            return False  # No metadata, first export

        try:
            with open(metadata_path, "r", encoding="utf-8") as f:
                metadata = json.load(f)

            old_mode = metadata.get("header_mode", "none")
            new_mode = "inline" if self.in_file_headers else ("separate" if self.save_headers else "none")

            if old_mode != new_mode:
                logger.debug(f"Header mode changed: {old_mode} -> {new_mode}")
                return True

            return False

        except Exception as e:
            logger.warning(f"Could not read metadata: {str(e)}")
            return False  # If we can't read metadata, don't force overwrite

    def _get_header_modes(self) -> Tuple[str, str]:
        """Get old and new header modes for display.

        Returns:
            Tuple: (old_mode_description, new_mode_description)
        """
        metadata_path = self.vba_dir / "vba_metadata.json"
        old_mode = "none (code only)"

        if metadata_path.exists():
            try:
                with open(metadata_path, "r", encoding="utf-8") as f:
                    metadata = json.load(f)
                mode = metadata.get("header_mode", "none")
                if mode == "inline":
                    old_mode = "inline (headers embedded in code files)"
                elif mode == "separate":
                    old_mode = "separate (headers in .header files)"
                else:
                    old_mode = "none (code only, no headers saved)"
            except Exception:
                pass

        # Determine new mode
        if self.in_file_headers:
            new_mode = "inline (headers embedded in code files)"
        elif self.save_headers:
            new_mode = "separate (headers in .header files)"
        else:
            new_mode = "none (code only, no headers saved)"

        return old_mode, new_mode

    def _cleanup_old_header_files(self) -> None:
        """Clean up old .header files when switching header storage modes.

        This prevents orphaned .header files from remaining when switching
        from --save-headers to --in-file-headers or no headers.
        """
        try:
            # Find all .header files
            header_files = list(self.vba_dir.glob("*.header"))
            if self.use_rubberduck_folders:
                header_files.extend(self.vba_dir.rglob("*.header"))

            if header_files:
                logger.info(f"Cleaning up {len(header_files)} old .header file(s)...")
                for header_file in header_files:
                    try:
                        header_file.unlink()
                        logger.debug(f"Removed old header file: {header_file.name}")
                    except OSError as e:
                        logger.warning(f"Could not remove {header_file.name}: {e}")
        except Exception as e:
            logger.warning(f"Error during header file cleanup: {e}")
            # Don't fail the export if cleanup fails

    def _check_existing_vba_files(self) -> list:
        """Check if VBA files already exist in the export directory.

        Returns:
            list: List of existing VBA file paths
        """
        existing_files = []
        vba_extensions = [".bas", ".cls", ".frm"]

        try:
            if self.use_rubberduck_folders:
                # Check recursively
                for ext in vba_extensions:
                    existing_files.extend(self.vba_dir.rglob(f"*{ext}"))
            else:
                # Check only root directory
                for ext in vba_extensions:
                    existing_files.extend(self.vba_dir.glob(f"*{ext}"))

            return existing_files
        except Exception as e:
            logger.debug(f"Error checking for existing files: {e}")
            return []

    def _save_metadata(self, encodings: Dict[str, Dict[str, Any]]) -> None:
        """Save metadata including encoding information.

        Args:
            encodings: Dictionary mapping module names to their encoding information

        Raises:
            VBAError: If metadata cannot be saved
        """
        try:
            metadata = {
                "source_document": str(self.doc_path),
                "export_date": datetime.datetime.now().isoformat(),
                "encoding_mode": "fixed",
                "header_mode": "inline" if self.in_file_headers else ("separate" if self.save_headers else "none"),
                "encodings": encodings,
            }

            metadata_path = self.vba_dir / "vba_metadata.json"
            with open(metadata_path, "w", encoding="utf-8") as f:
                json.dump(metadata, f, indent=2)

            logger.info(f"Metadata saved to {metadata_path}")

        except Exception as e:
            error_msg = "Failed to save metadata"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def export_component(self, component: Any, directory: Path) -> None:
        """Export a single VBA component."""
        temp_file = None
        try:
            logger.debug(f"Starting component export for {component.Name}")
            info = self.component_handler.get_component_info(component)
            name = info["name"]
            logger.debug(f"Exporting component {name} with save_headers={self.save_headers}")
            temp_file = resolve_path(f"{name}.tmp", directory)

            # Export to temp file
            logger.debug("About to call component.Export")
            component.Export(str(temp_file))
            logger.debug("Component.Export completed")

            # Read and process content
            with open(temp_file, "r", encoding=self.encoding) as f:
                content = f.read()

            # Split content
            header, code = self.component_handler.split_vba_content(content)

            # Extract Rubberduck folder if enabled
            folder_path = ""
            if self.use_rubberduck_folders:
                folder_path, code = self.component_handler.get_rubberduck_folder(code)

            # Determine target directory
            target_directory = directory
            if folder_path:
                target_directory = directory / folder_path
                target_directory.mkdir(parents=True, exist_ok=True)
                logger.debug(f"Created folder structure: {target_directory}")

            # Write files
            self._write_component_files(name, header, code, info, target_directory)
            logger.debug(f"Component files written for {name} in {target_directory}")

            # Handle form binaries if this is a UserForm
            if info["type"] == VBATypes.VBEXT_CT_MSFORM:
                self._handle_form_binary_export(name)

            logger.info(f"Exported: {name}" + (f" (folder: {folder_path})" if folder_path else ""))

        except Exception as e:
            logger.error(f"Failed to export component {component.Name}: {str(e)}")
            raise VBAError(f"Failed to export component {component.Name}") from e
        finally:
            if temp_file and Path(temp_file).exists():
                try:
                    Path(temp_file).unlink()
                except OSError:
                    pass

    def import_component(self, file_path: Path, components: Any) -> None:
        """Import a VBA component with automatic header format detection.

        This method automatically detects whether headers are embedded in the code file
        or stored in separate .header files, eliminating the need for user flags.

        Detection logic:
        1. Check if file contains inline headers (VERSION/BEGIN/Attribute at start)
        2. If not, look for separate .header file
        3. If neither, create minimal headers as needed

        Args:
            file_path: Path to the code file
            components: VBA components collection

        Raises:
            VBAError: If component import fails
        """
        try:
            name = file_path.stem

            # Auto-detect header format
            has_inline_headers = self.component_handler.has_inline_headers(file_path, encoding=self.encoding)

            # Check for separate header file
            header_file = file_path.with_suffix(f"{file_path.suffix}.header")
            has_separate_header = header_file.exists()

            # Warn if both formats exist (conflicting headers)
            if has_inline_headers and has_separate_header:
                logger.warning(
                    f"Both inline headers and separate header file found for {file_path.name}. "
                    f"Using inline headers (separate .header file will be ignored)."
                )

            # Pass detected format to get_module_type
            module_type = self.component_handler.get_module_type(
                file_path, in_file_headers=has_inline_headers, encoding=self.encoding
            )

            logger.debug(f"Processing module: {name} (Type: {module_type}, Inline headers: {has_inline_headers})")

            # Route to appropriate import handler based on detection
            # Precedence: inline > separate > minimal
            if has_inline_headers:
                self._import_with_in_file_headers(file_path, components, module_type)
            else:
                # Existing logic for separate header files (or no headers)
                self._import_with_separate_headers(file_path, components, module_type)

            # Handle any form binaries if needed
            if module_type == VBAModuleType.FORM:
                self._handle_form_binary_import(name)

            logger.info(f"Successfully processed: {file_path.name}")

            # Only try to save for non-Access applications
            if self.app_name != "Access":
                self.save_document()

        except Exception as e:
            logger.error(f"Failed to handle {file_path.name}: {str(e)}")
            raise VBAError(f"Failed to handle {file_path.name}") from e

    def _import_with_in_file_headers(self, file_path: Path, components: Any, module_type: VBAModuleType) -> None:
        """Import VBA component when headers are embedded in files."""
        name = file_path.stem

        # Read the complete file content
        with open(file_path, "r", encoding=self.encoding) as f:
            full_content = f.read().strip()

        # Split into header and code
        header, code = self.component_handler.split_vba_content(full_content)

        # Add Rubberduck folder annotation if enabled
        if self.use_rubberduck_folders:
            folder_path = self.component_handler.get_folder_from_file_path(file_path, self.vba_dir)
            if folder_path:
                code = self.component_handler.add_rubberduck_folder(code, folder_path)

        # Handle based on module type
        if module_type == VBAModuleType.DOCUMENT:
            self._update_document_module(name, code, components)
            return

        try:
            # Try to get existing component
            component = components(name)

            # For UserForms and Class modules with headers, always use full import via temporary file
            if module_type == VBAModuleType.FORM or (module_type == VBAModuleType.CLASS and header):
                logger.debug(f"Using full import for {module_type.name.lower()} with headers: {name}")
                components.Remove(component)
                self._import_via_temp_file(name, full_content, components, file_path.suffix, original_file=file_path)
            else:
                # For standard modules or class modules without headers, just update content
                logger.debug(f"Updating existing component: {name}")
                self._update_module_content(component, code)

        except Exception:
            # Component doesn't exist, create new
            if module_type == VBAModuleType.FORM or (module_type == VBAModuleType.CLASS and header):
                logger.debug(f"Creating new {module_type.name.lower()} with headers via import: {name}")
                self._import_via_temp_file(name, full_content, components, file_path.suffix, original_file=file_path)
            else:
                logger.debug(f"Creating new component: {name}")
                self._create_new_component(name, code, module_type, components)

    def _import_via_temp_file(
        self, name: str, full_content: str, components: Any, file_extension: str = ".cls", original_file: Path = None
    ) -> None:
        """Import UserForm or Class module with headers using VBA's Import method.

        This method handles both UserForms and Class modules that have header attributes
        that need to be preserved during import.

        Args:
            name: Component name
            full_content: Complete file content including headers
            components: VBA components collection
            file_extension: Original file extension (.frm, .cls, or .bas) - crucial for VBA to detect type!
            original_file: Optional path to original file (to avoid temp file issues with UserForms)
        """
        # For UserForms with binary files (.frx), use the original file directly
        # The temp file approach causes issues because:
        # 1. Binary .frx files aren't copied
        # 2. VBA Import uses VB_Name attribute, not filename
        # 3. Temp files can cause naming conflicts
        if original_file and file_extension == ".frm":
            try:
                logger.debug(f"Importing UserForm directly from original file: {original_file}")
                components.Import(str(original_file))
                logger.info(f"Imported UserForm with embedded headers: {name}")
                return
            except Exception as e:
                logger.warning(f"Direct import failed for {name}, falling back to temp file: {e}")

        # Use temp file for class modules or as fallback
        temp_file = self.vba_dir / f"{name}_temp{file_extension}"
        try:
            # Write complete content to temp file
            with open(temp_file, "w", encoding=self.encoding) as f:
                f.write(full_content)

            # Import the complete module using VBA's built-in import
            # This preserves all header attributes including VB_PredeclaredId
            # NOTE: The file extension is critical - VBA uses it to determine component type!
            components.Import(str(temp_file))

            logger.info(f"Imported module with embedded headers via temporary file: {name}")

        finally:
            # Clean up temp file
            if temp_file.exists():
                temp_file.unlink()

    def _create_new_component(self, name: str, code: str, module_type: VBAModuleType, components: Any) -> None:
        """Create a new VBA component with just the code portion."""
        if module_type == VBAModuleType.CLASS:
            component = components.Add(VBATypes.VBEXT_CT_CLASSMODULE)
        elif module_type == VBAModuleType.FORM:
            component = components.Add(VBATypes.VBEXT_CT_MSFORM)
        else:
            component = components.Add(VBATypes.VBEXT_CT_STDMODULE)

        component.Name = name

        # Add only the code portion - headers are auto-generated
        # Filter out hidden member attributes (Issue #16)
        if code.strip():
            component.CodeModule.AddFromString(_filter_attributes(code))

    def _import_with_separate_headers(self, file_path: Path, components: Any, module_type: VBAModuleType) -> None:
        """Import VBA component using separate header files (existing logic)."""
        name = file_path.stem

        # Read code file
        code = self._read_code_file(file_path)

        # Add Rubberduck folder annotation if enabled
        if self.use_rubberduck_folders:
            folder_path = self.component_handler.get_folder_from_file_path(file_path, self.vba_dir)
            if folder_path:
                code = self.component_handler.add_rubberduck_folder(code, folder_path)
                logger.debug(f"Added @Folder annotation: {folder_path}")

        # Handle based on module type
        if module_type == VBAModuleType.DOCUMENT:
            logger.debug(f"Updating document module: {name}")
            self._update_document_module(name, code, components)
            return

        try:
            # Try to get existing component
            component = components(name)

            if self._should_force_import(module_type):
                # Remove and reimport if required
                logger.debug(f"Forcing full import for: {name}")
                components.Remove(component)
                self._import_new_module(name, code, module_type, components)
            else:
                # Update existing module content in-place
                logger.debug(f"Updating existing component: {name}")
                self._update_module_content(component, code)

        except Exception:
            # Component doesn't exist, create new
            logger.debug(f"Creating new module: {name}")
            # For new modules, we need header information
            header = self._read_header_file(file_path)
            if not header and module_type in [VBAModuleType.CLASS, VBAModuleType.FORM]:
                header = self.component_handler.create_minimal_header(name, module_type)
                logger.debug(f"Created minimal header for new module: {name}")

            # Prepare content for new module
            content = self.component_handler.prepare_import_content(name, module_type, header, code)
            self._import_new_module(name, content, module_type, components)

    def _should_force_import(self, module_type: VBAModuleType) -> bool:
        """Determine if a module type requires full import instead of content update.

        Override in app-specific handlers if needed.

        Args:
            module_type: Type of the VBA module

        Returns:
            bool: True if module should be removed and reimported
        """
        # Force import for forms (always) and class modules (when they have headers)
        return module_type in [VBAModuleType.FORM, VBAModuleType.CLASS]

    def _import_new_module(
        self, name: str, content: str, module_type: VBAModuleType, components: Any, in_file_headers: bool = True
    ) -> None:
        """Create and import a new module.

        Args:
            name: Name of the module
            content: Module content
            module_type: Type of the VBA module
            components: VBA components collection
            in_file_headers: Whether content includes embedded headers
        """
        # Create appropriate module type
        if module_type == VBAModuleType.CLASS:
            component = components.Add(VBATypes.VBEXT_CT_CLASSMODULE)
        elif module_type == VBAModuleType.FORM:
            component = components.Add(VBATypes.VBEXT_CT_MSFORM)
        else:  # Standard module
            component = components.Add(VBATypes.VBEXT_CT_STDMODULE)

        component.Name = name

        if in_file_headers:
            # Split the content into header and code
            header, code = self.component_handler.split_vba_content(content)

            # Set header attributes via COM properties (if any special handling needed)
            self._apply_header_attributes(component, header)

            # Add only the code portion
            # Filter out hidden member attributes (Issue #16)
            if code.strip():
                component.CodeModule.AddFromString(_filter_attributes(code))
        else:
            self._update_module_content(component, content)

    def _apply_header_attributes(self, component: Any, header: str) -> None:
        """Apply header attributes to a VBA component via COM properties.

        Args:
            component: VBA component to configure
            header: Header content with attributes
        """
        if not header:
            return

        # For most modules, the basic attributes are automatically handled
        # when you create the component and set its name

        # Special handling for specific attributes if needed
        lines = header.splitlines()
        for line in lines:
            line = line.strip()
            if line.startswith("Attribute VB_"):
                # Most VB_ attributes are read-only and set automatically
                # Only a few can be modified via COM
                if "VB_Description" in line:
                    # Extract and set description if supported
                    match = re.search(r'Attribute VB_Description = "([^"]*)"', line)
                    if match:
                        try:
                            component.Description = match.group(1)
                        except Exception:
                            pass  # Not all components support description

    def _update_module_content(self, component: Any, content: str) -> None:
        """Update the content of an existing module.

        When updating content directly in the VBA editor (without full import),
        we must not include header information as it can't be processed by
        the VBA project.

        Args:
            component: VBA component to update
            content: New content to set
        """
        try:
            # For direct updates, we want just the code without any header
            # manipulation - the existing module already has its header
            if component.CodeModule.CountOfLines > 0:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)

            # Filter out hidden member attributes (Issue #16)
            if content.strip():
                component.CodeModule.AddFromString(_filter_attributes(content))

            logger.debug(f"Updated content for: {component.Name}")
        except Exception as e:
            logger.error(f"Failed to update content for {component.Name}: {str(e)}")
            raise VBAError("Failed to update module content") from e

    def _handle_form_binary_export(self, name: str) -> None:
        """Handle form binary (.frx) export.

        Sets exported .frx files as read-only to prevent accidental modification,
        which would corrupt the UserForm and cause deletion from VBA project.

        On re-export, automatically removes read-only flag before overwriting.
        """
        try:
            frx_source = resolve_path(f"{name}.frx", Path(self.doc.FullName).parent)
            if frx_source.exists():
                frx_target = resolve_path(f"{name}.frx", self.vba_dir)
                try:
                    # If target exists and is read-only, make it writable before copying
                    if frx_target.exists():
                        import stat

                        current_mode = os.stat(str(frx_target)).st_mode
                        if not (current_mode & stat.S_IWRITE):
                            # File is read-only, make it writable
                            os.chmod(str(frx_target), current_mode | stat.S_IWRITE)
                            logger.debug(f"Temporarily removed read-only flag from {frx_target.name}")

                    shutil.copy2(str(frx_source), str(frx_target))

                    # Make the exported .frx file read-only to prevent accidental modification
                    import stat

                    current_mode = os.stat(str(frx_target)).st_mode
                    os.chmod(str(frx_target), current_mode & ~stat.S_IWRITE)
                    logger.debug(f"Exported form binary (read-only): {frx_target}")
                except (OSError, shutil.Error) as e:
                    logger.error(f"Failed to copy form binary {name}.frx: {e}")
                    raise VBAError(f"Failed to export form binary {name}.frx") from e
        except PathError as e:
            raise VBAError(f"Failed to handle form binary path: {str(e)}") from e

    def _handle_form_binary_import(self, name: str) -> None:
        """Handle form binary (.frx) import.

        Temporarily removes read-only flag if needed to allow copying.
        """
        try:
            frx_source = resolve_path(f"{name}.frx", self.vba_dir)
            if frx_source.exists():
                frx_target = resolve_path(f"{name}.frx", Path(self.doc.FullName).parent)
                try:
                    # If target exists and is read-only, make it writable temporarily
                    if frx_target.exists():
                        import stat

                        current_mode = os.stat(str(frx_target)).st_mode
                        if not (current_mode & stat.S_IWRITE):
                            # File is read-only, make it writable
                            os.chmod(str(frx_target), current_mode | stat.S_IWRITE)

                    shutil.copy2(str(frx_source), str(frx_target))
                    logger.debug(f"Imported form binary: {frx_target}")
                except (OSError, shutil.Error) as e:
                    logger.error(f"Failed to copy form binary {name}.frx: {e}")
                    raise VBAError(f"Failed to import form binary {name}.frx") from e
        except PathError as e:
            raise VBAError(f"Failed to handle form binary path: {str(e)}") from e

    @abstractmethod
    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update an existing document module."""
        pass

    def _read_header_file(self, code_file: Path) -> str:
        """Read the header file if it exists."""
        if self.in_file_headers:
            # Extract header from the code file itself
            try:
                with open(code_file, "r", encoding=self.encoding) as f:
                    content = f.read().strip()
                header, _ = self.component_handler.split_vba_content(content)
                return header
            except Exception as e:
                logger.debug(f"Could not read header from code file {code_file}: {e}")
                return ""
        else:
            # Use existing logic for separate header files
            header_file = code_file.with_suffix(".header")
            if header_file.exists():
                try:
                    with open(header_file, "r", encoding=self.encoding) as f:
                        return f.read().strip()
                except Exception as e:
                    logger.debug(f"Could not read header file {header_file}: {e}")
            return ""

    def _read_code_file(self, code_file: Path) -> str:
        """Read the code file."""
        try:
            with open(code_file, "r", encoding=self.encoding) as f:
                content = f.read().strip()

            if self.in_file_headers:
                # Split content to extract only the code part
                _, code = self.component_handler.split_vba_content(content)
                return code
            else:
                return content
        except Exception as e:
            logger.error(f"Failed to read code file {code_file}: {str(e)}")
            raise VBAError(f"Failed to read VBA code file: {code_file}") from e

    def _write_component_files(self, name: str, header: str, code: str, info: Dict[str, Any], directory: Path) -> None:
        """Write component files with proper encoding.

        Args:
            name: Name of the VBA component
            header: Header content (may be empty)
            code: Code content
            info: Component information dictionary
            directory: Target directory
        """
        if self.in_file_headers and header:
            # Combine header and code in single file
            combined_content = f"{header}\n{code}"
            code_file = directory / f"{name}{info['extension']}"
            with open(code_file, "w", encoding=self.encoding) as f:
                f.write(combined_content + "\n")
            logger.debug(f"Saved code file with embedded header: {code_file}")
        else:
            # Save header if enabled and header content exists
            if self.save_headers and header:
                header_file = directory / f"{name}.header"
                with open(header_file, "w", encoding=self.encoding) as f:
                    f.write(header + "\n")
                logger.debug(f"Saved header file: {header_file}")

            # Save code file
            code_file = directory / f"{name}{info['extension']}"
            with open(code_file, "w", encoding=self.encoding) as f:
                f.write(code + "\n")
            logger.debug(f"Saved code file: {code_file}")

    def watch_changes(self) -> None:
        """Watch for changes in VBA files and update the document."""
        try:
            logger.info(f"Watching for changes in {self.vba_dir}...")
            last_check_time = time.time()
            check_interval = 5  # Check connection every 5 seconds

            # Setup file patterns for watchfiles
            if self.use_rubberduck_folders:
                # Watch recursively
                watch_path = self.vba_dir
                recursive = True
            else:
                # Watch only the root directory
                watch_path = self.vba_dir
                recursive = False

            # Define VBA file extensions we want to watch
            vba_extensions = {".bas", ".cls", ".frm"}

            # Use yield_on_timeout=True so watch yields even without file changes
            # This allows us to check document state periodically
            for changes in watch(
                watch_path,
                recursive=recursive,
                rust_timeout=check_interval * 1000,  # Convert seconds to milliseconds
                yield_on_timeout=True,
            ):
                try:
                    # Check connection periodically (now triggered by timeout or changes)
                    current_time = time.time()
                    if current_time - last_check_time >= check_interval:
                        if not self.is_document_open():
                            raise DocumentClosedError(self.document_type)
                        last_check_time = current_time
                        logger.debug("Connection check passed")

                    # Filter changes to only include VBA files (exclude temp files)
                    vba_changes = []
                    for change_type, file_path in changes:
                        path = Path(file_path)
                        # Only include files with VBA extensions, but exclude temp files
                        # Temp files have pattern: *_temp.{bas,cls,frm}
                        if path.suffix.lower() in vba_extensions and not path.stem.endswith("_temp"):
                            vba_changes.append((change_type, file_path))

                    if vba_changes:
                        logger.debug(f"Watchfiles detected VBA changes: {vba_changes}")

                    for change_type, path in vba_changes:
                        try:
                            path = Path(path)
                            if change_type == Change.deleted:
                                # Handle deleted files
                                logger.info(f"Detected deletion of {path.name}")
                                if not self.is_document_open():
                                    raise DocumentClosedError(self.document_type)

                                vba_project = self.get_vba_project()
                                components = vba_project.VBComponents
                                try:
                                    component = components(path.stem)
                                    components.Remove(component)
                                    logger.info(f"Removed component: {path.stem}")
                                    self.doc.Save()
                                except Exception:
                                    logger.debug(f"Component {path.stem} already removed or not found")

                            elif change_type in (Change.added, Change.modified):
                                # Handle both added and modified files the same way
                                action = "addition" if change_type == Change.added else "modification"
                                logger.debug(f"Processing {action} in {path}")
                                self.import_single_file(path)

                        except (DocumentClosedError, RPCError) as e:
                            raise e
                        except Exception as e:
                            logger.warning(f"Error handling changes (will retry): {str(e)}")
                            continue

                except (DocumentClosedError, RPCError) as error:
                    raise error
                except Exception as error:
                    logger.warning(f"Error in watch loop (will continue): {str(error)}")

                # Prevent excessive CPU usage but stay responsive
                time.sleep(0.5)

        except KeyboardInterrupt:
            logger.info("\nStopping VBA editor...")
        except (DocumentClosedError, RPCError) as error:
            raise error
        finally:
            logger.info("VBA editor stopped.")

    def import_vba(self) -> None:
        """Import VBA content into the Office document."""
        try:
            # First check if document is accessible
            if self.doc is None:
                self.open_document()
            _ = self.doc.Name  # Check connection

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            # Find all VBA files, recursively if Rubberduck folders are enabled
            vba_files = []
            if self.use_rubberduck_folders:
                # Search recursively
                for ext in [".cls", ".bas", ".frm"]:
                    vba_files.extend(self.vba_dir.rglob(f"*{ext}"))
            else:
                # Search only in root directory
                for ext in [".cls", ".bas", ".frm"]:
                    vba_files.extend(self.vba_dir.glob(f"*{ext}"))

            if not vba_files:
                logger.info("No VBA files found to import.")
                return

            logger.info(f"\nFound {len(vba_files)} VBA files to import:")
            for vba_file in vba_files:
                relative_path = vba_file.relative_to(self.vba_dir)
                logger.info(f"  - {relative_path}")

            # Import components
            for vba_file in vba_files:
                try:
                    self.import_component(vba_file, components)
                except Exception as e:
                    logger.error(f"Failed to import {vba_file.name}: {str(e)}")
                    continue

            # Save if we successfully imported files
            self.save_document()

        except Exception as e:
            if check_rpc_error(e):
                raise DocumentClosedError(self.document_type)
            raise VBAError(str(e))

    def import_single_file(self, file_path: Path) -> None:
        """Import a single VBA file that has changed.

        Args:
            file_path: Path to the changed VBA file
        """
        logger.info(f"Processing changes in {file_path.name}")

        try:
            # Check if document is still open
            if not self.is_document_open():
                raise DocumentClosedError(self.document_type)

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            # Import the component
            self.import_component(file_path, components)

            # Only try to save for non-Access applications
            if self.app_name != "Access":
                self.doc.Save()

        except (DocumentClosedError, RPCError):
            raise
        except Exception as e:
            logger.error(f"Failed to process {file_path.name}: {str(e)}")
            raise VBAError(f"Failed to import {file_path.name}") from e

    def export_vba(
        self, save_metadata: bool = False, overwrite: bool = True, interactive: bool = True, keep_open: bool = False
    ) -> None:
        """Export VBA modules to files.

        Args:
            save_metadata: Whether to save metadata file
            overwrite: Whether to overwrite existing files
            interactive: Whether to prompt for confirmation on warnings (set False to skip prompts)
            keep_open: Whether to keep document open after export (default: False = close after export)

        Raises:
            VBAExportWarning: When user confirmation is needed (only if interactive=True)
        """
        logger.debug("Starting export_vba operation")
        try:
            # Ensure document is open
            if not self.is_document_open():
                logger.debug("Document not open, opening...")
                self.open_document()

            vba_project = self.get_vba_project()

            components = vba_project.VBComponents
            if not components.Count:
                logger.info(f"No VBA components found in the {self.document_type}.")
                return

            # Check if files already exist and raise warning if interactive
            if interactive and overwrite:
                existing_files = self._check_existing_vba_files()
                if existing_files:
                    raise VBAExportWarning(
                        "existing_files", {"file_count": len(existing_files), "files": existing_files}
                    )

            # Check if header mode has changed and raise warning if interactive
            if interactive:
                header_mode_changed = self._check_header_mode_change()
                if header_mode_changed:
                    old_mode, new_mode = self._get_header_modes()
                    raise VBAExportWarning("header_mode_changed", {"old_mode": old_mode, "new_mode": new_mode})

            # If we get here, either interactive=False or no warnings were triggered
            # Clean up old header files if header mode changed (for non-interactive retries)
            if self._check_header_mode_change():
                overwrite = True
                self._cleanup_old_header_files()

            # Track exported files for metadata
            encoding_data = {}

            for component in components:
                try:
                    info = self.component_handler.get_component_info(component)
                    base_name = info["name"]
                    final_file = resolve_path(f"{base_name}{info['extension']}", self.vba_dir)
                    header_file = resolve_path(f"{base_name}.header", self.vba_dir) if self.save_headers else None

                    # When using in_file_headers, always export to ensure headers are embedded
                    # When using save_headers, check both code and header files
                    should_export = overwrite

                    if not overwrite:
                        if self.in_file_headers:
                            # For in-file headers, only skip if the file exists
                            # (we can't tell if it has headers without reading it, so safer to re-export)
                            should_export = not final_file.exists()
                        else:
                            # For separate headers, check both code and header files
                            files_to_check = [final_file]
                            if header_file:
                                files_to_check.append(header_file)

                            # Export if any file is missing
                            should_export = any(not f.exists() for f in files_to_check)

                    if should_export:
                        self.export_component(component, self.vba_dir)
                        encoding_data[info["name"]] = {"encoding": self.encoding, "type": info["type_name"]}
                    else:
                        logger.debug(f"Skipping existing file: {final_file}")

                except Exception as e:
                    logger.error(f"Failed to export component {component.Name}: {str(e)}")
                    continue

            self._check_form_safety(self.vba_dir)  # Check for forms before proceeding

            # Check if we have any UserForms (in exported data or existing files)
            has_forms = any(info.get("type") == "UserForm" for info in encoding_data.values())
            if not has_forms:
                # Also check for existing .frm files in case they were skipped
                has_forms = bool(list(self.vba_dir.glob("*.frm")))

            # Warn about .frx files if forms were exported
            if has_forms:
                logger.info("NOTE: UserForm binary files (.frx) are exported as read-only to prevent corruption.")
                logger.info("      Modifying .frx files will corrupt the UserForm and cause deletion from VBA project.")
                logger.info("      Only edit .frm files - .frx files are handled automatically.")

            # Save metadata if requested, or if we have forms (to track header mode)
            if save_metadata or has_forms:
                logger.debug("Saving metadata...")
                self._save_metadata(encoding_data)
                logger.debug("Metadata saved")
            else:
                logger.debug("Skipping metadata save (no forms and not requested)")

            # Show exported files to user if requested

            # Plattform independent way to open the directory commented out
            # as only Windows is supported for now

            # try:
            if self.open_folder:
                logger.debug("Opening export directory...")
                os.startfile(str(self.vba_dir))
            else:
                logger.info(f"VBA modules exported to: {self.vba_dir}")
            # except AttributeError:
            #     # os.startfile is Windows only, use platform-specific alternatives
            #     if sys.platform == "darwin":
            #         subprocess.run(["open", str(self.vba_dir)])
            #     else:
            #         subprocess.run(["xdg-open", str(self.vba_dir)])

            # Close document after export unless --keep-open flag is set
            if not keep_open:
                logger.debug("Closing document after export (use --keep-open to override)")
                self.close_document()

        except VBAExportWarning:
            # Let warnings propagate to CLI layer for user interaction
            raise
        except Exception as e:
            error_msg = "Failed to export VBA content"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e


class WordVBAHandler(OfficeVBAHandler):
    """Microsoft Word specific implementation of VBA operations.

    Provides Word-specific implementations of abstract methods from OfficeVBAHandler
    and any additional functionality specific to Word VBA projects.

    The handler manages operations like:
    - Importing/exporting VBA modules
    - Handling UserForm binaries (.frx files)
    - Managing ThisDocument module
    - Monitoring file changes
    """

    @property
    def app_name(self) -> str:
        """Name of the Office application."""
        return "Word"

    @property
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        return "Word.Application"

    def get_document_module_name(self) -> str:
        """Get the name of the document module."""
        return "ThisDocument"

    def _open_document_impl(self) -> Any:
        """Implementation-specific document opening logic."""
        return self.app.Documents.Open(str(self.doc_path))

    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update an existing document module for Word."""
        try:
            doc_component = components(name)

            # Clear existing code
            if doc_component.CodeModule.CountOfLines > 0:
                doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)

            # Add new code (filter hidden member attributes - Issue #16)
            if code.strip():
                doc_component.CodeModule.AddFromString(_filter_attributes(code))

            logger.info(f"Updated document module: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update document module {name}") from e


class ExcelVBAHandler(OfficeVBAHandler):
    """Microsoft Excel specific implementation of VBA operations.

    Provides Excel-specific implementations of abstract methods from OfficeVBAHandler
    and any additional functionality specific to Excel VBA projects.

    The handler manages operations like:
    - Importing/exporting VBA modules
    - Handling UserForm binaries (.frx files)
    - Managing ThisWorkbook and Sheet modules
    - Monitoring file changes
    """

    @property
    def app_name(self) -> str:
        """Name of the Office application."""
        return "Excel"

    @property
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        return "Excel.Application"

    def get_document_module_name(self) -> str:
        """Get the name of the document module."""
        return "ThisWorkbook"

    def _open_document_impl(self) -> Any:
        """Implementation-specific document opening logic."""
        return self.app.Workbooks.Open(str(self.doc_path))

    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update an existing document module for Excel."""
        try:
            # Handle ThisWorkbook and Sheet modules
            doc_component = components(name)

            # Clear existing code
            if doc_component.CodeModule.CountOfLines > 0:
                doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)

            # Add new code (filter hidden member attributes - Issue #16)
            if code.strip():
                doc_component.CodeModule.AddFromString(_filter_attributes(code))

            logger.info(f"Updated document module: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update document module {name}") from e


class AccessVBAHandler(OfficeVBAHandler):
    """Microsoft Access specific implementation of VBA operations.

    Handles Access-specific implementations for VBA module management, with special
    consideration for Access's unique behaviors around database and VBA project handling.

    Access differs from Word/Excel in several ways:
    - Uses VBE.ActiveVBProject instead of doc.VBProject
    - No document module equivalent (like ThisDocument/ThisWorkbook)
    - Different handling of database connections and saving
    - Forms handled differently (not supported in VBA editing)
    """

    def __init__(
        self,
        doc_path: str,
        vba_dir: Optional[str] = None,
        encoding: str = "cp1252",
        verbose: bool = False,
        save_headers: bool = False,
        use_rubberduck_folders: bool = False,
        open_folder: bool = False,
        in_file_headers: bool = False,
    ):
        """Initialize the Access VBA handler.

        Args:
            doc_path: Path to the Access database
            vba_dir: Directory for VBA files (defaults to current directory)
            encoding: Character encoding for VBA files (default: cp1252)
            verbose: Enable verbose logging
            save_headers: Whether to save VBA component headers to separate files
            use_rubberduck_folders: Whether to use Rubberduck folder structure
            open_folder: Whether to open the VBA directory after export
            in_file_headers: Whether to include headers directly in code files
        """
        try:
            # Let parent handle path resolution
            super().__init__(
                doc_path=doc_path,
                vba_dir=vba_dir,
                encoding=encoding,
                verbose=verbose,
                save_headers=save_headers,
                use_rubberduck_folders=use_rubberduck_folders,
                open_folder=open_folder,
                in_file_headers=in_file_headers,
            )

            # Handle Access-specific initialization
            try:
                # Try to get running instance first
                app = win32com.client.GetObject("Access.Application")
                try:
                    current_db = app.CurrentDb()
                    if current_db and str(self.doc_path) == current_db.Name:
                        logger.debug("Using already open database")
                        self.app = app
                        self.doc = current_db
                        return
                except Exception:
                    pass
            except Exception:
                pass

            # If we get here, we need to initialize a new instance
            logger.debug("No existing database connection found, initializing new instance")
            self.initialize_app()
            self.doc = self._open_document_impl()
            logger.debug("Database opened successfully")

        except VBAError:
            raise
        except Exception as e:
            raise VBAError(f"Failed to initialize Access VBA handler: {str(e)}") from e

    @property
    def app_name(self) -> str:
        """Name of the Office application."""
        return "Access"

    @property
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        return "Access.Application"

    @property
    def document_type(self) -> str:
        """Document type string for error messages."""
        return "database"

    def _open_document_impl(self) -> Any:
        """Open database in Access with proper error handling.

        Returns:
            The current database object

        Raises:
            RPCError: If connection to Access is lost
            VBAError: For other database access errors
        """
        try:
            # Check if database is already open
            try:
                current_db = self.app.CurrentDb()
                if current_db and str(self.doc_path) == current_db.Name:
                    logger.debug("Using already open database")
                    return current_db
            except Exception:
                pass  # Handle case where no database is open

            logger.debug(f"Opening database: {self.doc_path}")
            self.app.OpenCurrentDatabase(str(self.doc_path))
            return self.app.CurrentDb()

        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(
                    "\nLost connection to Access. The operation will be terminated.\n"
                    "This typically happens if Access was closed via the UI.\n"
                    "To continue:\n"
                    "1. Start Access\n"
                    "2. Run the access-vba command again"
                )
            raise VBAError(f"Failed to open database: {str(e)}") from e

    def get_document_module_name(self) -> str:
        """Get the name of the document module.

        Access databases don't have an equivalent to Word's ThisDocument
        or Excel's ThisWorkbook, so this returns an empty string.
        """
        return ""

    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update module code in Access.

        Access doesn't have document modules like Word/Excel, but we still need this
        method for the interface. For Access, we'll use it to update any module's content.

        Args:
            name: Name of the module
            code: New code to insert
            components: VBA components collection

        Raises:
            VBAError: If module update fails
        """
        try:
            component = components(name)

            # Clear existing code
            if component.CodeModule.CountOfLines > 0:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)

            # Add new code if not empty (filter hidden member attributes - Issue #16)
            if code.strip():
                component.CodeModule.AddFromString(_filter_attributes(code))

            logger.info(f"Updated module content: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update module {name}") from e

    def save_document(self) -> None:
        """Handle saving in Access.

        Access VBA projects save automatically when modules are modified.
        We only verify the database is still accessible and log appropriately.
        """
        try:
            if self.doc is not None:
                # Just verify database is still open/accessible
                _ = self.app.CurrentDb()
                logger.debug("Database verified accessible - Access auto-saves changes")
        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            # Don't raise other errors - Access handles saving automatically

    def is_document_open(self) -> bool:
        """Check if the database is still open and accessible.

        Returns:
            bool: True if database is open and accessible

        Raises:
            RPCError: If connection to Access is lost
            DocumentClosedError: If database is closed
        """
        try:
            if self.doc is None:
                return False

            current_db = self.app.CurrentDb()
            return current_db and str(self.doc_path) == current_db.Name

        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            raise DocumentClosedError(self.document_type)


class PowerPointVBAHandler(OfficeVBAHandler):
    """Microsoft PowerPoint specific implementation of VBA operations.

    PowerPoint has a unique VBA project structure:
    - No document-level module (unlike Word's ThisDocument or Excel's ThisWorkbook)
    - Each slide has its own module (e.g., "Slide1", "Slide2")
    - Standard modules, class modules, and forms work the same as other Office apps
    """

    @property
    def app_name(self) -> str:
        """Name of the Office application."""
        return "PowerPoint"

    @property
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        return "PowerPoint.Application"

    def get_document_module_name(self) -> str:
        """Get the name of the presentation module.

        PowerPoint has no document-level module, so return empty string.
        """
        return ""

    def document_is_read_only(self) -> bool:
        """Check if the PowerPoint presentation is read-only.

        PowerPoint doesn't have a direct read-only flag, so we check if the
        presentation is protected for editing.

        Returns:
            bool: True if the presentation is read-only
        """
        try:
            # Check if presentation is read-only
            if self.doc.ReadOnly:
                logger.warning("\nPresentation is opened in read-only mode!")
            return True
        except Exception as e:
            raise VBAError("Failed to check if presentation is read-only") from e

    def _open_document_impl(self) -> Any:
        """Implementation-specific presentation opening logic."""
        try:
            # Check if presentation is already open in app
            for pres in self.app.Presentations:
                if str(self.doc_path) == pres.FullName:
                    logger.debug("Using already open presentation")
                    return pres

            # If not found, open it
            logger.debug(f"Opening presentation: {self.doc_path}")
            return self.app.Presentations.Open(str(self.doc_path))

        except Exception as e:
            raise VBAError(f"Failed to open presentation: {str(e)}") from e

    def save_document(self) -> None:
        """Save the presentation if it's open.

        PowerPoint requires specific save handling different from Word/Excel.
        """
        if self.doc is not None:
            try:
                # PowerPoint uses Save() not SaveAs()
                self.doc.Save()
                logger.info("Presentation has been saved and left open for further editing")
            except Exception as e:
                if check_rpc_error(e):
                    raise RPCError(self.app_name) from e
                raise VBAError("Failed to save presentation") from e

    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update module code.

        Since PowerPoint has no document module, this acts like a regular module update.
        """
        try:
            component = components(name)

            # Clear existing code
            if component.CodeModule.CountOfLines > 0:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)

            # Add new code (filter hidden member attributes - Issue #16)
            if code.strip():
                component.CodeModule.AddFromString(_filter_attributes(code))

            logger.info(f"Updated module: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update module {name}") from e

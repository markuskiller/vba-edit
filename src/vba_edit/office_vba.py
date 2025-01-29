from abc import ABC, abstractmethod
import datetime
import json
import logging
import os
import re
import shutil
import sys
import time
from enum import Enum, auto
from pathlib import Path
from typing import Dict, Optional, Any, Tuple

# Third-party imports
import win32com.client
from watchgod import Change, RegExpWatcher

# Updated local imports
from vba_edit.path_utils import (
    resolve_path,
    get_document_paths,
)

from vba_edit.utils import (
    is_vba_access_error,
    get_vba_error_details,
)

from vba_edit.exceptions import (
    VBAError,
    VBAAccessError,
    DocumentClosedError,
    DocumentNotFoundError,
    RPCError,
    check_rpc_error,
    PathError,
)

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

# Command-line entry points for different Office applications
OFFICE_CLI_NAMES = {app: f"{app}-vba" for app in OFFICE_MACRO_EXTENSIONS.keys()}

# Currently supported apps in vba-edit
# "access" is only partially supported at this stage and will be included
# in list as soon as tests are adapted to handle it
SUPPORTED_APPS = [
    "word",
    "excel",
    # "access",
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
    def is_document_module(cls, name: str, app_type: str) -> bool:
        """Check if a name matches document module patterns for specific Office app.
        
        Args:
            name: Name to check
            app_type: Office application type ('word', 'excel', 'powerpoint', 'access')
            
        Returns:
            bool: True if name matches document module pattern for the app
            
        Raises:
            ValueError: If app_type is not supported
        """
        app_type = app_type.lower()
        
        if app_type not in {'word', 'excel', 'powerpoint', 'access'}:
            raise ValueError(f"Unsupported application type: {app_type}")
            
        # Access has no document modules
        if app_type == 'access':
            return False
            
        # Word document check
        if app_type == 'word':
            return name in cls.WORD_DOCUMENT_NAMES
            
        # Excel specific checks
        if app_type == 'excel':
            if name in cls.EXCEL_WORKBOOK_NAMES:
                return True
            return any(
                name.startswith(prefix) and name[len(prefix):].isdigit()
                for prefix in cls.EXCEL_SHEET_PREFIXES
            )
            
        # PowerPoint specific checks
        if app_type == 'powerpoint':
            return any(
                name.startswith(prefix) and name[len(prefix):].isdigit()
                for prefix in cls.POWERPOINT_SLIDE_PREFIXES
            )


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
    def __init__(self, app_name: str):
        """Initialize component handler.
        
        Args:
            app_name: Name of Office application ('word', 'excel', 'powerpoint', 'access')
        """
        self.app_name = app_name.lower()
        if self.app_name not in {'word', 'excel', 'powerpoint', 'access'}:
            raise ValueError(f"Unsupported application type: {app_name}")

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

    def get_module_type(self, file_path: Path) -> VBAModuleType:
        """Determine VBA module type from file extension and content.

        Args:
            file_path: Path to the VBA module file

        Returns:
            Appropriate VBAModuleType

        Raises:
            ValueError: If file extension is unknown
        """
        suffix = file_path.suffix.lower()
        name = file_path.stem

        # Check if it's a known document module name for this app type
        if VBADocumentNames.is_document_module(name, self.app_name):
            return VBAModuleType.DOCUMENT

        if suffix == ".bas":
            return VBAModuleType.STANDARD
        elif suffix == ".frm":
            return VBAModuleType.FORM
        elif suffix == ".cls":
            # For .cls files, check the header if available
            header_file = file_path.with_suffix(".header")
            if header_file.exists():
                with open(header_file, "r", encoding="utf-8") as f:
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
            Only module-level attributes (VB_Name, VB_GlobalNameSpace, VB_Creatable,
            VB_PredeclaredId, VB_Exposed) are considered part of the header.
            Procedure-level attributes are considered part of the code.
        """
        if not content.strip():
            return "", ""

        lines = content.splitlines()
        last_attr_idx = -1

        for i, line in enumerate(lines):
            stripped = line.strip()
            if stripped.startswith("Attribute VB_"):
                last_attr_idx = i
            elif last_attr_idx >= 0 and not stripped.startswith("Attribute VB_"):
                break

        if last_attr_idx == -1:
            return "", content

        header = "\n".join(lines[: last_attr_idx + 1])
        code = "\n".join(lines[last_attr_idx + 1 :])

        return header.strip(), code.strip()

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

    def prepare_import_content(self, name: str, module_type: VBAModuleType, header: str, code: str) -> str:
        """Prepare content for VBA component import.

        Args:
            name: Name of the VBA component
            module_type: Type of the VBA module
            header: Header content (may be empty)
            code: Code content

        Returns:
            Properly formatted content for import
        """
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

            if content.strip():
                component.CodeModule.AddFromString(content)

            logger.debug(f"Updated content for: {component.Name}")
        except Exception as e:
            logger.error(f"Failed to update content for {component.Name}: {str(e)}")
            raise VBAError("Failed to update module content") from e


class OfficeVBAHandler(ABC):
    """Abstract base class for handling VBA operations across Office applications.

    This class provides the foundation for application-specific VBA handlers,
    implementing common functionality while requiring specific implementations
    for application-dependent operations.

    Args:
        doc_path (str): Path to the Office document
        vba_dir (Optional[str]): Directory for VBA files (defaults to current directory)
        encoding (str): Character encoding for VBA files (default: cp1252)
        verbose (bool): Enable verbose logging
        save_headers (bool): Whether to save VBA component headers to separate files

    Attributes:
        doc_path (Path): Resolved path to the Office document
        vba_dir (Path): Resolved path to VBA directory
        encoding (str): Character encoding for file operations
        verbose (bool): Verbose logging flag
        save_headers (bool): Header saving flag
        app: Office application COM object
        doc: Office document COM object
        component_handler (VBAComponentHandler): Utility handler for VBA components
    """

    def __init__(
        self,
        doc_path: str,
        vba_dir: Optional[str] = None,
        encoding: str = "cp1252",
        verbose: bool = False,
        save_headers: bool = False,
    ):
        """Initialize the VBA handler."""
        try:
            # Let DocumentNotFoundError propagate as is - it's more fundamental than VBA errors
            self.doc_path, self.vba_dir = get_document_paths(doc_path, None, vba_dir)
            self.encoding = encoding
            self.verbose = verbose
            self.save_headers = save_headers
            self.app = None
            self.doc = None
            self.component_handler = VBAComponentHandler(app_name=self.app_name)

            # Configure logging
            log_level = logging.DEBUG if verbose else logging.INFO
            logger.setLevel(log_level)

            logger.debug(f"Initialized {self.__class__.__name__} with document: {doc_path}")
            logger.debug(f"VBA directory: {self.vba_dir}")
            logger.debug(f"Using encoding: {encoding}")
            logger.debug(f"Save headers: {save_headers}")

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
            VBAError: If .frm files are found and save_headers is False
        """
        if not self.save_headers:
            form_files = list(vba_dir.glob("*.frm"))
            if form_files:
                form_names = ", ".join(f.stem for f in form_files)
                error_msg = (
                    f"\nERROR: Found UserForm files ({form_names}) but --save-headers is not enabled!\n"
                    f"UserForms require their full header information to maintain form specifications.\n"
                    f"Please re-run the command with the --save-headers flag to preserve form settings."
                )
                logger.error(error_msg)
                sys.exit(1)

    def open_document(self) -> None:
        """Open the Office document."""
        try:
            if self.doc is None:
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

            # Write files
            self._write_component_files(name, header, code, info, directory)
            logger.debug(f"Component files written for {name}")

            logger.info(f"Exported: {name}")

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
        """Import a VBA component with app-specific handling.

        This method handles both new module creation and updates to existing modules.
        For updates, it will prefer in-place content updates where possible, only doing
        full imports when required by specific applications or module types.

        Args:
            file_path: Path to the code file
            components: VBA components collection

        Raises:
            VBAError: If component import fails
        """
        try:
            name = file_path.stem
            module_type = self.component_handler.get_module_type(file_path)

            logger.debug(f"Processing module: {name} (Type: {module_type})")

            # Read code file
            code = self._read_code_file(file_path)

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

    def _should_force_import(self, module_type: VBAModuleType) -> bool:
        """Determine if a module type requires full import instead of content update.

        Override in app-specific handlers if needed.

        Args:
            module_type: Type of the VBA module

        Returns:
            bool: True if module should be removed and reimported
        """
        # By default, only force import for forms
        return module_type == VBAModuleType.FORM

    def _import_new_module(self, name: str, content: str, module_type: VBAModuleType, components: Any) -> None:
        """Create and import a new module.

        Args:
            name: Name of the module
            content: Module content
            module_type: Type of the VBA module
            components: VBA components collection
        """
        # Create appropriate module type
        if module_type == VBAModuleType.CLASS:
            component = components.Add(VBATypes.VBEXT_CT_CLASSMODULE)
        elif module_type == VBAModuleType.FORM:
            component = components.Add(VBATypes.VBEXT_CT_MSFORM)
        else:  # Standard module
            component = components.Add(VBATypes.VBEXT_CT_STDMODULE)

        component.Name = name
        self._update_module_content(component, content)

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

            if content.strip():
                component.CodeModule.AddFromString(content)

            logger.debug(f"Updated content for: {component.Name}")
        except Exception as e:
            logger.error(f"Failed to update content for {component.Name}: {str(e)}")
            raise VBAError("Failed to update module content") from e

    def _handle_form_binary_export(self, name: str) -> None:
        """Handle form binary (.frx) export."""
        try:
            frx_source = resolve_path(f"{name}.frx", Path(self.doc.FullName).parent)
            if frx_source.exists():
                frx_target = resolve_path(f"{name}.frx", self.vba_dir)
                try:
                    shutil.copy2(str(frx_source), str(frx_target))
                    logger.debug(f"Exported form binary: {frx_target}")
                except (OSError, shutil.Error) as e:
                    logger.error(f"Failed to copy form binary {name}.frx: {e}")
                    raise VBAError(f"Failed to export form binary {name}.frx") from e
        except PathError as e:
            raise VBAError(f"Failed to handle form binary path: {str(e)}") from e

    def _handle_form_binary_import(self, name: str) -> None:
        """Handle form binary (.frx) import."""
        try:
            frx_source = resolve_path(f"{name}.frx", self.vba_dir)
            if frx_source.exists():
                frx_target = resolve_path(f"{name}.frx", Path(self.doc.FullName).parent)
                try:
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
        header_file = code_file.with_suffix(".header")
        if header_file.exists():
            with open(header_file, "r", encoding="utf-8") as f:
                return f.read().strip()
        return ""

    def _read_code_file(self, code_file: Path) -> str:
        """Read the code file."""
        with open(code_file, "r", encoding="utf-8") as f:
            return f.read().strip()

    def _write_component_files(self, name: str, header: str, code: str, info: Dict[str, Any], directory: Path) -> None:
        """Write component files with proper encoding.

        Args:
            name: Name of the VBA component
            header: Header content (may be empty)
            code: Code content
            info: Component information dictionary
            directory: Target directory
        """
        # Save header if enabled and header content exists
        if self.save_headers and header:
            header_file = directory / f"{name}.header"
            with open(header_file, "w", encoding="utf-8") as f:
                f.write(header + "\n")
            logger.debug(f"Saved header file: {header_file}")

        # Always save code file
        code_file = directory / f"{name}{info['extension']}"
        with open(code_file, "w", encoding="utf-8") as f:
            f.write(code + "\n")
        logger.debug(f"Saved code file: {code_file}")

    def watch_changes(self) -> None:
        """Watch for changes in VBA files and update the document."""
        try:
            logger.info(f"Watching for changes in {self.vba_dir}...")
            last_check_time = time.time()
            check_interval = 30  # Check connection every 30 seconds

            # Setup file watcher
            watcher = RegExpWatcher(
                self.vba_dir,
                re_files=r"^.*\.(cls|frm|bas)$",
            )

            while True:
                try:
                    # Check connection periodically
                    current_time = time.time()
                    if current_time - last_check_time >= check_interval:
                        if not self.is_document_open():
                            raise DocumentClosedError(self.document_type)
                        last_check_time = current_time
                        logger.debug("Connection check passed")

                    # Check for changes using watchgod
                    changes = watcher.check()
                    if changes:
                        logger.debug(f"Watchgod detected changes: {changes}")

                    for change_type, path in changes:
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

            # Find all VBA files
            vba_files = []
            for ext in [".cls", ".bas", ".frm"]:
                vba_files.extend(self.vba_dir.glob(f"*{ext}"))

            if not vba_files:
                logger.info("No VBA files found to import.")
                return

            logger.info(f"\nFound {len(vba_files)} VBA files to import:")
            for vba_file in vba_files:
                logger.info(f"  - {vba_file.name}")

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

    def export_vba(self, save_metadata: bool = False, overwrite: bool = True) -> None:
        """Export VBA modules to files."""
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

            # Track exported files for metadata
            encoding_data = {}

            for component in components:
                try:
                    info = self.component_handler.get_component_info(component)
                    base_name = info["name"]
                    final_file = resolve_path(f"{base_name}{info['extension']}", self.vba_dir)
                    header_file = resolve_path(f"{base_name}.header", self.vba_dir) if self.save_headers else None

                    # Handle both code and header files
                    files_to_check = [(final_file, False)]
                    if header_file:
                        files_to_check.append((header_file, True))

                    # Check each file
                    should_export = False
                    for file_path, is_header in files_to_check:
                        if overwrite or not file_path.exists():
                            should_export = True
                            break

                    if should_export:
                        self.export_component(component, self.vba_dir)
                        encoding_data[info["name"]] = {"encoding": self.encoding, "type": info["type_name"]}
                    else:
                        logger.debug(f"Skipping existing file: {final_file}")

                except Exception as e:
                    logger.error(f"Failed to export component {component.Name}: {str(e)}")
                    continue

            self._check_form_safety(self.vba_dir)  # Check for forms before proceeding

            # Save metadata if requested
            if save_metadata:
                logger.debug("Saving metadata...")
                self._save_metadata(encoding_data)
                logger.debug("Metadata saved")

            # Show exported files to user

            # Plattform independent way to open the directory commented out
            # as only Windows is supported for now

            # try:
            os.startfile(str(self.vba_dir))
            # except AttributeError:
            #     # os.startfile is Windows only, use platform-specific alternatives
            #     if sys.platform == "darwin":
            #         subprocess.run(["open", str(self.vba_dir)])
            #     else:
            #         subprocess.run(["xdg-open", str(self.vba_dir)])

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

            # Add new code
            if code.strip():
                doc_component.CodeModule.AddFromString(code)

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

            # Add new code
            if code.strip():
                doc_component.CodeModule.AddFromString(code)

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
    ):
        """Initialize the Access VBA handler.

        Args:
            doc_path: Path to the Access database
            vba_dir: Directory for VBA files (defaults to current directory)
            encoding: Character encoding for VBA files (default: cp1252)
            verbose: Enable verbose logging
            save_headers: Whether to save VBA component headers to separate files
        """
        try:
            # Let parent handle path resolution
            super().__init__(
                doc_path=doc_path,
                vba_dir=vba_dir,
                encoding=encoding,
                verbose=verbose,
                save_headers=save_headers,
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

            # Add new code if not empty
            if code.strip():
                component.CodeModule.AddFromString(code)

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

            # Add new code
            if code.strip():
                component.CodeModule.AddFromString(code)

            logger.info(f"Updated module: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update module {name}") from e

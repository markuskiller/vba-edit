"""
VBA Reference Management Module

This module provides functionality for managing VBA library references in Office documents.
It supports listing, adding, removing, and exporting/importing references via TOML configuration.

**Uses pure win32com** - xlwings is optional for convenience but NOT required.

Phase 1: Core reference management for single documents (v0.5.0)
Phase 2: CLI integration (v0.5.0)
Phase 3: Enhanced features and polish (v0.5.0)

Usage with win32com (recommended):
    import win32com.client
    from vba_edit.reference_manager import ReferenceManager
    
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open("file.xlsm")
    manager = ReferenceManager(wb)
    
Usage with xlwings (optional convenience):
    import xlwings as xw
    from vba_edit.reference_manager import ReferenceManager
    
    wb = xw.books.open("file.xlsm")
    manager = ReferenceManager(wb)  # Automatically extracts COM object
"""

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pywintypes
import win32com.client

from vba_edit.exceptions import (
    VBAAccessError,
    VBAError,
)

# Configure module logger
logger = logging.getLogger(__name__)


class ReferenceError(VBAError):
    """Exception raised for reference-related errors."""
    pass


class ReferenceManager:
    """Manage VBA library references in Office documents.
    
    This class provides methods to list, add, remove, and manage VBA library references
    in Microsoft Office documents using pure win32com COM automation.
    
    **xlwings is optional** - The manager accepts both xlwings workbook objects (for convenience)
    and win32com COM objects directly. When given an xlwings object, it automatically extracts
    the underlying COM object and uses pure win32com for all operations.
    
    Attributes:
        document: The Office document object (xlwings or win32com)
        vb_project: The VBA project COM object for the document
        
    Example with win32com (recommended):
        >>> import win32com.client
        >>> excel = win32com.client.Dispatch("Excel.Application")
        >>> wb = excel.Workbooks.Open("file.xlsm")
        >>> manager = ReferenceManager(wb)
        >>> refs = manager.list_references()
        
    Example with xlwings (optional):
        >>> import xlwings as xw
        >>> wb = xw.books.open("file.xlsm")
        >>> manager = ReferenceManager(wb)  # Extracts wb.api automatically
        >>> refs = manager.list_references()
    """
    
    # GUID validation pattern: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    GUID_PATTERN = re.compile(
        r'^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$'
    )
    
    def __init__(self, document: Any):
        """Initialize the ReferenceManager.
        
        Accepts both win32com COM objects and xlwings objects. When given an xlwings
        object (has .api attribute), automatically extracts the underlying COM object.
        All subsequent operations use pure win32com.
        
        Args:
            document: Office document object (win32com COM object or xlwings workbook)
            
        Raises:
            VBAAccessError: If VBA project access is denied
            ReferenceError: If document doesn't support VBA references
        """
        self.document = document
        self.vb_project = None
        
        logger.debug("Initializing ReferenceManager")
        
        try:
            # Handle xlwings objects (have .api attribute)
            if hasattr(document, 'api'):
                logger.debug("Detected xlwings document object")
                self.vb_project = document.api.VBProject
            # Handle win32com objects directly
            elif hasattr(document, 'VBProject'):
                logger.debug("Detected win32com document object")
                self.vb_project = document.VBProject
            else:
                raise ReferenceError(
                    "Document object does not support VBA projects. "
                    "Expected xlwings workbook or win32com Office document."
                )
            
            # Test access to VBA project
            try:
                _ = self.vb_project.Name
                logger.debug(f"Successfully accessed VBA project: {self.vb_project.Name}")
            except pywintypes.com_error as e:
                error_code = e.args[0] if e.args else None
                if error_code == -2147352567:  # Access denied
                    logger.error("VBA project access denied - trust access not enabled")
                    raise VBAAccessError(
                        'Access to VBA project denied. Please enable '
                        '"Trust access to the VBA project object model" in Trust Center settings.'
                    )
                raise
                
        except pywintypes.com_error as e:
            logger.error(f"COM error initializing ReferenceManager: {e}")
            raise ReferenceError(f"Failed to access VBA project: {e}")
    
    def list_references(self) -> List[Dict[str, Any]]:
        """List all VBA references in the document.
        
        Returns a list of dictionaries containing reference information:
        - name: Reference name (e.g., "Scripting", "Excel")
        - guid: Reference GUID (e.g., "{420B2830-E718-11CF-893D-00A0C9054228}")
        - major: Major version number
        - minor: Minor version number
        - priority: Loading priority (lower = earlier)
        - builtin: Whether it's a built-in reference
        - broken: Whether the reference is broken/missing
        - description: Reference description (optional)
        - path: Reference file path (if available)
        
        Returns:
            List of reference dictionaries
            
        Raises:
            ReferenceError: If unable to access references
            
        Example:
            >>> refs = manager.list_references()
            >>> print(f"Found {len(refs)} references")
            >>> for ref in refs:
            ...     status = "BROKEN" if ref['broken'] else "OK"
            ...     print(f"{status}: {ref['name']} v{ref['major']}.{ref['minor']}")
        """
        logger.debug("Listing VBA references")
        references = []
        
        try:
            refs_collection = self.vb_project.References
            ref_count = refs_collection.Count
            logger.debug(f"Found {ref_count} references in VBA project")
            
            # VBA collections are 1-indexed
            for i in range(1, ref_count + 1):
                ref = refs_collection.Item(i)
                
                try:
                    ref_info = {
                        'name': ref.Name,
                        'guid': ref.Guid,
                        'major': ref.Major,
                        'minor': ref.Minor,
                        'priority': i,  # Position in collection = loading priority
                        'builtin': ref.BuiltIn,
                        'broken': ref.IsBroken,
                    }
                    
                    # Optional fields that might not always be available
                    try:
                        ref_info['description'] = ref.Description
                    except (AttributeError, pywintypes.com_error):
                        ref_info['description'] = ""
                    
                    try:
                        ref_info['path'] = ref.FullPath
                    except (AttributeError, pywintypes.com_error):
                        ref_info['path'] = ""
                    
                    references.append(ref_info)
                    
                    status = "⚠ BROKEN" if ref.IsBroken else "✓"
                    builtin_str = " [BUILTIN]" if ref.BuiltIn else ""
                    logger.debug(
                        f"{status} Reference {i}: {ref.Name} "
                        f"v{ref.Major}.{ref.Minor}{builtin_str}"
                    )
                    
                except pywintypes.com_error as e:
                    logger.warning(f"Failed to read reference {i}: {e}")
                    continue
            
            logger.info(f"Listed {len(references)} VBA references")
            return references
            
        except pywintypes.com_error as e:
            logger.error(f"Failed to access VBA references: {e}")
            raise ReferenceError(f"Unable to list references: {e}")
    
    def reference_exists(
        self,
        guid: Optional[str] = None,
        name: Optional[str] = None
    ) -> bool:
        """Check if a reference exists in the VBA project.
        
        Searches by GUID (preferred) or name. GUID matching is more reliable
        since reference names can be ambiguous.
        
        Args:
            guid: Reference GUID (e.g., "{420B2830-E718-11CF-893D-00A0C9054228}")
            name: Reference name (e.g., "Scripting")
            
        Returns:
            True if reference exists, False otherwise
            
        Raises:
            ValueError: If neither guid nor name is provided
            ReferenceError: If unable to check references
            
        Example:
            >>> # Check by GUID (recommended)
            >>> exists = manager.reference_exists(
            ...     guid="{420B2830-E718-11CF-893D-00A0C9054228}"
            ... )
            >>> 
            >>> # Check by name (less reliable)
            >>> exists = manager.reference_exists(name="Scripting")
        """
        if guid is None and name is None:
            raise ValueError("Either guid or name must be provided")
        
        search_by = f"GUID {guid}" if guid else f"name '{name}'"
        logger.debug(f"Checking if reference exists by {search_by}")
        
        try:
            references = self.list_references()
            
            for ref in references:
                if guid and ref['guid'].upper() == guid.upper():
                    logger.debug(f"Reference found by GUID: {ref['name']}")
                    return True
                if name and ref['name'].lower() == name.lower():
                    logger.debug(f"Reference found by name: {ref['name']} ({ref['guid']})")
                    return True
            
            logger.debug(f"Reference not found by {search_by}")
            return False
            
        except ReferenceError:
            logger.error("Failed to check reference existence")
            raise
    
    def add_reference(
        self,
        guid: str,
        name: str,
        major: int = 1,
        minor: int = 0,
        skip_if_exists: bool = True
    ) -> bool:
        """Add a VBA reference to the document.
        
        Adds a reference by GUID. If the reference already exists and skip_if_exists
        is True, the operation is skipped with a log message.
        
        Args:
            guid: Reference GUID (e.g., "{420B2830-E718-11CF-893D-00A0C9054228}")
            name: Reference name for logging (e.g., "Scripting")
            major: Major version number (default: 1)
            minor: Minor version number (default: 0)
            skip_if_exists: If True, skip adding if reference already exists (default: True)
            
        Returns:
            True if reference was added, False if skipped (already exists)
            
        Raises:
            ValueError: If GUID format is invalid
            ReferenceError: If unable to add reference
            
        Example:
            >>> # Add Microsoft Scripting Runtime
            >>> manager.add_reference(
            ...     guid="{420B2830-E718-11CF-893D-00A0C9054228}",
            ...     name="Scripting",
            ...     major=1,
            ...     minor=0
            ... )
            True
        """
        # Validate GUID format
        if not self.GUID_PATTERN.match(guid):
            raise ValueError(
                f"Invalid GUID format: {guid}. "
                f"Expected format: {{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}}"
            )
        
        logger.debug(f"Adding reference: {name} ({guid}) v{major}.{minor}")
        
        # Check if already exists
        if skip_if_exists and self.reference_exists(guid=guid):
            logger.info(f"Reference already exists, skipping: {name} ({guid})")
            return False
        
        try:
            self.vb_project.References.AddFromGuid(guid, major, minor)
            logger.info(f"✓ Added reference: {name} ({guid}) v{major}.{minor}")
            return True
            
        except pywintypes.com_error as e:
            error_msg = str(e)
            logger.error(f"Failed to add reference {name}: {error_msg}")
            
            # Provide helpful error messages
            if "class not registered" in error_msg.lower():
                raise ReferenceError(
                    f"Reference library not found: {name} ({guid}). "
                    f"The library may not be installed on this system."
                )
            elif "invalid guid" in error_msg.lower():
                raise ReferenceError(
                    f"Invalid reference GUID: {guid}. "
                    f"The GUID may be incorrect or not registered."
                )
            else:
                raise ReferenceError(f"Failed to add reference {name}: {e}")
    
    def remove_reference(
        self,
        guid: Optional[str] = None,
        name: Optional[str] = None,
        skip_if_missing: bool = True
    ) -> bool:
        """Remove a VBA reference from the document.
        
        Removes a reference by GUID (preferred) or name. If the reference doesn't exist
        and skip_if_missing is True, the operation is skipped with a log message.
        
        NOTE: Built-in references (like VBA, Excel, Word) cannot be removed.
        
        Args:
            guid: Reference GUID (e.g., "{420B2830-E718-11CF-893D-00A0C9054228}")
            name: Reference name (e.g., "Scripting")
            skip_if_missing: If True, skip removal if reference doesn't exist (default: True)
            
        Returns:
            True if reference was removed, False if skipped (doesn't exist)
            
        Raises:
            ValueError: If neither guid nor name is provided
            ReferenceError: If unable to remove reference or if reference is built-in
            
        Example:
            >>> # Remove by GUID (recommended)
            >>> manager.remove_reference(
            ...     guid="{420B2830-E718-11CF-893D-00A0C9054228}"
            ... )
            True
            >>> 
            >>> # Remove by name
            >>> manager.remove_reference(name="Scripting")
            True
        """
        if guid is None and name is None:
            raise ValueError("Either guid or name must be provided")
        
        search_by = f"GUID {guid}" if guid else f"name '{name}'"
        logger.debug(f"Removing reference by {search_by}")
        
        try:
            # Find the reference
            refs_collection = self.vb_project.References
            target_ref = None
            
            for i in range(1, refs_collection.Count + 1):
                ref = refs_collection.Item(i)
                
                if guid and ref.Guid.upper() == guid.upper():
                    target_ref = ref
                    break
                if name and ref.Name.lower() == name.lower():
                    target_ref = ref
                    break
            
            if target_ref is None:
                if skip_if_missing:
                    logger.info(f"Reference not found, skipping removal: {search_by}")
                    return False
                else:
                    raise ReferenceError(f"Reference not found: {search_by}")
            
            # Check if it's a built-in reference
            if target_ref.BuiltIn:
                raise ReferenceError(
                    f"Cannot remove built-in reference: {target_ref.Name}. "
                    f"Built-in references are required by the Office application."
                )
            
            # Remove the reference
            ref_name = target_ref.Name
            ref_guid = target_ref.Guid
            refs_collection.Remove(target_ref)
            
            logger.info(f"✓ Removed reference: {ref_name} ({ref_guid})")
            return True
            
        except pywintypes.com_error as e:
            logger.error(f"Failed to remove reference: {e}")
            raise ReferenceError(f"Unable to remove reference: {e}")
    
    def export_to_toml(self, output_file: Union[str, Path]) -> None:
        """Export VBA references to a TOML configuration file.
        
        Creates a TOML file with all non-built-in references. Built-in references
        are excluded as they are always present and application-specific.
        
        TOML Format:
            [[references]]
            name = "Scripting"
            guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
            major = 1
            minor = 0
            description = "Microsoft Scripting Runtime"
        
        Args:
            output_file: Path to output TOML file
            
        Raises:
            ReferenceError: If unable to export references
            IOError: If unable to write to file
            
        Example:
            >>> manager.export_to_toml("references.toml")
            >>> # Creates references.toml with all non-built-in references
        """
        output_path = Path(output_file)
        logger.debug(f"Exporting references to TOML: {output_path}")
        
        try:
            references = self.list_references()
            
            # Filter out built-in references
            user_refs = [ref for ref in references if not ref['builtin']]
            
            if not user_refs:
                logger.warning("No user references to export (only built-in references found)")
            
            logger.debug(f"Exporting {len(user_refs)} user references (excluded {len(references) - len(user_refs)} built-in)")
            
            # Build TOML content manually (tomli_w not yet imported)
            # Will be replaced with tomli_w.dump() when dependency is added
            toml_content = "# VBA References Configuration\n"
            toml_content += "# Generated by vba-edit reference manager\n"
            toml_content += "# Built-in references are excluded\n\n"
            
            for ref in user_refs:
                toml_content += "[[references]]\n"
                toml_content += f'name = "{ref["name"]}"\n'
                toml_content += f'guid = "{ref["guid"]}"\n'
                toml_content += f'major = {ref["major"]}\n'
                toml_content += f'minor = {ref["minor"]}\n'
                
                if ref.get('description'):
                    # Escape quotes in description
                    desc = ref['description'].replace('"', '\\"')
                    toml_content += f'description = "{desc}"\n'
                
                toml_content += "\n"
            
            # Write to file
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(toml_content, encoding='utf-8')
            
            logger.info(f"✓ Exported {len(user_refs)} references to: {output_path}")
            
        except ReferenceError:
            logger.error("Failed to export references")
            raise
        except IOError as e:
            logger.error(f"Failed to write TOML file: {e}")
            raise ReferenceError(f"Unable to write to {output_path}: {e}")
    
    def import_from_toml(
        self,
        input_file: Union[str, Path],
        skip_existing: bool = True
    ) -> Dict[str, int]:
        """Import VBA references from a TOML configuration file.
        
        Reads references from a TOML file and adds them to the document.
        Existing references can be skipped (default) or cause an error.
        
        Args:
            input_file: Path to input TOML file
            skip_existing: If True, skip references that already exist (default: True)
            
        Returns:
            Dictionary with operation statistics:
            - added: Number of references added
            - skipped: Number of references skipped (already exist)
            - failed: Number of references that failed to add
            
        Raises:
            FileNotFoundError: If TOML file doesn't exist
            ReferenceError: If unable to parse TOML or add references
            
        Example:
            >>> stats = manager.import_from_toml("references.toml")
            >>> print(f"Added: {stats['added']}, Skipped: {stats['skipped']}")
        """
        input_path = Path(input_file)
        logger.debug(f"Importing references from TOML: {input_path}")
        
        if not input_path.exists():
            raise FileNotFoundError(f"TOML file not found: {input_path}")
        
        try:
            # Read and parse TOML
            # NOTE: Will use tomli for Python < 3.11, tomllib for Python >= 3.11
            import sys
            if sys.version_info >= (3, 11):
                import tomllib
                with open(input_path, 'rb') as f:
                    data = tomllib.load(f)
            else:
                import tomli
                with open(input_path, 'rb') as f:
                    data = tomli.load(f)
            
            if 'references' not in data:
                raise ReferenceError(
                    f"Invalid TOML file: missing 'references' section in {input_path}"
                )
            
            references = data['references']
            logger.debug(f"Found {len(references)} references in TOML file")
            
            stats = {'added': 0, 'skipped': 0, 'failed': 0}
            
            for ref in references:
                # Validate required fields
                required_fields = ['name', 'guid', 'major', 'minor']
                missing = [f for f in required_fields if f not in ref]
                if missing:
                    logger.warning(
                        f"Skipping invalid reference (missing {', '.join(missing)}): "
                        f"{ref.get('name', 'unknown')}"
                    )
                    stats['failed'] += 1
                    continue
                
                try:
                    added = self.add_reference(
                        guid=ref['guid'],
                        name=ref['name'],
                        major=ref['major'],
                        minor=ref['minor'],
                        skip_if_exists=skip_existing
                    )
                    
                    if added:
                        stats['added'] += 1
                    else:
                        stats['skipped'] += 1
                        
                except ReferenceError as e:
                    logger.warning(f"Failed to add reference {ref['name']}: {e}")
                    stats['failed'] += 1
            
            logger.info(
                f"Import complete: {stats['added']} added, "
                f"{stats['skipped']} skipped, {stats['failed']} failed"
            )
            
            return stats
            
        except ImportError as e:
            logger.error(f"TOML library not available: {e}")
            raise ReferenceError(
                "TOML parsing library not available. "
                "Please install tomli: pip install tomli"
            )
        except Exception as e:
            logger.error(f"Failed to import references from TOML: {e}")
            raise ReferenceError(f"Unable to import from {input_path}: {e}")

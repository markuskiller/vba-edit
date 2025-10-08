"""Custom argument parser formatters for enhanced help output."""

import argparse
import textwrap
from typing import Optional


class EnhancedHelpFormatter(argparse.RawDescriptionHelpFormatter):
    """Enhanced help formatter with better organization.
    
    Features:
    - Better indentation and spacing
    - Consistent formatting across all commands
    - Group support
    """

    def __init__(self, prog, indent_increment=2, max_help_position=28, width=None):
        """Initialize the enhanced help formatter.
        
        Args:
            prog: Program name
            indent_increment: Number of spaces to indent
            max_help_position: Maximum position for help text
            width: Maximum line width (None = auto-detect)
        """
        super().__init__(prog, indent_increment, max_help_position, width)
        self._section_heading = None  # Store just the heading text for reference

    def start_section(self, heading):
        """Start a new section with custom heading formatting.
        
        Args:
            heading: Section heading text
        """
        # Store heading for reference in other methods
        self._section_heading = heading
        # Capitalize first letter and add single colon for consistency
        if heading:
            # Capitalize first letter of each word for section headings
            heading = heading.title() if heading.islower() else heading
            # Remove any existing colons before adding our own
            heading = heading.rstrip(':')
        super().start_section(heading)

    def _format_action(self, action):
        """Format an individual action with better alignment.
        
        Args:
            action: Action to format
            
        Returns:
            Formatted action string
        """
        # Capitalize help text if it starts with lowercase
        if action.help and action.help[0].islower():
            action.help = action.help[0].upper() + action.help[1:]
        
        # Get original formatting
        result = super()._format_action(action)
        
        # Add extra spacing after action groups for readability
        if hasattr(self, '_section_heading') and self._section_heading and self._section_heading.lower() in ['commands', 'subcommands']:
            result += '\n'
        
        return result

    def _split_lines(self, text, width):
        """Split lines for help text with proper indentation.
        
        Args:
            text: Text to split
            width: Maximum width
            
        Returns:
            List of split lines
        """
        lines = []
        for line in text.splitlines():
            if line.strip():
                lines.extend(textwrap.wrap(line, width, break_long_words=False, break_on_hyphens=False))
            else:
                lines.append('')
        return lines


class GroupedHelpFormatter(EnhancedHelpFormatter):
    """Help formatter with explicit argument groups.
    
    Organizes arguments into logical groups:
    - Required Arguments
    - Configuration Options
    - Encoding Options
    - Header Options
    - Folder Organization Options  
    - Output Options
    - Global Options
    """

    def __init__(self, prog, indent_increment=2, max_help_position=28, width=None):
        """Initialize the grouped help formatter.
        
        Args:
            prog: Program name
            indent_increment: Number of spaces to indent
            max_help_position: Maximum position for help text
            width: Maximum line width
        """
        super().__init__(prog, indent_increment, max_help_position, width)
        self._action_groups = []

    def add_argument_group(self, title=None, description=None):
        """Add an argument group with custom formatting.
        
        Args:
            title: Group title
            description: Group description
            
        Returns:
            Argument group
        """
        group = super().add_argument_group(title, description)
        self._action_groups.append(group)
        return group


def create_parser_with_groups(parser, include_file=True, include_vba_dir=True, 
                               include_encoding=True, include_headers=True,
                               include_folders=True, include_metadata=False):
    """Create organized argument groups for a parser.
    
    This function sets up logical groupings of arguments to make help output
    clearer and more organized.
    
    Args:
        parser: ArgumentParser or subparser to add groups to
        include_file: Include file/vba_directory arguments
        include_vba_dir: Include vba_directory argument
        include_encoding: Include encoding-related arguments
        include_headers: Include header-related arguments
        include_folders: Include folder organization arguments
        include_metadata: Include metadata arguments
        
    Returns:
        Dictionary of created groups
    """
    groups = {}
    
    # Required Arguments group (if applicable)
    if include_file:
        groups['required'] = parser.add_argument_group(
            'Required Arguments',
            'Document and directory specifications'
        )
    
    # Configuration Options group
    groups['config'] = parser.add_argument_group(
        'Configuration Options',
        'Config file and verbose/logging settings'
    )
    
    # Encoding Options group
    if include_encoding:
        groups['encoding'] = parser.add_argument_group(
            'Encoding Options',
            'Character encoding for VBA files'
        )
    
    # Header Options group  
    if include_headers:
        groups['headers'] = parser.add_argument_group(
            'Header Options',
            'VBA module header storage'
        )
    
    # Folder Organization Options group
    if include_folders:
        groups['folders'] = parser.add_argument_group(
            'Folder Organization Options',
            'RubberduckVBA folder support'
        )
    
    # Metadata Options group
    if include_metadata:
        groups['metadata'] = parser.add_argument_group(
            'Metadata Options',
            'Module metadata storage'
        )
    
    # Output Options group
    groups['output'] = parser.add_argument_group(
        'Output Options',
        'Force overwrite and open folder behavior'
    )
    
    return groups


def add_examples_epilog(command_name, examples):
    """Create an examples epilog section for help text.
    
    Automatically replaces placeholder 'word-vba' with the actual command name.
    
    Args:
        command_name: Name of the command (e.g., 'excel-vba', 'word-vba')
        examples: List of (description, command) tuples
        
    Returns:
        Formatted epilog string
    """
    epilog_parts = ['\nExamples:']
    
    for desc, cmd in examples:
        # Replace word-vba with the actual command name
        cmd = cmd.replace('word-vba', command_name)
        epilog_parts.append(f'  # {desc}')
        epilog_parts.append(f'  {cmd}')
        epilog_parts.append('')
    
    return '\n'.join(epilog_parts)


# Example usage patterns for different commands (using word-vba as template)
EDIT_EXAMPLES = [
    ('Edit VBA in active document', 'word-vba edit'),
    ('Edit with specific VBA directory', 'word-vba edit --vba-directory path/to/vba'),
    ('Edit with RubberduckVBA folders', 'word-vba edit --rubberduck-folders --open-folder'),
    ('Edit with config file', 'word-vba edit --conf config.toml --verbose'),
]

IMPORT_EXAMPLES = [
    ('Import VBA from directory', 'word-vba import --file document.docm --vba-directory vba/'),
    ('Import with custom encoding', 'word-vba import -f doc.docm --encoding cp850'),
    ('Import with config file', 'word-vba import --conf config.toml'),
]

EXPORT_EXAMPLES = [
    ('Export VBA to directory', 'word-vba export --file document.docm'),
    ('Export with metadata', 'word-vba export -f doc.docm --save-metadata'),
    ('Export with inline headers', 'word-vba export --file doc.docm --in-file-headers'),
    ('Force overwrite existing', 'word-vba export --file doc.docm --force-overwrite'),
]

CHECK_EXAMPLES = [
    ('Check VBA access for Word', 'word-vba check'),
    ('Check all Office apps', 'word-vba check all'),
]

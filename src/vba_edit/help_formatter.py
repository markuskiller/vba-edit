"""Custom argument parser formatters for enhanced help output."""

import argparse
import re
import sys
import textwrap

from vba_edit.console import RICH_AVAILABLE


def strip_rich_markup(text: str) -> str:
    """Strip Rich markup tags while preserving literal brackets.

    Removes Rich styling tags like [bold], [/cyan], etc. while preserving
    literal brackets like [--file FILE], [CTRL+S], etc.

    Args:
        text: Text potentially containing Rich markup tags

    Returns:
        Text with Rich markup removed
    """
    # Remove Rich style tags: [style] or [/style] or [style params]
    text = re.sub(
        r"\[/?(?:bold|dim|cyan|bright_cyan|italic|underline|strike|reverse|blink|conceal|white|black|red|green|yellow|blue|magenta|white|default|bright_\w+|on_\w+|heading|usage|metavar|choices|command|option|action|number|path|file|success|error|warning|info)(?:\s+[^\]]+)?\]",
        "",
        text,
    )
    # Remove link markup: [link=...] ... [/link]
    text = re.sub(r"\[/?link[^\]]*\]", "", text)
    return text


def print_help_with_rich(text):
    """Print help text using rich console if available.

    This function handles the rich markup tags in help text and prints
    them with appropriate styling when rich is available.

    Args:
        text: Help text (potentially with rich markup tags)
    """
    # Import module to check if colors were disabled
    import vba_edit.console as console_module

    if RICH_AVAILABLE and not console_module._colors_disabled:
        # Use rich console to print (will render markup tags and apply highlighting)
        # Note: highlight=True allows our custom highlighter to work on plain text portions
        console_module.console.print(text, end="", highlight=True, soft_wrap=True)
    else:
        # Strip Rich markup tags but preserve literal brackets like [--file FILE]
        text = strip_rich_markup(text)
        sys.stdout.write(text)


class ColorizedArgumentParser(argparse.ArgumentParser):
    """ArgumentParser that uses rich console for help output.

    This subclass intercepts help printing and routes it through
    rich console if available and colors are enabled.

    Disables Python 3.14+ built-in argparse colors in favor of our Rich colors.
    """

    def __init__(self, *args, **kwargs):
        """Initialize parser with color support control.

        Python 3.14+ added built-in color support to argparse with a 'color' parameter.
        We always disable it (color=False) because we use Rich for custom colorization.
        This ensures our custom color scheme (uv/ruff style) is used instead of argparse's
        default colors.

        For Python <3.14, the 'color' parameter doesn't exist and is safely ignored.
        """
        # Always disable argparse's built-in colors - we use Rich instead
        # Only available in Python 3.14+, safely ignored in earlier versions
        kwargs.setdefault("color", False)

        try:
            super().__init__(*args, **kwargs)
        except TypeError as e:
            # Python <3.14 doesn't support 'color' parameter
            if "color" in str(e):
                kwargs.pop("color", None)
                super().__init__(*args, **kwargs)
            else:
                raise

    def print_help(self, file=None):
        """Print help message using rich console if available.

        Args:
            file: Output file (ignored, always uses stdout)
        """
        help_text = self.format_help()
        print_help_with_rich(help_text)

    def _print_message(self, message, file=None):
        """Print message using rich console if available.

        Args:
            message: Message to print
            file: Output file (ignored when using rich)
        """
        if message:
            print_help_with_rich(message)


class EnhancedHelpFormatter(argparse.RawDescriptionHelpFormatter):
    """Enhanced help formatter with better organization and optional colorization.

    Features:
    - Better indentation and spacing
    - Consistent formatting across all commands
    - Group support
    - Optional syntax highlighting with rich (when available and colors enabled)
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
        # NOTE: Don't cache _use_colors here - check dynamically via _should_use_colors()
        # This ensures --no-color flag is respected even when checked before parser creation

    def _should_use_colors(self):
        """Check if colors should be used (dynamic check).

        Returns:
            True if colors should be used, False otherwise
        """
        # Import module to check if colors were disabled
        import vba_edit.console as console_module

        return RICH_AVAILABLE and not console_module._colors_disabled

    def _colorize(self, text, style):
        """Apply color styling to text if colors are enabled.

        Args:
            text: Text to colorize
            style: Rich style name (e.g., 'command', 'option', 'metavar')

        Returns:
            Styled text if colors enabled, plain text otherwise
        """
        if not self._should_use_colors() or not text:
            return text
        return f"[{style}]{text}[/{style}]"

    def _format_usage(self, usage, actions, groups, prefix):
        """Format usage string with optional colorization.

        Args:
            usage: Usage string
            actions: List of actions
            groups: List of groups
            prefix: Prefix string (e.g., 'usage: ')

        Returns:
            Formatted usage string
        """
        result = super()._format_usage(usage, actions, groups, prefix)

        # NOTE: Usage colorization disabled - causes ANSI escape code leakage
        # The usage line is formatted by argparse before reaching Rich console,
        # which causes raw escape codes (like '35m') to appear in output.
        # The usage line is already clear without colorization.

        # if self._use_colors:
        #     # Colorize command names (prog name)
        #     result = re.sub(r"\b(\w+-vba)\b", lambda m: self._colorize(m.group(1), "command"), result)
        #     # Colorize optional arguments in square brackets
        #     result = re.sub(r"(\[--?[\w-]+[^\]]*\])", lambda m: self._colorize(m.group(1), "dim"), result)
        #     # Colorize required arguments in angle brackets
        #     result = re.sub(r"(<[^>]+>)", lambda m: self._colorize(m.group(1), "metavar"), result)

        return result

    def start_section(self, heading):
        """Start a new section with custom heading formatting and optional colorization.

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
            heading = heading.rstrip(":")
            # Colorize section headings (bold green in uv style)
            if self._should_use_colors():
                heading = self._colorize(heading, "heading")
        super().start_section(heading)

    def _format_action(self, action):
        """Format an individual action with better alignment and optional colorization.

        Uses 3 regex passes over the entire string instead of line-by-line processing.
        This is simpler and more maintainable than manual string splitting.

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

        # Apply colorization if enabled using 3 regex passes
        if self._should_use_colors() and result:
            # Pass 1: Colorize option flags (--option, -o)
            # Matches options at start of line or after comma/whitespace
            # Avoids matching options inside help text (after 2+ spaces)
            result = re.sub(
                r"(^|\s)(--?[\w-]+)(?=\s*(?:,|\[|[A-Z_]|\s{2,}|$))",
                lambda m: m.group(1) + self._colorize(m.group(2), "option"),
                result,
                flags=re.MULTILINE,
            )

            # Pass 2: Colorize metavars (FILE, DIR, ENCODING, [LOGFILE])
            # First, colorize metavars in square brackets (e.g., [LOGFILE])
            result = re.sub(
                r"(\[[A-Z_]{2,}\])",
                lambda m: self._colorize(m.group(1), "metavar"),
                result,
                flags=re.MULTILINE,
            )
            # Then, colorize regular metavars (uppercase words before help text or at end of line)
            result = re.sub(
                r"\b([A-Z_]{2,})\b(?=\s{2,}|$)",
                lambda m: self._colorize(m.group(1), "metavar"),
                result,
                flags=re.MULTILINE,
            )

            # Pass 3: Colorize command names in Commands/Subcommands section
            # Only if we're in the commands section
            if (
                hasattr(self, "_section_heading")
                and self._section_heading
                and self._section_heading.lower() in ["commands", "subcommands"]
            ):
                # Matches command name at start of line (after whitespace) followed by 2+ spaces
                result = re.sub(
                    r"^(\s+)(\w+)(?=\s{2,})",
                    lambda m: m.group(1) + self._colorize(m.group(2), "command"),
                    result,
                    flags=re.MULTILINE,
                )

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
                lines.append("")
        return lines

    def format_help(self):
        """Format help text with optional rich rendering.

        Returns:
            Formatted help text
        """
        help_text = super().format_help()

        # If rich is available and colors enabled, use rich console to print
        # (The markup tags [option], [command], etc. will be rendered)
        if self._should_use_colors() and RICH_AVAILABLE:
            # Return the help text with markup - argparse will print it,
            # but we've already added the markup tags
            return help_text

        return help_text


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


def create_parser_with_groups(
    parser,
    include_file=True,
    include_vba_dir=True,
    include_encoding=True,
    include_headers=True,
    include_folders=True,
    include_metadata=False,
):
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
        groups["required"] = parser.add_argument_group("Required Arguments", "Document and directory specifications")

    # Configuration Options group
    groups["config"] = parser.add_argument_group("Configuration Options", "Config file and verbose/logging settings")

    # Encoding Options group
    if include_encoding:
        groups["encoding"] = parser.add_argument_group("Encoding Options", "Character encoding for VBA files")

    # Header Options group
    if include_headers:
        groups["headers"] = parser.add_argument_group("Header Options", "VBA module header storage")

    # Folder Organization Options group
    if include_folders:
        groups["folders"] = parser.add_argument_group("Folder Organization Options", "RubberduckVBA folder support")

    # Metadata Options group
    if include_metadata:
        groups["metadata"] = parser.add_argument_group("Metadata Options", "Module metadata storage")

    # Output Options group
    groups["output"] = parser.add_argument_group("Output Options", "Force overwrite and open folder behavior")

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
    epilog_parts = ["\nExamples:"]

    for desc, cmd in examples:
        # Replace word-vba with the actual command name
        cmd = cmd.replace("word-vba", command_name)
        epilog_parts.append(f"  # {desc}")
        epilog_parts.append(f"  {cmd}")
        epilog_parts.append("")

    return "\n".join(epilog_parts)


# Example usage patterns for different commands (using word-vba as template)
EDIT_EXAMPLES = [
    ("Edit VBA in active document", "word-vba edit"),
    ("Edit with specific VBA directory", "word-vba edit --vba-directory path/to/vba"),
    ("Edit with RubberduckVBA folders", "word-vba edit --rubberduck-folders --open-folder"),
    ("Edit with config file", "word-vba edit --conf config.toml --verbose"),
]

IMPORT_EXAMPLES = [
    ("Import VBA from directory", "word-vba import --file document.docm --vba-directory vba/"),
    ("Import with custom encoding", "word-vba import -f doc.docm --encoding cp850"),
    ("Import with config file", "word-vba import --conf config.toml"),
]

EXPORT_EXAMPLES = [
    ("Export VBA to directory", "word-vba export --file document.docm"),
    ("Export with metadata", "word-vba export -f doc.docm --save-metadata"),
    ("Export with inline headers", "word-vba export --file doc.docm --in-file-headers"),
    ("Force overwrite existing", "word-vba export --file doc.docm --force-overwrite"),
    ("Keep document open after export", "word-vba export --file doc.docm --keep-open"),
]

CHECK_EXAMPLES = [
    ("Check VBA access for Word", "word-vba check"),
    ("Check all Office apps", "word-vba check all"),
]

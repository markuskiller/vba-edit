"""Console output utilities with rich styling (uv/ruff-style aesthetics).

This module provides colorized console output using the rich library.
It automatically handles:
- Terminal detection (no colors in pipes/files)
- NO_COLOR environment variable
- --no-color CLI flag
- Cross-platform support (Windows, Linux, macOS)

Usage:
    from vba_edit.console import success, warning, error, info, console

    success("Export completed successfully!")
    warning("Found existing files")
    error("Failed to open document")
    info("Using document: workbook.xlsm")

    # Advanced usage with markup
    console.print("[success]✓[/success] Exported: [path]Module1.bas[/path]")
"""

import re
import sys

try:
    from rich.console import Console
    from rich.highlighter import Highlighter
    from rich.theme import Theme
    from rich.text import Text

    class HelpTextHighlighter(Highlighter):
        """Custom highlighter for CLI help text.

        Highlights:
        - Technical terms (TOML, JSON, XML, HTTP, etc.) in bright blue
        - Example commands (lines starting with whitespace + command name) in dim
        - Does NOT highlight random capitalized words
        - Works alongside Rich markup tags
        """

        # Technical terms/formats we want to highlight
        TECH_TERMS = {
            # File formats and data standards
            "TOML",
            "JSON",
            "XML",
            "YAML",
            "HTML",
            "CSV",
            "SQL",
            # Protocols and web
            "HTTP",
            "HTTPS",
            "FTP",
            "SSH",
            "API",
            "REST",
            "URI",
            "URL",
            # Encodings
            "UTF-8",
            "UTF8",
            "ASCII",
            "cp1252",
            # VBA and Office
            "VBA",
            "COM",
            "Module",
            "Class",
            "Form",
            "UserForm",
            "Macro",
            "Procedure",
            "Function",
            # Office applications
            "Excel",
            "Word",
            # NOTE: "Access" intentionally excluded - conflicts with "Trust Access to VBA"
            # Use "MS Access" instead to refer to Microsoft Access application
            "MS Access",
            "PowerPoint",
            "Outlook",
            "Project",
            "Visio",
            # File extensions and filenames
            ".bas",
            ".cls",
            ".frm",
            "vba_edit.log",
            # ".xlsm", ".docm", ".accdb", ".pptm",
            # Dev tools and libraries
            "RubberduckVBA",
            "@Folder",
            "xlwings",
            "xlwings vba",
            "VS Code",
            "Git",
            # Operations (lowercase for command descriptions)
            "export",
            "import",
            "sync",
            "watch",
            "watches",
            "edit",
            # Command verbs (capitalized for emphasis in sentences)
            "Edit",
            "Export",
            "Import",
            "Check",
            "Sync",
            # Keyboard shortcuts
            "[CTRL+S]",
            "[SHIFT]",
            "[ALT]",
            "[ENTER]",
            "[RETURN]",
        }

        def highlight(self, text: Text) -> None:
            """Apply highlighting to text.

            This works by adding styles to specific ranges without
            overriding existing styles from markup tags.

            IMPORTANT: Order matters! We dim example lines FIRST, then
            check for dim style when adding technical term highlights.
            """
            plain_text = text.plain
            import re

            # STEP 1: Highlight example command lines and continuation lines FIRST
            # This must happen before technical term highlighting so we can detect dimmed regions
            lines = plain_text.split("\n")
            offset = 0
            in_example = False  # Track if we're in a multi-line example
            in_usage = False  # Track if we're in the usage synopsis (indented options)
            in_important = False  # Track if we're in an IMPORTANT warning

            for line in lines:
                # Check if this is an IMPORTANT warning (starts with "IMPORTANT:")
                if line.strip().startswith("IMPORTANT:"):
                    in_important = True
                    line_start = offset
                    line_end = offset + len(line)
                    text.stylize("dim yellow", line_start, line_end)

                # Check if we're continuing an IMPORTANT warning (indented to column 12)
                elif in_important and len(line) >= 11 and line[:11] == " " * 11 and line[11] != " ":
                    # Line starts at column 12 (11 spaces + content) - continuation of IMPORTANT
                    line_start = offset
                    line_end = offset + len(line)
                    text.stylize("dim yellow", line_start, line_end)

                # Check if IMPORTANT warning ended (empty line or non-matching indent)
                elif in_important and (not line.strip() or not (len(line) >= 11 and line[:11] == " " * 11)):
                    in_important = False

                # Check if we're in usage synopsis (indented lines with brackets)
                # Don't dim the first "usage:" line itself, only the indented option lines
                if in_usage and re.match(r"^\s+\[", line):
                    # Usage synopsis option line - make it dim
                    line_start = offset
                    line_end = offset + len(line)
                    text.stylize("dim", line_start, line_end)

                # Check if this is the start of usage synopsis
                elif line.startswith("usage:"):
                    in_usage = True
                    # Don't dim the usage: line itself - let technical terms be highlighted

                # Check if usage synopsis ended (empty line or non-indented line that's not a bracket line)
                elif in_usage and not line.strip():
                    in_usage = False

                # Check if line looks like an example (starts with whitespace, contains command)
                # Matches: "  excel-vba edit" or "  excel-vba edit -f file.xlsm"
                if re.match(r"^\s{2,}(excel-vba|word-vba|access-vba|powerpoint-vba)\s+\w+", line):
                    # This is an example line - make it dim
                    line_start = offset
                    line_end = offset + len(line)
                    text.stylize("dim", line_start, line_end)
                    in_example = True

                # Check if this is a continuation line (indented, starts with #)
                elif in_example and re.match(r"^\s+#", line):
                    # This is a comment continuation - also make it dim
                    line_start = offset
                    line_end = offset + len(line)
                    text.stylize("dim", line_start, line_end)

                else:
                    # Not an example or continuation - reset the flag
                    in_example = False

                offset += len(line) + 1  # +1 for newline

            # STEP 2: Now highlight technical terms (but skip dimmed regions)
            for term in self.TECH_TERMS:
                # Find all occurrences of the term
                # Use word boundaries only if term starts/ends with word characters
                escaped_term = re.escape(term)
                if term[0].isalnum() and term[-1].isalnum():
                    # Standard word - use word boundaries
                    pattern = rf"\b({escaped_term})\b"
                else:
                    # Special characters (like [CTRL+S]) - no word boundaries
                    pattern = rf"({escaped_term})"

                for match in re.finditer(pattern, plain_text):
                    start, end = match.span(1)
                    # Only apply if this range doesn't already have a style
                    # Check if any span overlaps with this range
                    # Also skip if the text is already dimmed (examples, usage synopsis)
                    has_existing_style = False
                    is_dimmed = False
                    for span_start, span_end, style in text.spans:
                        if not (end <= span_start or start >= span_end):
                            # Overlaps - check if it's a dim style
                            if style and "dim" in str(style):
                                is_dimmed = True
                                break
                            # Other overlap - skip this highlighting
                            has_existing_style = True
                            break

                    if not has_existing_style and not is_dimmed:
                        text.stylize("cyan", start, end)

    # Define our color theme (uv-style - October 2025)
    custom_theme = Theme(
        {
            # Message types
            "success": "bold green",
            "error": "bold red",
            "warning": "bold yellow",
            "info": "cyan",
            # Elements
            "dim": "dim",
            "path": "cyan",  # uv: regular cyan for paths (no bold, no bright)
            "file": "cyan",  # uv: regular cyan for files (no bold, no bright)
            "command": "bold bright_cyan",  # uv: bold bright cyan for commands
            "option": "bold bright_cyan",  # uv: bold bright cyan for options like --file, -f
            "action": "green",  # uv: green for actions
            "number": "cyan",  # uv: cyan for numbers
            # Help text styling (matches uv October 2025)
            "heading": "bold bright_green",  # Section headings: File Options, etc.
            "usage": "bold white",  # Usage line
            "metavar": "cyan",  # Metavars: FILE, DIR, etc. (regular cyan)
            "choices": "dim cyan",  # Choices in dim cyan
        }
    )

    # Create console instances with custom highlighter
    # Uses our HelpTextHighlighter for smart, predictable highlighting:
    # - Technical terms (TOML, JSON, etc.) in bright blue
    # - Example command lines in dim grey
    # - No random capitalized word highlighting
    help_highlighter = HelpTextHighlighter()
    console = Console(theme=custom_theme, highlight=False, highlighter=help_highlighter)
    error_console = Console(stderr=True, theme=custom_theme, highlight=False, highlighter=help_highlighter)

    RICH_AVAILABLE = True

except ImportError:
    # Fallback: create dummy console that strips markup
    class DummyConsole:
        """Fallback console when rich is not available."""

        def __init__(self, stderr=False):
            self.file = sys.stderr if stderr else sys.stdout
            self.no_color = True

        def print(self, *args, **kwargs):
            """Print with rich markup stripped."""
            # Combine args into single string
            text = " ".join(str(arg) for arg in args)
            # Remove [style]...[/style] markup
            text = re.sub(r"\[/?[^\]]+\]", "", text)
            # Filter kwargs to only print-compatible ones
            print_kwargs = {k: v for k, v in kwargs.items() if k in ("file", "end", "sep")}
            print_kwargs.setdefault("file", self.file)
            print(text, **print_kwargs)

    console = DummyConsole()
    error_console = DummyConsole(stderr=True)
    RICH_AVAILABLE = False


def success(message: str, **kwargs):
    """Print success message with checkmark (green).

    Args:
        message: The success message to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        success("Export completed successfully!")
        # Output: ✓ Export completed successfully! (in green)
    """
    console.print(f"[success]✓[/success] {message}", **kwargs)


def error(message: str, **kwargs):
    """Print error message with X mark (red).

    Args:
        message: The error message to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        error("Failed to open document")
        # Output: ✗ Failed to open document (in red)
    """
    error_console.print(f"[error]✗[/error] {message}", **kwargs)


def warning(message: str, **kwargs):
    """Print warning message with warning symbol (yellow).

    Args:
        message: The warning message to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        warning("Found existing files")
        # Output: ⚠ Found existing files (in yellow)
    """
    console.print(f"[warning]⚠[/warning] {message}", **kwargs)


def info(message: str, **kwargs):
    """Print info message (cyan).

    Args:
        message: The info message to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        info("Using document: workbook.xlsm")
        # Output: Using document: workbook.xlsm (in cyan)
    """
    console.print(f"[info]{message}[/info]", **kwargs)


def dim(message: str, **kwargs):
    """Print dimmed message (gray).

    Args:
        message: The message to print in dim style
        **kwargs: Additional arguments passed to console.print()

    Example:
        dim("(Press Ctrl+C to stop)")
        # Output: (Press Ctrl+C to stop) (in gray)
    """
    console.print(f"[dim]{message}[/dim]", **kwargs)


def print_command(command: str, **kwargs):
    """Print command name (cyan bold).

    Args:
        command: The command name to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        print_command("excel-vba export")
        # Output: excel-vba export (in cyan bold)
    """
    console.print(f"[command]{command}[/command]", **kwargs)


def print_path(path: str, **kwargs):
    """Print file path (blue).

    Args:
        path: The file path to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        print_path("src/vba_edit/Module1.bas")
        # Output: src/vba_edit/Module1.bas (in blue)
    """
    console.print(f"[path]{path}[/path]", **kwargs)


def print_action(action: str, **kwargs):
    """Print action (magenta).

    Args:
        action: The action text to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        print_action("Watching for changes...")
        # Output: Watching for changes... (in magenta)
    """
    console.print(f"[action]{action}[/action]", **kwargs)


def disable_colors():
    """Disable all colors (for --no-color flag).

    This function sets the no_color flag on both console instances,
    which causes them to output plain text without any styling.
    """
    global console, error_console

    # Replace with DummyConsole to fully disable color and markup
    class DummyConsole:
        def __init__(self, stderr=False):
            self.file = sys.stderr if stderr else sys.stdout
            self.no_color = True

        def print(self, *args, **kwargs):
            text = " ".join(str(arg) for arg in args)
            # Strip Rich markup tags but preserve literal brackets like [CTRL+S]
            # Only remove tags that are Rich styles: [style] or [/style] or [link=...] patterns
            # This regex matches: [word] or [/word] or [word params] but NOT [WORD+WORD] or [WORD-WORD]
            text = re.sub(
                r"\[/?(?:bold|dim|cyan|bright_cyan|italic|underline|strike|reverse|blink|conceal|white|black|red|green|yellow|blue|magenta|white|default|bright_\w+|on_\w+)(?:\s+[^\]]+)?\]",
                "",
                text,
            )
            # Also remove link markup: [link=...] ... [/link]
            text = re.sub(r"\[/?link[^\]]*\]", "", text)
            print_kwargs = {k: v for k, v in kwargs.items() if k in ("file", "end", "sep")}
            print_kwargs.setdefault("file", self.file)
            print(text, **print_kwargs)

    console = DummyConsole()
    error_console = DummyConsole(stderr=True)


def enable_colors():
    """Enable colors (opposite of disable_colors).

    Only has effect if rich is available and terminal supports colors.
    """
    if RICH_AVAILABLE:
        console.no_color = False
        error_console.no_color = False


# Export public API
__all__ = [
    # Console instances
    "console",
    "error_console",
    # Message functions
    "success",
    "error",
    "warning",
    "info",
    "dim",
    # Formatting functions
    "print_command",
    "print_path",
    "print_action",
    # Control functions
    "disable_colors",
    "enable_colors",
    # Status
    "RICH_AVAILABLE",
]

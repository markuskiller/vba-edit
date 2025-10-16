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

import sys

# Module-level flag to track if colors are disabled
# This is set by disable_colors() and checked by help formatter
_colors_disabled = False

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
        # Used by both help formatter AND semantic log formatter for consistency
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
            # VBA file extensions
            ".bas",
            ".cls",
            ".frm",
            ".frx",
            # Office file extensions
            ".xlsm",
            ".xlsb",
            ".xlsx",
            ".docm",
            ".docx",
            ".dotm",
            ".accdb",
            ".accde",
            ".pptm",
            ".pptx",
            ".potm",
            # Other file extensions
            ".log",
            ".txt",
            ".toml",
            ".json",
            ".xml",
            # Filenames
            "vba_edit.log",
            # Dev tools and libraries
            "RubberduckVBA",
            "@Folder",
            "xlwings",
            "xlwings vba",
            "VS Code",
            "Git",
            # Commands and operations (used in both help and logging)
            "export",
            "import",
            "sync",
            "watch",
            "watches",
            "edit",
            "check",
            # Command verbs (capitalized for emphasis in sentences)
            "Edit",
            "Export",
            "Import",
            "Check",
            "Sync",
            "Starting",  # Used in "Starting edit session" messages
            # Action verbs (lowercase - used in progress messages)
            "Opening",
            "opening",
            "exported",
            "imported",
            "editing",
            # Keyboard shortcuts
            "[CTRL+S]",
            "[SHIFT]",
            "[ALT]",
            "[ENTER]",
            "[RETURN]",
        }

        def _dim_important_warnings(self, text: Text, lines: list[str], offset_tracker: list[int]) -> None:
            """Dim IMPORTANT warnings (first priority).

            Args:
                text: Rich Text object to apply styles to
                lines: List of text lines
                offset_tracker: Single-item list tracking current offset [offset]
            """

            offset = offset_tracker[0]
            in_important = False

            for line in lines:
                if line.strip().startswith("IMPORTANT:"):
                    in_important = True
                    text.stylize("dim yellow", offset, offset + len(line))
                elif in_important and len(line) >= 11 and line[:11] == " " * 11 and line[11] != " ":
                    # Continuation line (indented to column 12)
                    text.stylize("dim yellow", offset, offset + len(line))
                elif in_important and (not line.strip() or len(line) < 11 or line[:11] != " " * 11):
                    in_important = False

                offset += len(line) + 1
            offset_tracker[0] = 0  # Reset for next pass

        def _dim_usage_synopsis(self, text: Text, lines: list[str], offset_tracker: list[int]) -> None:
            """Dim usage synopsis option lines (second priority).

            Args:
                text: Rich Text object to apply styles to
                lines: List of text lines
                offset_tracker: Single-item list tracking current offset [offset]
            """
            import re

            offset = offset_tracker[0]
            in_usage = False

            for line in lines:
                if line.startswith("usage:"):
                    in_usage = True
                elif in_usage and not line.strip():
                    in_usage = False
                elif in_usage and re.match(r"^\s+\[", line):
                    text.stylize("dim", offset, offset + len(line))

                offset += len(line) + 1
            offset_tracker[0] = 0  # Reset for next pass

        def _dim_example_lines(self, text: Text, lines: list[str], offset_tracker: list[int]) -> None:
            """Dim example command lines and comment continuations (third priority).

            Args:
                text: Rich Text object to apply styles to
                lines: List of text lines
                offset_tracker: Single-item list tracking current offset [offset]
            """
            import re

            offset = offset_tracker[0]
            in_example = False

            for line in lines:
                if re.match(r"^\s{2,}(excel-vba|word-vba|access-vba|powerpoint-vba)\s+\w+", line):
                    text.stylize("dim", offset, offset + len(line))
                    in_example = True
                elif in_example and re.match(r"^\s+#", line):
                    text.stylize("dim", offset, offset + len(line))
                else:
                    in_example = False

                offset += len(line) + 1
            offset_tracker[0] = 0  # Reset for next pass

        def _highlight_technical_terms(self, text: Text, plain_text: str) -> None:
            """Highlight technical terms (final pass - skips dimmed regions).

            Args:
                text: Rich Text object to apply styles to
                plain_text: Plain text version for pattern matching
            """
            import re

            for term in self.TECH_TERMS:
                # Build pattern with word boundaries for alphanumeric terms
                escaped_term = re.escape(term)
                if term[0].isalnum() and term[-1].isalnum():
                    pattern = rf"\b({escaped_term})\b"
                else:
                    pattern = rf"({escaped_term})"

                for match in re.finditer(pattern, plain_text):
                    start, end = match.span(1)
                    # Skip if already styled or dimmed
                    if self._is_range_styled_or_dimmed(text, start, end):
                        continue
                    text.stylize("cyan", start, end)

        def _is_range_styled_or_dimmed(self, text: Text, start: int, end: int) -> bool:
            """Check if a range already has styling or is dimmed.

            Args:
                text: Rich Text object
                start: Start position
                end: End position

            Returns:
                True if range is already styled or dimmed
            """
            return any(end > span_start and start < span_end for span_start, span_end, style in text.spans)

        def highlight(self, text: Text) -> None:
            """Apply highlighting to text.

            This works by adding styles to specific ranges without
            overriding existing styles from markup tags.

            IMPORTANT: Order matters! We dim example lines FIRST, then
            check for dim style when adding technical term highlights.
            """

            plain_text = text.plain
            lines = plain_text.split("\n")
            offset_tracker = [0]  # Mutable list to pass offset by reference

            # Apply styling in priority order
            self._dim_important_warnings(text, lines, offset_tracker)
            self._dim_usage_synopsis(text, lines, offset_tracker)
            self._dim_example_lines(text, lines, offset_tracker)
            self._highlight_technical_terms(text, plain_text)

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
            # Import at function level to avoid circular imports
            from vba_edit.help_formatter import strip_rich_markup

            # Combine args into single string
            text = " ".join(str(arg) for arg in args)
            # Remove Rich markup using shared utility
            text = strip_rich_markup(text)
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


def print_exception(exc: Exception, **kwargs):
    """Print exception message with Rich markup rendering.

    This function properly renders Rich markup tags in exception messages,
    making IMPORTANT warnings, commands, and paths colorized.
    Use this instead of logger.error(str(e)) for exceptions that contain
    Rich markup tags like [warning], [command], [path], etc.

    Args:
        exc: The exception to print
        **kwargs: Additional arguments passed to console.print()

    Example:
        try:
            raise DocumentClosedError()
        except DocumentClosedError as e:
            print_exception(e)
            # Output: Colorized message with yellow IMPORTANT, cyan commands, etc.
    """
    # Print the exception message with markup rendering enabled
    error_console.print(str(exc), **kwargs)


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
    global console, error_console, _colors_disabled

    # Set module-level flag for help formatter
    _colors_disabled = True

    # Import at function level to avoid circular imports
    from vba_edit.help_formatter import strip_rich_markup

    # Replace with DummyConsole to fully disable color and markup
    class DummyConsole:
        def __init__(self, stderr=False):
            self.file = sys.stderr if stderr else sys.stdout
            self.no_color = True

        def print(self, *args, **kwargs):
            text = " ".join(str(arg) for arg in args)
            # Strip Rich markup using shared utility
            text = strip_rich_markup(text)
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

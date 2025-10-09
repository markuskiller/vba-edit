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
    from rich.theme import Theme

    # Define our color theme (uv/ruff-style)
    custom_theme = Theme(
        {
            # Message types
            "success": "bold green",
            "error": "bold red",
            "warning": "bold yellow",
            "info": "cyan",
            # Elements
            "dim": "dim",
            "path": "blue",
            "file": "blue",
            "command": "bold cyan",
            "option": "bold magenta",
            "action": "magenta",
            "number": "yellow",
            # Help text
            "usage": "bold",
            "metavar": "bold cyan",
            "choices": "dim",
        }
    )

    # Create console instances
    console = Console(theme=custom_theme, highlight=False)
    error_console = Console(stderr=True, theme=custom_theme, highlight=False)

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
    console.no_color = True
    error_console.no_color = True


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

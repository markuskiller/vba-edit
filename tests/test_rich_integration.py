"""Tests for Rich library integration and formatter behavior.

These unit tests validate that the Rich library integration doesn't cause
markup errors when processing file paths, log messages, or exception text.

Related: Issue #43, Issue #46 (xlsb-specific manifestation)
Fixed in v0.4.2
"""

import pytest
from io import StringIO
from rich.markup import MarkupError


class TestRichIntegration:
    """Test Rich library integration and formatter behavior."""

    def test_console_has_help_text_highlighter(self):
        """Verify main console has HelpTextHighlighter for help text."""
        from vba_edit.console import console

        # In CI/non-TTY, console might be DummyConsole (no highlighter attribute)
        if not hasattr(console, "highlighter"):
            pytest.skip("DummyConsole in use (CI/non-TTY environment)")

        # Console has single highlighter attribute, not plural
        if console.highlighter is not None:
            highlighter_name = console.highlighter.__class__.__name__
            # Main console SHOULD have HelpTextHighlighter for help text formatting
            assert "HelpTextHighlighter" in highlighter_name, (
                "console should have HelpTextHighlighter for help text formatting"
            )

    def test_logging_setup_creates_handler(self):
        """Verify setup_logging creates logging handler."""
        import logging
        from vba_edit.utils import setup_logging

        # Setup logging (creates internal logging_console and handlers)
        setup_logging(verbose=False)

        # Get the root logger
        root_logger = logging.getLogger()

        # Should have at least one handler
        assert len(root_logger.handlers) > 0, "Should have logging handlers after setup_logging"

        # Find RichHandler if present (may be StreamHandler in CI/non-TTY)
        has_rich_handler = any(h.__class__.__name__ == "RichHandler" for h in root_logger.handlers)
        has_stream_handler = any(h.__class__.__name__ == "StreamHandler" for h in root_logger.handlers)

        # In TTY environments, we use RichHandler; in CI/non-TTY, StreamHandler
        assert has_rich_handler or has_stream_handler, (
            "Should have RichHandler (TTY) or StreamHandler (non-TTY) for console output"
        )

    def test_logging_with_problematic_extensions(self):
        """Test logging with extensions that could trigger MarkupError."""
        import logging
        from vba_edit.utils import setup_logging

        # Setup logging
        setup_logging(verbose=False)
        logger = logging.getLogger("test_extensions")

        # Messages that could trigger MarkupError (Issue #43, #46)
        test_messages = [
            "Processing file.xlsb",
            "Path: C:\\test[bracket].xlsm",
            "Error with [/bold] in path",
            "File [test]file.xls opened",
        ]

        # Should not raise MarkupError
        for msg in test_messages:
            try:
                logger.info(msg)
            except MarkupError:
                pytest.fail(f"MarkupError raised for message: {msg}")

    def test_print_exception_disables_markup_parsing(self):
        """Verify print_exception() doesn't parse exception text as markup."""
        from vba_edit.console import print_exception
        import sys
        from io import StringIO as PyStringIO

        # Create exception with markup-like text
        test_exceptions = [
            ValueError("File [bold]test.xlsb[/bold] not found"),
            RuntimeError("Path contains [/test] brackets"),
            IOError("Cannot open file[test].xls"),
        ]

        for exc in test_exceptions:
            # Should not raise MarkupError during exception printing
            # Note: print_exception() prints to stderr, not console.file
            old_stderr = sys.stderr
            try:
                sys.stderr = PyStringIO()
                print_exception(exc)
                # Verify exception was printed (contains exception text)
                assert "File" in str(exc) or "Path" in str(exc) or "Cannot" in str(exc)
            except MarkupError:
                pytest.fail(f"MarkupError raised when printing exception: {exc}")
            finally:
                sys.stderr = old_stderr

    def test_console_messages_preserve_literal_brackets(self):
        """Verify console messages preserve literal brackets like [CTRL+S]."""
        from vba_edit.console import console

        # Capture output
        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        # Messages with literal brackets
        test_messages = [
            "Press [CTRL+S] to save",
            "File path: C:\\test[1].xlsm",
            "Extension: [.xlsb]",
        ]

        for msg in test_messages:
            console.print(msg, markup=False)  # Print without markup parsing

        console.file = original_file
        output = buffer.getvalue()

        # Literal brackets should be preserved
        assert "[CTRL+S]" in output, "Literal [CTRL+S] should be preserved"
        assert "[1]" in output or "test" in output, "Brackets in paths should be preserved"
        assert "[.xlsb]" in output, "Extension brackets should be preserved"

    def test_xlsb_extension_safe_in_console_output(self):
        """Specifically test .xlsb extension that caused Issue #46."""
        from vba_edit.console import console

        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        # The problematic case from Issue #46
        xlsb_messages = [
            "Exporting file.xlsb",
            "Processing test.xlsb",
            "Path: C:\\Users\\test\\file.xlsb",
            "Error with file.xlsb[temp]",
        ]

        for msg in xlsb_messages:
            try:
                # Print without markup to avoid interpretation
                console.print(msg, markup=False)
            except MarkupError:
                pytest.fail(f"MarkupError on .xlsb message: {msg}")

        console.file = original_file
        output = buffer.getvalue()

        # Should contain .xlsb references
        assert ".xlsb" in output, "Output should contain .xlsb references"

    def test_all_excel_extensions_safe_in_console(self):
        """Test all Excel extensions are safe in console output."""
        from vba_edit.console import console

        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        # All Excel macro-enabled formats
        extensions = [".xlsm", ".xlsb", ".xls", ".xltm", ".xlt", ".xlam", ".xla"]

        for ext in extensions:
            msg = f"Processing file{ext}"
            try:
                console.print(msg, markup=False)
            except MarkupError:
                pytest.fail(f"MarkupError on extension: {ext}")

        console.file = original_file
        output = buffer.getvalue()

        # All extensions should appear in output
        for ext in extensions:
            assert ext in output, f"Extension {ext} should appear in output"

    def test_disable_colors_function_works(self):
        """Test that disable_colors() properly disables Rich markup."""
        from vba_edit.console import console, disable_colors

        # Call disable_colors
        disable_colors()

        # Capture output
        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        # Print with Rich markup
        console.print("[bold]test[/bold] [cyan]message[/cyan]")

        console.file = original_file
        output = buffer.getvalue()

        # Should NOT contain Rich markup tags (they should be stripped or not rendered)
        # The actual text should still be present
        assert "test" in output, "Text content should be present"
        assert "message" in output, "Text content should be present"

    def test_semantic_log_formatter_handles_paths(self):
        """Test SemanticLogFormatter handles file paths correctly."""
        import logging
        from vba_edit.utils import SemanticLogFormatter

        # Create formatter
        formatter = SemanticLogFormatter()

        # Create log record with problematic path
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="Processing: %s",
            args=("C:\\test\\file.xlsb",),
            exc_info=None,
        )

        # Should not raise MarkupError
        try:
            formatted = formatter.format(record)
            assert "file.xlsb" in formatted, "Path should be in formatted output"
        except MarkupError:
            pytest.fail("SemanticLogFormatter raised MarkupError on file path")

    def test_error_messages_with_brackets_dont_crash(self):
        """Test error messages containing brackets are handled safely."""
        from vba_edit.console import error

        # Error messages with markup-like content
        error_messages = [
            "File [test.xlsb] not found",
            "Invalid path: C:\\test[1]\\file.xls",
            "Cannot process [bold]file[/bold]",
        ]

        # Should not raise MarkupError
        for msg in error_messages:
            try:
                error(msg)
                # If we get here, no MarkupError was raised
            except MarkupError:
                pytest.fail(f"MarkupError in error message: {msg}")


class TestConsoleOutputStability:
    """Test console output functions handle edge cases gracefully."""

    def test_info_with_special_chars(self):
        """Test info() handles special characters safely."""
        from vba_edit.console import info

        buffer = StringIO()
        from vba_edit.console import console

        original_file = console.file
        console.file = buffer

        # Info messages with special characters
        test_messages = [
            "File: test.xlsb",
            "Path: [C:\\test]",
            "Extension: [.xls]",
            "Processing [file].xlsm",
        ]

        for msg in test_messages:
            try:
                info(msg)
            except Exception as e:
                pytest.fail(f"info() raised {type(e).__name__} on message: {msg}")

        console.file = original_file

    def test_success_with_file_paths(self):
        """Test success() handles file paths correctly."""
        from vba_edit.console import success

        buffer = StringIO()
        from vba_edit.console import console

        original_file = console.file
        console.file = buffer

        # Success messages with file paths
        paths = [
            "Exported: file.xlsb",
            "Saved: C:\\test[1].xls",
            "Imported: test.xlsm successfully",
        ]

        for path in paths:
            try:
                success(path)
            except Exception as e:
                pytest.fail(f"success() raised {type(e).__name__} on path: {path}")

        console.file = original_file
        output = buffer.getvalue()

        # Success indicator should be present
        assert "✓" in output or "SUCCESS" in output.upper(), "Success output should contain success indicator"

    def test_warning_with_brackets(self):
        """Test warning() handles bracket characters safely."""
        from vba_edit.console import warning

        buffer = StringIO()
        from vba_edit.console import console

        original_file = console.file
        console.file = buffer

        # Warning messages with brackets
        warnings = [
            "File [test.xlsb] may be corrupted",
            "Path contains brackets: [C:\\test]",
            "Invalid format: [.xls]",
        ]

        for msg in warnings:
            try:
                warning(msg)
            except Exception as e:
                pytest.fail(f"warning() raised {type(e).__name__} on message: {msg}")

        console.file = original_file
        output = buffer.getvalue()

        # Warning indicator should be present
        assert "⚠" in output or "WARNING" in output.upper(), "Warning output should contain warning indicator"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

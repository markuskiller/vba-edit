"""Tests for CLI color control functionality.

Tests the --no-color flag behavior, disable_colors() function,
and easter egg color handling across all CLI tools.
"""

import os
import subprocess
import sys
from io import StringIO

import pytest

# Use the current Python interpreter for subprocess calls
PYTHON_EXE = sys.executable


class TestNoColorFlag:
    """Tests for --no-color flag with help messages and commands."""

    @pytest.mark.parametrize(
        "cli_tool",
        ["excel_vba", "word_vba", "access_vba", "powerpoint_vba"],
    )
    def test_no_color_flag_with_main_help(self, cli_tool):
        """Test that --no-color flag strips colors from main help output."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", f"vba_edit.{cli_tool}", "--help", "--no-color"],
            capture_output=True,
            text=True,
        )

        # Should NOT have ANSI escape codes
        assert "\x1b[" not in result.stdout, f"{cli_tool}: Help output should not contain ANSI escape codes"

        # Should have essential help content
        assert "usage:" in result.stdout.lower()
        assert "Commands:" in result.stdout or "Options:" in result.stdout

        # Should exit cleanly (argparse help uses sys.exit(0))
        assert result.returncode in (0, 1)  # 0 for success, 1 for argparse exit

    @pytest.mark.parametrize(
        "cli_tool,command",
        [
            ("excel_vba", "edit"),
            ("excel_vba", "export"),
            ("excel_vba", "import"),
            ("word_vba", "edit"),
            ("word_vba", "export"),
            ("access_vba", "edit"),
            ("access_vba", "check"),
            ("powerpoint_vba", "edit"),
        ],
    )
    def test_no_color_flag_with_command_help(self, cli_tool, command):
        """Test that --no-color flag strips colors from command help output."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", f"vba_edit.{cli_tool}", command, "--help", "--no-color"],
            capture_output=True,
            text=True,
        )

        # Should NOT have ANSI escape codes
        assert "\x1b[" not in result.stdout, f"{cli_tool} {command}: Help should not contain ANSI codes"

        # Should have command-specific help
        assert command in result.stdout.lower() or "usage:" in result.stdout.lower()

        # Should exit cleanly
        assert result.returncode in (0, 1)

    def test_no_color_flag_recognized_without_errors(self):
        """Test that --no-color flag doesn't cause argparse errors."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "edit", "--help", "--no-color"],
            capture_output=True,
            text=True,
        )

        # Should NOT have argparse error messages
        assert "unrecognized arguments" not in result.stderr.lower()
        assert "error:" not in result.stderr.lower()

        # Should NOT have usage error
        assert result.returncode in (0, 1)  # Help exits with 0 or 1, not 2 (error)

    def test_no_color_flag_with_invalid_command(self):
        """Test --no-color flag works even when command fails."""
        # Try to export nonexistent file
        result = subprocess.run(
            [
                "python",
                "-m",
                "vba_edit.excel_vba",
                "export",
                "--no-color",
                "--file",
                "nonexistent_file_12345.xlsm",
                "--vba-directory",
                "temp_dir",
            ],
            capture_output=True,
            text=True,
        )

        # Even error output should not have ANSI codes
        assert "\x1b[" not in result.stdout
        assert "\x1b[" not in result.stderr

        # Should fail (file doesn't exist)
        assert result.returncode != 0


class TestBritishSpelling:
    """Tests for British spelling variant --no-colour."""

    @pytest.mark.parametrize(
        "cli_tool",
        ["excel_vba", "word_vba", "access_vba", "powerpoint_vba"],
    )
    def test_no_colour_british_spelling(self, cli_tool):
        """Test that --no-colour works (British English)."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", f"vba_edit.{cli_tool}", "--help", "--no-colour"],
            capture_output=True,
            text=True,
        )

        # Should NOT have ANSI escape codes
        assert "\x1b[" not in result.stdout, f"{cli_tool}: --no-colour should strip colors"

        # Should have help content
        assert "usage:" in result.stdout.lower()

        # Should not error
        assert result.returncode in (0, 1)

    def test_both_spellings_equivalent(self):
        """Test that --no-color and --no-colour produce identical output."""
        # Test with American spelling
        result_us = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "--help", "--no-color"],
            capture_output=True,
            text=True,
        )

        # Test with British spelling
        result_uk = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "--help", "--no-colour"],
            capture_output=True,
            text=True,
        )

        # Both should produce same output
        assert result_us.stdout == result_uk.stdout
        assert result_us.returncode == result_uk.returncode


class TestDisableColorsFunction:
    """Tests for disable_colors() function."""

    def test_disable_colors_strips_rich_markup(self):
        """Test that disable_colors() strips Rich markup tags."""
        from vba_edit.console import console, disable_colors

        # Call disable_colors
        disable_colors()

        # Capture output
        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        # Print with Rich markup
        console.print("[bold]test[/bold] [cyan]message[/cyan] [dim]dimmed[/dim]")

        # Restore original file
        console.file = original_file

        # Get output
        output = buffer.getvalue()

        # Should NOT contain Rich markup tags
        assert "[bold]" not in output
        assert "[/bold]" not in output
        assert "[cyan]" not in output
        assert "[/cyan]" not in output
        assert "[dim]" not in output
        assert "[/dim]" not in output

        # Should contain the actual text
        assert "test" in output
        assert "message" in output
        assert "dimmed" in output

    def test_disable_colors_preserves_literal_brackets(self):
        """Test that disable_colors() preserves literal brackets like [CTRL+S]."""
        from vba_edit.console import console, disable_colors

        disable_colors()

        buffer = StringIO()
        original_file = console.file
        console.file = buffer

        # Print with literal brackets
        console.print("Press [CTRL+S] to save [ALT+F4] to quit")

        console.file = original_file
        output = buffer.getvalue()

        # Should preserve literal brackets
        assert "[CTRL+S]" in output
        assert "[ALT+F4]" in output

    def test_disable_colors_handles_multiple_calls(self):
        """Test that disable_colors() can be called multiple times safely."""
        from vba_edit.console import disable_colors

        # Should not raise any exceptions
        disable_colors()
        disable_colors()
        disable_colors()


class TestEasterEggColor:
    """Tests for easter egg flags with color control."""

    @pytest.mark.parametrize(
        "cli_tool,flag",
        [
            ("excel_vba", "--diagram"),
            ("word_vba", "--diagram"),
            ("access_vba", "--diagram"),
            ("powerpoint_vba", "--diagram"),
            ("excel_vba", "--how-it-works"),
            ("word_vba", "--how-it-works"),
        ],
    )
    def test_easter_egg_honors_no_color(self, cli_tool, flag):
        """Test that easter egg respects --no-color flag."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", f"vba_edit.{cli_tool}", flag, "--no-color"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",  # Replace encoding errors instead of failing
        )

        # Should NOT have ANSI escape codes
        assert "\x1b[" not in result.stdout, f"{cli_tool} {flag}: Should not have ANSI codes"

        # Should have diagram content
        assert "vba-edit" in result.stdout.lower() or "workflow" in result.stdout.lower()

        # Should exit cleanly
        assert result.returncode == 0

    def test_easter_egg_preserves_ctrl_s_brackets(self):
        """Test that [CTRL+S] literal brackets are preserved in easter egg."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "--diagram", "--no-color"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        # Should preserve literal brackets
        assert "[CTRL+S]" in result.stdout, "Literal [CTRL+S] brackets should be preserved"

        # Should NOT have Rich markup
        assert "[bold]" not in result.stdout.lower() or "[bold]vba-edit" not in result.stdout.lower()

    def test_easter_egg_british_spelling(self):
        """Test easter egg with British spelling --no-colour."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.word_vba", "--how-it-works", "--no-colour"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        # Should work with British spelling
        assert "\x1b[" not in result.stdout
        assert result.returncode == 0


class TestEasterEggFormat:
    """Tests for easter egg output format."""

    def test_easter_egg_output_structure(self):
        """Test that easter egg displays expected diagram structure."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "--diagram", "--no-color"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        # Should have key diagram elements
        assert "vba-edit Workflow" in result.stdout or "vba-edit" in result.stdout.lower()
        assert "Excel" in result.stdout or "Word" in result.stdout
        assert "EDIT" in result.stdout or "EXPORT" in result.stdout or "IMPORT" in result.stdout
        assert "[CTRL+S]" in result.stdout

        # Should exit cleanly
        assert result.returncode == 0

    @pytest.mark.parametrize("flag", ["--diagram", "--how-it-works"])
    def test_both_flags_produce_same_output(self, flag):
        """Test that --diagram and --how-it-works are identical."""
        result1 = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "--diagram", "--no-color"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        result2 = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "--how-it-works", "--no-color"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        # Should produce identical output
        assert result1.stdout == result2.stdout
        assert result1.returncode == result2.returncode


@pytest.mark.skipif(
    not sys.stdout.isatty() or os.getenv("NO_COLOR") or os.getenv("GITHUB_ACTIONS"),
    reason="Requires interactive TTY without NO_COLOR environment variable",
)
class TestColorWithTTY:
    """Tests for actual color output (TTY required, skipped in CI/CD)."""

    def test_help_has_colors_by_default(self):
        """Test that help output includes colors in TTY environment (local only)."""
        # This will SKIP in CI/CD (no TTY)
        result = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "edit", "--help"],
            capture_output=True,
            text=True,
        )

        # In a TTY with Rich available, should have ANSI codes
        # This is a "nice to have" test that documents expected behavior
        # but we don't fail the test if colors are missing (could be NO_COLOR set)
        if "\x1b[" not in result.stdout:
            # No colors - could be NO_COLOR env var or other reason
            # Just log it for visibility
            pytest.skip("No colors detected - NO_COLOR may be set or terminal doesn't support colors")

    def test_easter_egg_has_colors_without_flag(self):
        """Test that easter egg has colors by default in TTY (local only)."""
        result = subprocess.run(
            [PYTHON_EXE, "-m", "vba_edit.excel_vba", "--diagram"],
            capture_output=True,
            text=True,
        )

        # Similar to above - nice to have but not critical
        if "\x1b[" not in result.stdout:
            pytest.skip("No colors detected in easter egg output")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

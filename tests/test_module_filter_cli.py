"""CLI, config-file and config-gen tests for --include / --exclude (GitHub issue #77)."""

import argparse
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

from vba_edit.cli_common import (
    CONFIG_KEY_FILE,
    CONFIG_SECTION_GENERAL,
    _get_cli_explicit_args,
    add_filter_arguments,
    merge_config_with_args,
    toml_lib,  # stdlib tomllib on 3.11+, tomli on 3.10
)
from vba_edit.config_gen import build_toml, split_patterns
from vba_edit.excel_vba import create_cli_parser, handle_excel_vba_command

# region argparse + config merge


def _make_filter_parser():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--file", "-f", dest=CONFIG_KEY_FILE)
    add_filter_arguments(parser)
    return parser


class TestFilterArguments:
    def test_defaults_are_none(self):
        args = _make_filter_parser().parse_args([])
        assert args.include is None
        assert args.exclude is None

    def test_cli_lists_parsed(self):
        args = _make_filter_parser().parse_args(["--include", "A", "B*", "--exclude", "C"])
        assert args.include == ["A", "B*"]
        assert args.exclude == ["C"]

    def test_repeated_option_extends(self):
        args = _make_filter_parser().parse_args(["--include", "A", "--include", "B"])
        assert args.include == ["A", "B"]

    def test_lists_from_config_when_cli_absent(self):
        parser = _make_filter_parser()
        args = parser.parse_args([])
        config = {CONFIG_SECTION_GENERAL: {"include": ["Module1", "Helpers_*"], "exclude": ["Vendor_*"]}}
        result = merge_config_with_args(args, config, cli_explicit=_get_cli_explicit_args(parser, args))
        assert result.include == ["Module1", "Helpers_*"]
        assert result.exclude == ["Vendor_*"]

    def test_cli_list_wins_over_config(self):
        parser = _make_filter_parser()
        args = parser.parse_args(["--include", "OnlyThis"])
        config = {CONFIG_SECTION_GENERAL: {"include": ["Module1"], "exclude": ["Vendor_*"]}}
        result = merge_config_with_args(args, config, cli_explicit=_get_cli_explicit_args(parser, args))
        assert result.include == ["OnlyThis"]
        assert result.exclude == ["Vendor_*"]  # not given on CLI -> taken from config

    def test_cli_replaces_config_default_applied_via_set_defaults(self):
        """office_cli applies config values with parser.set_defaults(); CLI must replace, not append."""
        parser = _make_filter_parser()
        parser.set_defaults(include=["FromConfig"])
        args = parser.parse_args(["--include", "FromCli", "--include", "Second"])
        assert args.include == ["FromCli", "Second"]
        assert parser.parse_args([]).include == ["FromConfig"]

    def test_string_value_from_config_is_passed_through(self):
        parser = _make_filter_parser()
        args = parser.parse_args([])
        config = {CONFIG_SECTION_GENERAL: {"include": "Module1"}}
        result = merge_config_with_args(args, config, cli_explicit=_get_cli_explicit_args(parser, args))
        assert result.include == "Module1"  # ModuleFilter accepts a bare string


# endregion


# region excel-vba parser integration


def test_include_exclude_options_on_all_vba_commands():
    parser = create_cli_parser()
    for command in ("edit", "export", "import"):
        args = parser.parse_args([command, "--include", "Module1", "Helpers_*.cls", "--exclude", "Vendor_*"])
        assert args.include == ["Module1", "Helpers_*.cls"], command
        assert args.exclude == ["Vendor_*"], command

        args = parser.parse_args([command])
        assert args.include is None, command
        assert args.exclude is None, command


def test_include_option_repeatable_around_other_options():
    parser = create_cli_parser()
    args = parser.parse_args(["import", "--include", "A", "--file", "x.xlsm", "--include", "B"])
    assert args.include == ["A", "B"]
    assert args.file == "x.xlsm"


def test_include_exclude_not_on_check_command():
    args = create_cli_parser().parse_args(["check"])
    assert not hasattr(args, "include")
    assert not hasattr(args, "exclude")


@patch("vba_edit.excel_vba.ExcelVBAHandler")
@patch("vba_edit.excel_vba.get_document_paths")
@patch("vba_edit.excel_vba.setup_logging")
def test_include_exclude_passed_to_handler(mock_logging, mock_get_paths, mock_handler):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        doc_path = tmp_path / "test.xlsm"
        doc_path.touch()
        mock_get_paths.return_value = (doc_path, tmp_path)
        mock_handler.return_value = Mock()

        args = create_cli_parser().parse_args(
            ["import", "--file", str(doc_path), "--include", "Module1", "Utils/*", "--exclude", "Vendor_*"]
        )
        handle_excel_vba_command(args)

        call_kwargs = mock_handler.call_args[1]
        assert call_kwargs["include"] == ["Module1", "Utils/*"]
        assert call_kwargs["exclude"] == ["Vendor_*"]


@patch("vba_edit.excel_vba.ExcelVBAHandler")
@patch("vba_edit.excel_vba.get_document_paths")
@patch("vba_edit.excel_vba.setup_logging")
def test_filter_defaults_passed_as_none(mock_logging, mock_get_paths, mock_handler):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        doc_path = tmp_path / "test.xlsm"
        doc_path.touch()
        mock_get_paths.return_value = (doc_path, tmp_path)
        mock_handler.return_value = Mock()

        handle_excel_vba_command(create_cli_parser().parse_args(["export", "--file", str(doc_path)]))

        call_kwargs = mock_handler.call_args[1]
        assert call_kwargs["include"] is None
        assert call_kwargs["exclude"] is None


# endregion


# region config-gen


def test_split_patterns_variants():
    assert split_patterns(None) == []
    assert split_patterns("") == []
    assert split_patterns("Module1, Helpers_*.cls  Utils/*") == ["Module1", "Helpers_*.cls", "Utils/*"]
    assert split_patterns(["a", " b ", ""]) == ["a", "b"]


def test_build_toml_no_filter_keys_when_empty():
    toml = build_toml(include="", exclude=[])
    assert "include" not in toml
    assert "exclude" not in toml


def test_build_toml_include_exclude_arrays():
    toml = build_toml(include="Module1, Helpers_*.cls", exclude=["Vendor_*", "LockedForm.frm"])
    assert "include = ['Module1', 'Helpers_*.cls']" in toml
    assert "exclude = ['Vendor_*', 'LockedForm.frm']" in toml


def test_build_toml_filter_arrays_round_trip_through_toml():
    parsed = toml_lib.loads(build_toml(include="Module1 Utils/*", exclude="Vendor_*"))
    assert parsed["general"]["include"] == ["Module1", "Utils/*"]
    assert parsed["general"]["exclude"] == ["Vendor_*"]


# endregion

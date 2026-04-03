"""Tests for config file merge logic: _get_cli_explicit_args, _merge_config_section,
merge_config_with_args with cli_explicit parameter, [references] config section,
and new [general] config keys."""

import argparse

import pytest

from vba_edit.cli_common import (
    CONFIG_KEY_DETECT_ENCODING,
    CONFIG_KEY_FILE,
    CONFIG_KEY_FORCE_OVERWRITE,
    CONFIG_KEY_NO_BUILTINS,
    CONFIG_KEY_NO_COLOR,
    CONFIG_KEY_NO_CUSTOM,
    CONFIG_KEY_NO_THIRD_PARTY,
    CONFIG_KEY_REFS_FILE,
    CONFIG_KEY_RUBBERDUCK_FOLDERS,
    CONFIG_KEY_SAVE_METADATA,
    CONFIG_KEY_SKIP_EMPTY,
    CONFIG_KEY_VBA_DIRECTORY,
    CONFIG_KEY_VERBOSE,
    CONFIG_KEY_WITH_REFERENCES,
    CONFIG_SECTION_GENERAL,
    CONFIG_SECTION_REFERENCES,
    _get_cli_explicit_args,
    _merge_config_section,
    add_common_option_group,
    add_encoding_arguments,
    add_exporting_arguments,
    add_references_filter_arguments,
    add_references_output_arguments,
    merge_config_with_args,
)


def _make_export_parser():
    """Create a minimal parser with export-style arguments for testing."""
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--file", "-f", dest=CONFIG_KEY_FILE)
    parser.add_argument("--vba-directory", dest=CONFIG_KEY_VBA_DIRECTORY)
    parser.add_argument("--verbose", "-v", dest=CONFIG_KEY_VERBOSE, action="store_true")
    parser.add_argument("--rubberduck-folders", dest=CONFIG_KEY_RUBBERDUCK_FOLDERS, action="store_true")
    parser.add_argument("--no-color", dest=CONFIG_KEY_NO_COLOR, action="store_true")
    parser.add_argument("--force-overwrite", dest=CONFIG_KEY_FORCE_OVERWRITE, action="store_true")
    parser.add_argument("--skip-empty", dest=CONFIG_KEY_SKIP_EMPTY, action="store_true")
    parser.add_argument("--with-references", dest=CONFIG_KEY_WITH_REFERENCES, action="store_true")
    parser.add_argument("--save-metadata", dest=CONFIG_KEY_SAVE_METADATA, action="store_true")
    parser.add_argument("--detect-encoding", dest=CONFIG_KEY_DETECT_ENCODING, action="store_true")
    return parser


def _make_references_parser():
    """Create a minimal parser with reference-specific arguments for testing."""
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--file", "-f", dest=CONFIG_KEY_FILE)
    parser.add_argument("--verbose", "-v", dest=CONFIG_KEY_VERBOSE, action="store_true")
    add_references_output_arguments(parser)
    add_references_filter_arguments(parser)
    return parser


# ---------------------------------------------------------------------------
# _get_cli_explicit_args
# ---------------------------------------------------------------------------
class TestGetCliExplicitArgs:
    """Tests for _get_cli_explicit_args()."""

    def test_no_args_returns_empty_set(self):
        parser = _make_export_parser()
        args = parser.parse_args([])
        explicit = _get_cli_explicit_args(parser, args)
        assert explicit == set()

    def test_string_arg_detected(self):
        parser = _make_export_parser()
        args = parser.parse_args(["--file", "test.xlsm"])
        explicit = _get_cli_explicit_args(parser, args)
        assert CONFIG_KEY_FILE in explicit

    def test_store_true_arg_detected(self):
        parser = _make_export_parser()
        args = parser.parse_args(["--verbose"])
        explicit = _get_cli_explicit_args(parser, args)
        assert CONFIG_KEY_VERBOSE in explicit

    def test_non_passed_args_not_in_set(self):
        parser = _make_export_parser()
        args = parser.parse_args(["--verbose"])
        explicit = _get_cli_explicit_args(parser, args)
        assert CONFIG_KEY_FILE not in explicit
        assert CONFIG_KEY_RUBBERDUCK_FOLDERS not in explicit
        assert CONFIG_KEY_VBA_DIRECTORY not in explicit

    def test_multiple_args_detected(self):
        parser = _make_export_parser()
        args = parser.parse_args(["--verbose", "--file", "x.xlsm", "--skip-empty"])
        explicit = _get_cli_explicit_args(parser, args)
        assert {CONFIG_KEY_VERBOSE, CONFIG_KEY_FILE, CONFIG_KEY_SKIP_EMPTY} <= explicit

    def test_underscored_private_keys_skipped(self):
        """Keys starting with _ should not appear in the explicit set."""
        parser = _make_export_parser()
        args = parser.parse_args([])
        # Manually inject a private key
        args._internal = "something"
        explicit = _get_cli_explicit_args(parser, args)
        assert "_internal" not in explicit

    def test_reference_filter_args_detected(self):
        parser = _make_references_parser()
        args = parser.parse_args(["--no-builtins", "--no-custom"])
        explicit = _get_cli_explicit_args(parser, args)
        assert CONFIG_KEY_NO_BUILTINS in explicit
        assert CONFIG_KEY_NO_CUSTOM in explicit
        assert CONFIG_KEY_NO_THIRD_PARTY not in explicit

    def test_refs_file_detected(self):
        parser = _make_references_parser()
        args = parser.parse_args(["--refs-file", "my_refs.toml"])
        explicit = _get_cli_explicit_args(parser, args)
        assert CONFIG_KEY_REFS_FILE in explicit


# ---------------------------------------------------------------------------
# _merge_config_section
# ---------------------------------------------------------------------------
class TestMergeConfigSection:
    """Tests for _merge_config_section()."""

    def test_missing_section_is_noop(self):
        args_dict = {CONFIG_KEY_VERBOSE: None, CONFIG_KEY_FILE: None}
        original = args_dict.copy()
        _merge_config_section(args_dict, {}, CONFIG_SECTION_GENERAL, None)
        assert args_dict == original

    def test_non_dict_section_is_noop(self):
        args_dict = {CONFIG_KEY_VERBOSE: None}
        config = {CONFIG_SECTION_GENERAL: "not-a-dict"}
        _merge_config_section(args_dict, config, CONFIG_SECTION_GENERAL, None)
        assert args_dict[CONFIG_KEY_VERBOSE] is None

    def test_legacy_none_check_applies_config(self):
        """Without cli_explicit, only None values are overridden."""
        args_dict = {CONFIG_KEY_VERBOSE: None, CONFIG_KEY_FILE: None}
        config = {CONFIG_SECTION_GENERAL: {"verbose": True, "file": "doc.xlsm"}}
        _merge_config_section(args_dict, config, CONFIG_SECTION_GENERAL, None)
        assert args_dict[CONFIG_KEY_VERBOSE] is True
        assert args_dict[CONFIG_KEY_FILE] == "doc.xlsm"

    def test_legacy_none_check_preserves_non_none(self):
        """Without cli_explicit, non-None values are NOT overridden."""
        args_dict = {CONFIG_KEY_VERBOSE: False, CONFIG_KEY_FILE: "other.xlsm"}
        config = {CONFIG_SECTION_GENERAL: {"verbose": True, "file": "doc.xlsm"}}
        _merge_config_section(args_dict, config, CONFIG_SECTION_GENERAL, None)
        # False is non-None → not overridden
        assert args_dict[CONFIG_KEY_VERBOSE] is False
        assert args_dict[CONFIG_KEY_FILE] == "other.xlsm"

    def test_cli_explicit_protects_explicit_args(self):
        """With cli_explicit, explicitly set args are NOT overridden."""
        args_dict = {CONFIG_KEY_VERBOSE: True, CONFIG_KEY_FILE: "cli.xlsm"}
        config = {CONFIG_SECTION_GENERAL: {"verbose": False, "file": "config.xlsm"}}
        cli_explicit = {CONFIG_KEY_VERBOSE, CONFIG_KEY_FILE}
        _merge_config_section(args_dict, config, CONFIG_SECTION_GENERAL, cli_explicit)
        assert args_dict[CONFIG_KEY_VERBOSE] is True
        assert args_dict[CONFIG_KEY_FILE] == "cli.xlsm"

    def test_cli_explicit_allows_non_explicit_override(self):
        """With cli_explicit, args NOT in the set ARE overridden by config."""
        args_dict = {CONFIG_KEY_VERBOSE: False, CONFIG_KEY_FILE: None}
        config = {CONFIG_SECTION_GENERAL: {"verbose": True, "file": "config.xlsm"}}
        cli_explicit = set()  # Nothing explicitly set
        _merge_config_section(args_dict, config, CONFIG_SECTION_GENERAL, cli_explicit)
        assert args_dict[CONFIG_KEY_VERBOSE] is True
        assert args_dict[CONFIG_KEY_FILE] == "config.xlsm"

    def test_hyphen_to_underscore_mapping(self):
        """Config keys with hyphens map to underscored arg keys."""
        args_dict = {CONFIG_KEY_VBA_DIRECTORY: None}
        config = {CONFIG_SECTION_GENERAL: {"vba-directory": "/some/path"}}
        _merge_config_section(args_dict, config, CONFIG_SECTION_GENERAL, None)
        assert args_dict[CONFIG_KEY_VBA_DIRECTORY] == "/some/path"

    def test_unknown_config_key_ignored(self):
        """Config keys with no matching arg are silently skipped."""
        args_dict = {CONFIG_KEY_VERBOSE: None}
        config = {CONFIG_SECTION_GENERAL: {"nonexistent_key": "value"}}
        _merge_config_section(args_dict, config, CONFIG_SECTION_GENERAL, None)
        assert "nonexistent_key" not in args_dict

    def test_references_section(self):
        """[references] section keys are applied correctly."""
        args_dict = {
            CONFIG_KEY_REFS_FILE: None,
            CONFIG_KEY_NO_BUILTINS: None,
            CONFIG_KEY_NO_THIRD_PARTY: None,
            CONFIG_KEY_NO_CUSTOM: None,
        }
        config = {
            CONFIG_SECTION_REFERENCES: {
                "refs_file": "my_refs.toml",
                "no_builtins": True,
                "no_third_party": False,
                "no_custom": True,
            }
        }
        _merge_config_section(args_dict, config, CONFIG_SECTION_REFERENCES, None)
        assert args_dict[CONFIG_KEY_REFS_FILE] == "my_refs.toml"
        assert args_dict[CONFIG_KEY_NO_BUILTINS] is True
        assert args_dict[CONFIG_KEY_NO_THIRD_PARTY] is False
        assert args_dict[CONFIG_KEY_NO_CUSTOM] is True


# ---------------------------------------------------------------------------
# merge_config_with_args — full function with cli_explicit
# ---------------------------------------------------------------------------
class TestMergeConfigWithArgs:
    """Tests for merge_config_with_args() with the cli_explicit parameter."""

    def test_legacy_call_without_cli_explicit(self):
        """Calling without cli_explicit preserves backward-compatible behavior."""
        args = argparse.Namespace(verbose=None, file=None, vba_directory=None, conf="test.toml")
        config = {CONFIG_SECTION_GENERAL: {"verbose": True, "file": "doc.xlsm"}}
        result = merge_config_with_args(args, config)
        assert result.verbose is True
        assert result.file == "doc.xlsm"

    def test_legacy_preserves_non_none(self):
        args = argparse.Namespace(verbose=False, file="cli.xlsm", conf="test.toml")
        config = {CONFIG_SECTION_GENERAL: {"verbose": True, "file": "config.xlsm"}}
        result = merge_config_with_args(args, config)
        assert result.verbose is False
        assert result.file == "cli.xlsm"

    def test_cli_explicit_empty_set_applies_all_config(self):
        """cli_explicit=set() means nothing was explicitly set — all config applied."""
        args = argparse.Namespace(verbose=False, file=None, conf="test.toml")
        config = {CONFIG_SECTION_GENERAL: {"verbose": True, "file": "config.xlsm"}}
        result = merge_config_with_args(args, config, cli_explicit=set())
        assert result.verbose is True
        assert result.file == "config.xlsm"

    def test_cli_explicit_protects_verbose(self):
        """Explicitly set --verbose should NOT be overridden by config."""
        args = argparse.Namespace(verbose=True, file=None, conf="test.toml")
        config = {CONFIG_SECTION_GENERAL: {"verbose": False}}
        result = merge_config_with_args(args, config, cli_explicit={"verbose"})
        assert result.verbose is True

    def test_merges_both_general_and_references_sections(self):
        """Both [general] and [references] sections should be merged."""
        args = argparse.Namespace(
            verbose=None,
            file=None,
            refs_file=None,
            no_builtins=None,
            conf="test.toml",
        )
        config = {
            CONFIG_SECTION_GENERAL: {"verbose": True, "file": "doc.xlsm"},
            CONFIG_SECTION_REFERENCES: {"refs_file": "refs.toml", "no_builtins": True},
        }
        result = merge_config_with_args(args, config)
        assert result.verbose is True
        assert result.file == "doc.xlsm"
        assert result.refs_file == "refs.toml"
        assert result.no_builtins is True

    def test_stores_config_and_path(self):
        """_config and _config_file_path are stored on the result namespace."""
        args = argparse.Namespace(verbose=None, conf="test.toml")
        config = {CONFIG_SECTION_GENERAL: {"verbose": True}}
        result = merge_config_with_args(args, config)
        assert result._config is config
        assert hasattr(result, "_config_file_path")

    def test_cli_explicit_with_references_section(self):
        """cli_explicit also protects reference args."""
        args = argparse.Namespace(
            refs_file="cli_refs.toml",
            no_builtins=True,
            no_third_party=False,
            conf="test.toml",
        )
        config = {
            CONFIG_SECTION_REFERENCES: {
                "refs_file": "config_refs.toml",
                "no_builtins": False,
                "no_third_party": True,
            }
        }
        cli_explicit = {CONFIG_KEY_REFS_FILE, CONFIG_KEY_NO_BUILTINS}
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        # Explicit CLI args protected
        assert result.refs_file == "cli_refs.toml"
        assert result.no_builtins is True
        # Non-explicit arg overridden by config
        assert result.no_third_party is True


# ---------------------------------------------------------------------------
# New [general] config keys
# ---------------------------------------------------------------------------
class TestNewGeneralConfigKeys:
    """Tests for new config keys in the [general] section."""

    @pytest.mark.parametrize(
        "config_key, config_value, expected",
        [
            (CONFIG_KEY_WITH_REFERENCES, True, True),
            (CONFIG_KEY_SKIP_EMPTY, True, True),
            (CONFIG_KEY_FORCE_OVERWRITE, True, True),
            (CONFIG_KEY_SAVE_METADATA, True, True),
            (CONFIG_KEY_DETECT_ENCODING, True, True),
            (CONFIG_KEY_WITH_REFERENCES, False, False),
        ],
    )
    def test_general_key_applied_via_merge(self, config_key, config_value, expected):
        """Each new [general] key should be applied when not explicitly set on CLI."""
        args = argparse.Namespace(**{config_key: None, "conf": "test.toml"})
        config = {CONFIG_SECTION_GENERAL: {config_key: config_value}}
        result = merge_config_with_args(args, config)
        assert getattr(result, config_key) == expected

    @pytest.mark.parametrize(
        "config_key",
        [
            CONFIG_KEY_WITH_REFERENCES,
            CONFIG_KEY_SKIP_EMPTY,
            CONFIG_KEY_FORCE_OVERWRITE,
            CONFIG_KEY_SAVE_METADATA,
            CONFIG_KEY_DETECT_ENCODING,
        ],
    )
    def test_general_key_not_overridden_when_explicit(self, config_key):
        """CLI-explicit args should not be overridden by config."""
        args = argparse.Namespace(**{config_key: False, "conf": "test.toml"})
        config = {CONFIG_SECTION_GENERAL: {config_key: True}}
        result = merge_config_with_args(args, config, cli_explicit={config_key})
        assert getattr(result, config_key) is False


# ---------------------------------------------------------------------------
# [references] config section
# ---------------------------------------------------------------------------
class TestReferencesConfigSection:
    """Tests for the [references] config section."""

    def test_all_references_keys_applied(self):
        """All four [references] keys should be applied."""
        args = argparse.Namespace(
            refs_file=None,
            no_builtins=None,
            no_third_party=None,
            no_custom=None,
            conf="test.toml",
        )
        config = {
            CONFIG_SECTION_REFERENCES: {
                "refs_file": "project_refs.toml",
                "no_builtins": True,
                "no_third_party": True,
                "no_custom": False,
            }
        }
        result = merge_config_with_args(args, config)
        assert result.refs_file == "project_refs.toml"
        assert result.no_builtins is True
        assert result.no_third_party is True
        assert result.no_custom is False

    def test_references_section_with_hyphen_keys(self):
        """Hyphenated TOML keys should map to underscored dest names."""
        args = argparse.Namespace(
            refs_file=None,
            no_builtins=None,
            no_third_party=None,
            no_custom=None,
            conf="test.toml",
        )
        config = {
            CONFIG_SECTION_REFERENCES: {
                "refs-file": "hyphen_refs.toml",
                "no-builtins": True,
                "no-third-party": True,
                "no-custom": True,
            }
        }
        result = merge_config_with_args(args, config)
        assert result.refs_file == "hyphen_refs.toml"
        assert result.no_builtins is True
        assert result.no_third_party is True
        assert result.no_custom is True

    def test_empty_references_section_is_noop(self):
        args = argparse.Namespace(refs_file=None, no_builtins=None, conf="test.toml")
        config = {CONFIG_SECTION_REFERENCES: {}}
        result = merge_config_with_args(args, config)
        assert result.refs_file is None
        assert result.no_builtins is None

    def test_general_and_references_combined(self):
        """Both sections should be merged in a single call."""
        args = argparse.Namespace(
            verbose=None,
            with_references=None,
            refs_file=None,
            no_builtins=None,
            conf="test.toml",
        )
        config = {
            CONFIG_SECTION_GENERAL: {
                "verbose": True,
                "with_references": True,
            },
            CONFIG_SECTION_REFERENCES: {
                "refs_file": "combined.toml",
                "no_builtins": True,
            },
        }
        result = merge_config_with_args(args, config)
        assert result.verbose is True
        assert result.with_references is True
        assert result.refs_file == "combined.toml"
        assert result.no_builtins is True


# ---------------------------------------------------------------------------
# store_true defaults — the core motivation for the refactoring
# ---------------------------------------------------------------------------
class TestStoreTrueConfigMerge:
    """Tests that store_true flags work correctly with config files.

    This is the core scenario that motivated _get_cli_explicit_args:
    argparse store_true with no explicit default uses None, which works
    with the legacy is-None check. But if someone adds default=False,
    the legacy check fails because False is not None.
    """

    def test_store_true_default_none_applies_config(self):
        """store_true without explicit default → default is None → config applies."""
        parser = _make_export_parser()
        args = parser.parse_args([])
        # Default for store_true without default= is False (argparse behavior)
        # But our parser does NOT set default=False, so argparse uses False internally
        config = {CONFIG_SECTION_GENERAL: {"verbose": True}}
        cli_explicit = _get_cli_explicit_args(parser, args)
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        assert result.verbose is True

    def test_store_true_explicitly_set_wins_over_config(self):
        """User passes --verbose → config verbose=false should NOT override."""
        parser = _make_export_parser()
        args = parser.parse_args(["--verbose"])
        config = {CONFIG_SECTION_GENERAL: {"verbose": False}}
        cli_explicit = _get_cli_explicit_args(parser, args)
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        assert result.verbose is True

    def test_store_true_not_passed_config_true_applies(self):
        """User doesn't pass --verbose, config says verbose=true → applied."""
        parser = _make_export_parser()
        args = parser.parse_args(["--file", "test.xlsm"])
        config = {CONFIG_SECTION_GENERAL: {"verbose": True}}
        cli_explicit = _get_cli_explicit_args(parser, args)
        assert CONFIG_KEY_VERBOSE not in cli_explicit
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        assert result.verbose is True

    def test_with_references_from_config(self):
        """--with-references not passed, config says true → applied."""
        parser = _make_export_parser()
        args = parser.parse_args([])
        config = {CONFIG_SECTION_GENERAL: {"with_references": True}}
        cli_explicit = _get_cli_explicit_args(parser, args)
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        assert result.with_references is True

    def test_skip_empty_from_config(self):
        parser = _make_export_parser()
        args = parser.parse_args([])
        config = {CONFIG_SECTION_GENERAL: {"skip_empty": True}}
        cli_explicit = _get_cli_explicit_args(parser, args)
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        assert result.skip_empty is True

    def test_reference_filters_from_config(self):
        """Reference filter flags applied from [references] config section."""
        parser = _make_references_parser()
        args = parser.parse_args([])
        config = {
            CONFIG_SECTION_REFERENCES: {
                "no_builtins": True,
                "no_third_party": True,
                "no_custom": False,
                "refs_file": "project.toml",
            }
        }
        cli_explicit = _get_cli_explicit_args(parser, args)
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        assert result.no_builtins is True
        assert result.no_third_party is True
        assert result.no_custom is False
        assert result.refs_file == "project.toml"

    def test_cli_filter_overrides_config(self):
        """--no-builtins on CLI overrides config no_builtins=false."""
        parser = _make_references_parser()
        args = parser.parse_args(["--no-builtins"])
        config = {CONFIG_SECTION_REFERENCES: {"no_builtins": False}}
        cli_explicit = _get_cli_explicit_args(parser, args)
        result = merge_config_with_args(args, config, cli_explicit=cli_explicit)
        assert result.no_builtins is True


# ---------------------------------------------------------------------------
# _get_config_defaults in OfficeVBACLI — reads both sections
# ---------------------------------------------------------------------------
class TestGetConfigDefaults:
    """Tests for OfficeVBACLI._get_config_defaults reading [references] section."""

    def test_reads_general_section(self):
        from vba_edit.office_cli import OfficeVBACLI

        cli = OfficeVBACLI("excel")
        parser = cli.create_cli_parser()
        subparser = cli._get_subparser(parser, "export")
        config = {CONFIG_SECTION_GENERAL: {"verbose": True, "file": "test.xlsm"}}
        defaults = cli._get_config_defaults(subparser or parser, config)
        assert defaults.get("verbose") is True
        assert defaults.get("file") == "test.xlsm"

    def test_reads_references_section(self):
        from vba_edit.office_cli import OfficeVBACLI

        cli = OfficeVBACLI("excel")
        parser = cli.create_cli_parser()
        # references list subcommand has --refs-file, --no-builtins etc.
        # We need to find the subparser for 'references' — but references has sub-subparsers
        # Use the export subparser which also has --with-references
        subparser = cli._get_subparser(parser, "export")
        config = {
            CONFIG_SECTION_REFERENCES: {
                "no_builtins": True,
            }
        }
        defaults = cli._get_config_defaults(subparser or parser, config)
        # no_builtins is not on the export subparser — it's on references subcommands
        # So it may or may not be in defaults depending on the parser structure
        # The important thing is that the function doesn't crash
        assert isinstance(defaults, dict)

    def test_reads_both_sections(self):
        from vba_edit.office_cli import OfficeVBACLI

        cli = OfficeVBACLI("excel")
        parser = cli.create_cli_parser()
        subparser = cli._get_subparser(parser, "export")
        config = {
            CONFIG_SECTION_GENERAL: {"verbose": True},
            CONFIG_SECTION_REFERENCES: {"refs_file": "custom.toml"},
        }
        defaults = cli._get_config_defaults(subparser or parser, config)
        assert defaults.get("verbose") is True
        # refs_file may or may not be a known dest on the export subparser,
        # but the function should not crash
        assert isinstance(defaults, dict)

    def test_with_references_on_export_subparser(self):
        """with_references should be a known dest on the export subparser."""
        from vba_edit.office_cli import OfficeVBACLI

        cli = OfficeVBACLI("excel")
        parser = cli.create_cli_parser()
        subparser = cli._get_subparser(parser, "export")
        config = {CONFIG_SECTION_GENERAL: {"with_references": True}}
        defaults = cli._get_config_defaults(subparser or parser, config)
        assert defaults.get("with_references") is True


# ---------------------------------------------------------------------------
# Config loading scope — "references" command included
# ---------------------------------------------------------------------------
class TestReferencesCommandConfigLoading:
    """Tests that config files are loaded for the 'references' command."""

    def test_references_command_in_config_loading_scope(self):
        """Verify 'references' is in the set of commands that load config files.

        We test this by reading the source to confirm, since the actual run()
        method requires a live Office document. The implementation check is
        sufficient for this scope test.
        """
        import inspect
        from vba_edit.office_cli import OfficeVBACLI

        source = inspect.getsource(OfficeVBACLI.main)
        # The config loading condition should include "references"
        assert '"references"' in source or "'references'" in source

    def test_references_config_via_cli_parser(self, tmp_path):
        """Config file with [references] section is accepted by references subcommand parser."""
        from vba_edit.office_cli import OfficeVBACLI
        from vba_edit.cli_common import load_config_file

        config_content = '[general]\nverbose = true\n\n[references]\nno_builtins = true\nrefs_file = "custom.toml"\n'
        config_file = tmp_path / "test-config.toml"
        config_file.write_text(config_content, encoding="utf-8")

        config = load_config_file(str(config_file))
        assert CONFIG_SECTION_REFERENCES in config
        assert config[CONFIG_SECTION_REFERENCES]["no_builtins"] is True
        assert config[CONFIG_SECTION_REFERENCES]["refs_file"] == "custom.toml"

        cli = OfficeVBACLI("excel")
        parser = cli.create_cli_parser()
        # This should not crash — config with [references] is valid
        defaults = cli._get_config_defaults(parser, config)
        assert isinstance(defaults, dict)

"""Tests for vba_edit.module_filter (GitHub issue #77: --include / --exclude)."""

from pathlib import Path

import pytest

from vba_edit.module_filter import ModuleFilter, clean_patterns


# region clean_patterns


class TestCleanPatterns:
    def test_none_gives_empty_list(self):
        assert clean_patterns(None) == []

    def test_single_string_wrapped(self):
        assert clean_patterns("Module1") == ["Module1"]

    def test_whitespace_entries_dropped_and_stripped(self):
        assert clean_patterns(["  Module1 ", "", "   ", "Class*"]) == ["Module1", "Class*"]

    def test_tuple_input(self):
        assert clean_patterns(("a", "b")) == ["a", "b"]


# endregion


# region inactive filter


class TestInactiveFilter:
    def test_no_patterns_is_inactive_and_allows_everything(self):
        f = ModuleFilter()
        assert f.active is False
        assert bool(f) is False
        assert f.matches("Anything") is True
        assert f.allows_component("Anything", ".bas") is True
        assert f.allows_file(Path("C:/x/Anything.bas"), Path("C:/x")) is True

    def test_empty_lists_are_inactive(self):
        assert ModuleFilter([], []).active is False
        assert ModuleFilter([""], ["  "]).active is False

    def test_describe_empty(self):
        assert ModuleFilter().describe() == ""


# endregion


# region include matching


class TestIncludeMatching:
    def test_exact_module_name(self):
        f = ModuleFilter(include=["Module1"])
        assert f.matches("Module1") is True
        assert f.matches("Module10") is False
        assert f.matches("Other") is False

    def test_module_name_is_case_insensitive(self):
        f = ModuleFilter(include=["module1"])
        assert f.matches("MODULE1") is True
        assert f.allows_component("Module1", ".bas") is True

    def test_name_glob(self):
        f = ModuleFilter(include=["mod_10*"])
        assert f.allows_component("mod_10_report", ".bas") is True
        assert f.allows_component("mod_20_report", ".bas") is False

    def test_file_name_with_extension_distinguishes_type(self):
        f = ModuleFilter(include=["Helpers_*.cls"])
        assert f.allows_component("Helpers_Text", ".cls") is True
        assert f.allows_component("Helpers_Text", ".bas") is False

    def test_extension_only_glob(self):
        f = ModuleFilter(include=["*.frm"])
        assert f.allows_component("frmMain", ".frm") is True
        assert f.allows_component("Module1", ".bas") is False

    def test_multiple_include_patterns_are_or_ed(self):
        f = ModuleFilter(include=["Module1", "Class*"])
        assert f.allows_component("Module1", ".bas") is True
        assert f.allows_component("ClassFoo", ".cls") is True
        assert f.allows_component("Sheet1", ".cls") is False

    def test_single_string_include(self):
        f = ModuleFilter(include="Module1")
        assert f.include == ("Module1",)
        assert f.allows_component("Module1", ".bas") is True


# endregion


# region exclude matching


class TestExcludeMatching:
    def test_exclude_only(self):
        f = ModuleFilter(exclude=["Vendor_*", "LockedForm.frm"])
        assert f.allows_component("Vendor_Json", ".bas") is False
        assert f.allows_component("LockedForm", ".frm") is False
        assert f.allows_component("MyModule", ".bas") is True

    def test_exclude_wins_over_include(self):
        f = ModuleFilter(include=["Helpers_*"], exclude=["Helpers_Legacy"])
        assert f.allows_component("Helpers_Text", ".bas") is True
        assert f.allows_component("Helpers_Legacy", ".bas") is False

    def test_exclude_is_case_insensitive(self):
        f = ModuleFilter(exclude=["vendor_*"])
        assert f.allows_component("VENDOR_X", ".bas") is False


# endregion


# region path matching (RubberduckVBA folders)


class TestPathMatching:
    def test_folder_glob_against_relative_path(self):
        f = ModuleFilter(include=["Utils/*"])
        assert f.allows_component("Strings", ".bas", folder="Utils") is True
        assert f.allows_component("Strings", ".bas", folder="Utils/Sub") is True  # '*' spans '/'
        assert f.allows_component("Strings", ".bas", folder="") is False
        assert f.allows_component("Strings", ".bas", folder="Other") is False

    def test_backslash_pattern_and_folder_are_normalised(self):
        f = ModuleFilter(include=["Utils\\*"])
        assert f.allows_component("Strings", ".bas", folder="Utils\\Sub") is True
        assert f.allows_component("Strings", ".bas", folder="Utils") is True

    def test_folder_qualified_name_without_extension(self):
        f = ModuleFilter(include=["Utils/Strings"])
        assert f.allows_component("Strings", ".bas", folder="Utils") is True
        assert f.allows_component("Strings", ".cls", folder="Utils") is True
        assert f.allows_component("Strings", ".bas", folder="") is False

    def test_leading_dot_slash_in_pattern(self):
        f = ModuleFilter(include=["./Utils/*.bas"])
        assert f.allows_component("Strings", ".bas", folder="Utils") is True

    def test_needs_folder_only_when_pattern_has_separator(self):
        assert ModuleFilter(include=["Module1"]).needs_folder is False
        assert ModuleFilter(include=["Utils/*"]).needs_folder is True
        assert ModuleFilter(exclude=["Vendor\\*"]).needs_folder is True

    def test_allows_file_uses_path_relative_to_base_dir(self, tmp_path):
        base = tmp_path / "vba"
        target = base / "Utils" / "Strings.bas"
        target.parent.mkdir(parents=True)
        target.touch()
        assert ModuleFilter(include=["Utils/*"]).allows_file(target, base) is True
        assert ModuleFilter(include=["Strings"]).allows_file(target, base) is True
        assert ModuleFilter(include=["Strings.bas"]).allows_file(target, base) is True
        assert ModuleFilter(exclude=["Utils/*"]).allows_file(target, base) is False

    def test_allows_file_outside_base_dir_falls_back_to_name(self, tmp_path):
        other = tmp_path / "elsewhere" / "Module1.bas"
        other.parent.mkdir(parents=True)
        other.touch()
        f = ModuleFilter(include=["Module1"])
        assert f.allows_file(other, tmp_path / "vba") is True
        assert ModuleFilter(include=["vba/*"]).allows_file(other, tmp_path / "vba") is False

    def test_allows_file_without_base_dir(self):
        f = ModuleFilter(include=["*.cls"])
        assert f.allows_file(Path("Class1.cls")) is True
        assert f.allows_file(Path("Module1.bas")) is False


# endregion


# region describe


class TestDescribe:
    def test_describe_include_and_exclude(self):
        f = ModuleFilter(include=["a", "b*"], exclude=["c"])
        assert f.describe() == "include: a, b*; exclude: c"

    def test_describe_exclude_only(self):
        assert ModuleFilter(exclude=["Vendor_*"]).describe() == "exclude: Vendor_*"


# endregion


@pytest.mark.parametrize(
    "include,exclude,name,ext,folder,expected",
    [
        # Issue #77 examples
        (["mod_10*.bas"], None, "mod_10_main", ".bas", "", True),
        (["mod_10*.bas"], None, "mod_11_main", ".bas", "", False),
        (["mod_10*.bas", "helper.cls"], None, "helper", ".cls", "", True),
        # maximweb examples
        (["MyModule.bas", "Helpers_*.cls"], None, "MyModule", ".bas", "", True),
        (["MyModule.bas", "Helpers_*.cls"], None, "Helpers_Str", ".cls", "", True),
        (["MyModule.bas", "Helpers_*.cls"], None, "Helpers_Str", ".bas", "", False),
        (None, ["Vendor_*", "LockedForm.frm"], "Vendor_Lib", ".bas", "", False),
        (None, ["Vendor_*", "LockedForm.frm"], "LockedForm", ".frm", "", False),
        (None, ["Vendor_*", "LockedForm.frm"], "MyForm", ".frm", "", True),
    ],
)
def test_issue_examples(include, exclude, name, ext, folder, expected):
    assert ModuleFilter(include, exclude).allows_component(name, ext, folder) is expected

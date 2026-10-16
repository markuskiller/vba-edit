"""Include/exclude filtering of VBA components and source files.

Implements the module selection feature requested in GitHub issue #77:
``--include`` / ``--exclude`` patterns on ``edit``, ``export`` and ``import``
restrict which VBA modules are touched.

A pattern is a shell-style glob (``fnmatch``) and is tested, case-insensitively,
against three views of a component:

* the bare module name              ``Module1``
* the source file name              ``Module1.bas``
* the path relative to the VBA dir  ``Utils/Module1.bas`` (RubberduckVBA folders)

Backslashes in patterns and paths are normalised to forward slashes, so
``Utils\\*`` and ``Utils/*`` are equivalent.  Include patterns are applied
first; anything that survives is then checked against the exclude patterns.

UserForm binaries (``.frx``) and separate ``.header`` files are never
matched directly: they always follow the decision made for their ``.frm`` /
code file, because the handler only ever copies them as part of exporting or
importing that component.
"""

from __future__ import annotations

import fnmatch
from pathlib import Path
from typing import Iterable, Optional, Sequence, Union

PatternInput = Union[None, str, Iterable[str]]

__all__ = ["ModuleFilter", "clean_patterns"]


def clean_patterns(patterns: PatternInput) -> list[str]:
    """Normalise user-supplied pattern input to a list of non-empty strings.

    Accepts ``None``, a single string, or any iterable of strings (CLI
    ``nargs="+"`` lists, TOML arrays).  Whitespace-only entries are dropped.
    """
    if patterns is None:
        return []
    if isinstance(patterns, str):
        patterns = [patterns]
    cleaned: list[str] = []
    for raw in patterns:
        text = str(raw).strip()
        if text:
            cleaned.append(text)
    return cleaned


def _normalize(value: str) -> str:
    """Lower-case, forward-slash, no leading ``./`` or surrounding slashes."""
    text = value.strip().replace("\\", "/")
    while text.startswith("./"):
        text = text[2:]
    return text.strip("/").lower()


class ModuleFilter:
    """Decide whether a VBA component or source file should be processed.

    Args:
        include: Patterns a component must match (any of them) to be processed.
            Empty/None means "everything is included".
        exclude: Patterns that remove a component even if it was included.
    """

    def __init__(self, include: PatternInput = None, exclude: PatternInput = None) -> None:
        self.include: tuple[str, ...] = tuple(clean_patterns(include))
        self.exclude: tuple[str, ...] = tuple(clean_patterns(exclude))
        self._include_norm = tuple(_normalize(p) for p in self.include)
        self._exclude_norm = tuple(_normalize(p) for p in self.exclude)

    # ------------------------------------------------------------------ state

    @property
    def active(self) -> bool:
        """True when at least one include or exclude pattern is set."""
        return bool(self.include or self.exclude)

    @property
    def needs_folder(self) -> bool:
        """True when any pattern contains a path separator.

        Callers use this to decide whether it is worth computing the
        RubberduckVBA folder of a component (which costs a COM round-trip).
        """
        return any("/" in p for p in self._include_norm + self._exclude_norm)

    def describe(self) -> str:
        """Human readable summary for log output, e.g. ``include: a, b; exclude: c``."""
        parts = []
        if self.include:
            parts.append("include: " + ", ".join(self.include))
        if self.exclude:
            parts.append("exclude: " + ", ".join(self.exclude))
        return "; ".join(parts)

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"ModuleFilter(include={list(self.include)!r}, exclude={list(self.exclude)!r})"

    def __bool__(self) -> bool:
        return self.active

    # --------------------------------------------------------------- matching

    @staticmethod
    def _match_any(patterns: Sequence[str], candidates: Iterable[str]) -> bool:
        return any(fnmatch.fnmatchcase(candidate, pattern) for pattern in patterns for candidate in candidates)

    def matches(self, name: str, filename: Optional[str] = None, relative_path: Optional[str] = None) -> bool:
        """Return True if a component identified by these views passes the filter.

        Args:
            name: Bare module name (``Module1``).
            filename: Source file name including extension (``Module1.bas``).
            relative_path: Path relative to the VBA directory, any separator style.
        """
        if not self.active:
            return True

        candidates = {_normalize(name)}
        if filename:
            candidates.add(_normalize(filename))
        if relative_path:
            rel = _normalize(relative_path)
            candidates.add(rel)
            # Also allow matching the folder-qualified name without extension,
            # e.g. pattern "Utils/Module1" against "Utils/Module1.bas".
            stem, dot, _ext = rel.rpartition(".")
            if dot and "/" not in _ext:
                candidates.add(stem)

        if self._include_norm and not self._match_any(self._include_norm, candidates):
            return False
        if self._exclude_norm and self._match_any(self._exclude_norm, candidates):
            return False
        return True

    def allows_component(self, name: str, extension: str, folder: str = "") -> bool:
        """Filter decision for an in-document component (export side).

        Args:
            name: Component name.
            extension: Target file extension including the dot (``.bas``).
            folder: RubberduckVBA folder path (``Utils`` or ``Utils/Sub``), or ``""``.
        """
        filename = f"{name}{extension}"
        relative = f"{folder}/{filename}" if folder else filename
        return self.matches(name, filename, relative)

    def allows_file(self, path: Path, base_dir: Optional[Path] = None) -> bool:
        """Filter decision for a source file on disk (import / watch side).

        ``base_dir`` is the VBA directory; the relative path is derived from it
        when possible and falls back to the bare file name otherwise.
        """
        relative: str = path.name
        if base_dir is not None:
            relative = self._relative_to(path, base_dir)
        return self.matches(path.stem, path.name, relative)

    @staticmethod
    def _relative_to(path: Path, base_dir: Path) -> str:
        try:
            return path.relative_to(base_dir).as_posix()
        except ValueError:
            pass
        try:
            return path.resolve().relative_to(base_dir.resolve()).as_posix()
        except (ValueError, OSError):
            return path.name

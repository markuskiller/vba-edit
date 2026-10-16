"""Conservative two-way VBA editing. All COM calls stay on the watch-loop thread.

Only standard modules and Excel/Word document modules are writable. Hidden
attributes, structural changes and incomplete operations require explicit review.
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import tempfile
import time
import uuid
from contextlib import ExitStack
from dataclasses import asdict, dataclass, field, replace
from pathlib import Path

from watchfiles import watch

from vba_edit.exceptions import DocumentClosedError, RPCError, VBAError

logger = logging.getLogger(__name__)
STATE_VERSION = 1
INTERVAL_MS = 2500
SETTLED_SCANS = 2
EXTENSIONS = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}
MEMBER_ATTRIBUTE = re.compile(r"^\s*Attribute\s+[^=]*\.[^=]*=", re.IGNORECASE)


def normalize(text: str) -> str:
    """Ignore newline conventions and blank boundary lines, never attribute loss."""
    return text.replace("\r\n", "\n").replace("\r", "\n").strip("\n")


@dataclass(frozen=True)
class Snapshot:
    name: str
    kind: int
    path: str
    body: str
    header: str = ""
    attributes: tuple[str, ...] = ()
    raw: str = field(default="", compare=False)

    @classmethod
    def parse(cls, name: str, kind: int, path: str, raw: str) -> Snapshot:
        lines = normalize(raw).split("\n")
        boundary = 0
        in_block = False
        for index, line in enumerate(lines):
            stripped = line.strip()
            if stripped.upper().startswith("VERSION "):
                boundary = index + 1
            elif stripped.upper() == "BEGIN" or stripped.upper().startswith("BEGIN "):
                in_block = True
                boundary = index + 1
            elif in_block:
                boundary = index + 1
                if stripped.upper() == "END":
                    in_block = False
            elif re.match(r"Attribute\s+VB_", stripped, re.IGNORECASE):
                boundary = index + 1
            else:
                break
        body_lines = lines[boundary:]
        attributes = tuple(line for line in body_lines if MEMBER_ATTRIBUTE.match(line))
        body = normalize("\n".join(line for line in body_lines if not MEMBER_ATTRIBUTE.match(line)))
        return cls(name, kind, path, body, normalize("\n".join(lines[:boundary])), attributes, raw)

    @classmethod
    def load(cls, value: dict) -> Snapshot:
        return cls(**{**value, "attributes": tuple(value.get("attributes", []))})


def plan(baseline: Snapshot | None, office: Snapshot | None, files: Snapshot | None, *, startup=False) -> str:
    """Pure reconciliation decision. Read failures are handled before calling this."""
    if office is None or files is None:
        return "conflict"
    if startup and baseline is not None and office != baseline:
        return "conflict"
    if office == files:
        return "adopt"
    if baseline is None:
        return "conflict"
    if office == baseline:
        return "import"
    if files == baseline:
        return "export"
    return "conflict"


def atomic_json(path: Path, value: dict) -> None:
    atomic_bytes(path, (json.dumps(value, indent=2, ensure_ascii=False) + "\n").encode("utf-8"))


def atomic_bytes(path: Path, value: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # The suffix is deliberately not importable by Rubberduck rglob.
    temporary = path.with_name(f".{uuid.uuid4().hex}.bin")
    try:
        with temporary.open("xb") as stream:
            stream.write(value)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


class SessionLock:
    """OS lock, released even after a crash; the on-disk file is not authority."""

    def __init__(self, path: Path):
        self.path = path
        self.stream = None

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.stream = self.path.open("a+b")
        self.stream.seek(0, os.SEEK_END)
        if self.stream.tell() == 0:
            self.stream.write(b"\0")
            self.stream.flush()
        self.stream.seek(0)
        try:
            if os.name == "nt":
                import msvcrt

                msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(self.stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            self.stream.close()
            raise VBAError(f"Another edit session holds {self.path}") from exc
        return self

    def __exit__(self, *_):
        self.stream.close()


class OfficeAdapter:
    """Handler bridge; exports retain the file representation, scans read bodies."""

    def __init__(self, handler):
        self.handler = handler
        self.root = handler.vba_dir.resolve()
        self.state_dir = self.root / ".vba-edit"
        if self.state_dir.is_symlink() or not self.state_dir.resolve().is_relative_to(self.root):
            raise VBAError("Synchronization state directory must stay inside the source directory.")
        self.project_name = str(handler.get_vba_project().Name)
        self.excluded: set[str] = set()
        self.warned_exclusions: set[tuple[str, int]] = set()
        self.warned_filtered: set[str] = set()
        self.component_names: set[str] = set()

    def _passes_filter(self, component, name: str, kind: int) -> bool:
        """Apply the handler's --include/--exclude filter to an Office component."""
        h = self.handler
        module_filter = h.module_filter
        folder = ""
        if h.use_rubberduck_folders and module_filter.needs_folder:
            count = int(component.CodeModule.CountOfLines)
            code = str(component.CodeModule.Lines(1, count)) if count else ""
            folder, _ = h.component_handler.get_rubberduck_folder(code)
        return module_filter.allows_component(name, EXTENSIONS.get(kind, ".cls"), folder.replace(os.sep, "/"))

    def identity(self) -> dict:
        h = self.handler
        return {
            "document": str(h.doc_path.resolve()).casefold(),
            "host": h.app_name,
            "project": self.project_name,
            "encoding": h.encoding,
            "headers": "inline" if h.in_file_headers else "separate" if h.save_headers else "none",
            "folders": h.use_rubberduck_folders,
        }

    def inventory(self) -> dict[str, tuple[int, str]]:
        h = self.handler
        if not h.is_document_open():
            raise DocumentClosedError(h.document_type)
        project = h.get_vba_project()
        if str(project.Name) != self.project_name:
            raise VBAError("VBA project identity changed; restart after reviewing synchronization state.")
        if h.app_name == "Access" and Path(str(h.app.CurrentDb().Name)).resolve() != h.doc_path.resolve():
            raise VBAError("Access database identity changed.")
        # Mode values: run=0, break=1, design=2. Missing/failed reads are unknown.
        if int(project.Mode) != 2:
            raise VBAError("Office is in run/break mode; synchronization is suspended.")
        result = {}
        self.component_names = set()
        self.excluded = set()
        module_filter = getattr(h, "module_filter", None)
        for component in project.VBComponents:
            kind = int(component.Type)
            name = str(component.Name)
            self.component_names.add(name)
            if module_filter and not self._passes_filter(component, name, kind):
                if name not in self.warned_filtered:
                    logger.info("Excluded %s: does not match --include/--exclude", name)
                    self.warned_filtered.add(name)
                self.excluded.add(name)
                continue
            if kind != 1 and not (kind == 100 and h.app_name in {"Excel", "Word"}):
                if (name, kind) not in self.warned_exclusions:
                    logger.warning("Excluded %s: component type %s is not supported in bidirectional mode", name, kind)
                    self.warned_exclusions.add((name, kind))
                self.excluded.add(name)
                continue
            if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
                raise VBAError(f"Unsafe component name: {name!r}")
            count = int(component.CodeModule.CountOfLines)
            body = normalize(str(component.CodeModule.Lines(1, count))) if count else ""
            result[name] = (kind, body)
        return result

    def export(self, name: str, kind: int) -> Snapshot:
        h = self.handler
        component = h.get_vba_project().VBComponents(name)
        self.state_dir.mkdir(parents=True, exist_ok=True)
        temporary = self.state_dir / f"export-{uuid.uuid4().hex}.txt"
        try:
            component.Export(str(temporary))
            raw = temporary.read_text(encoding=h.encoding)
        finally:
            temporary.unlink(missing_ok=True)
        snapshot = Snapshot.parse(name, kind, name + EXTENSIONS[kind], raw)
        count = int(component.CodeModule.CountOfLines)
        visible = normalize(str(component.CodeModule.Lines(1, count))) if count else ""
        if snapshot.body != visible:
            raise VBAError("Export differs from live code (encoding loss or concurrent edit); refusing to synchronize.")
        if h.use_rubberduck_folders:
            folder, _ = h.component_handler.get_rubberduck_folder(snapshot.body)
            if folder:
                snapshot = replace(snapshot, path=(Path(folder) / snapshot.path).as_posix())
        self.path(snapshot.path)  # Validate before any file operation.
        return snapshot

    def path(self, relative: str) -> Path:
        path = (self.root / relative).resolve()
        if not path.is_relative_to(self.root) or path.is_relative_to(self.state_dir.resolve()):
            raise VBAError(f"Source path escapes the managed directory: {relative}")
        return path

    def files(self) -> dict[str, Snapshot]:
        h = self.handler
        result = {}
        paths = self.root.rglob("*") if h.use_rubberduck_folders else self.root.glob("*")
        for path in paths:
            if path.is_relative_to(self.state_dir) or path.suffix.lower() not in {".bas", ".cls", ".frm"}:
                continue
            if not path.is_file():
                continue
            self.path(path.relative_to(self.root).as_posix())
            name = path.stem
            if name in self.excluded:
                continue
            module_filter = getattr(h, "module_filter", None)
            if module_filter and not module_filter.allows_file(path, self.root):
                continue
            if name not in self.component_names and any(n.casefold() == name.casefold() for n in self.component_names):
                raise VBAError(f"Source filename case must match the Office component: {name}")
            if any(key.casefold() == name.casefold() for key in result):
                raise VBAError(f"Duplicate source component name: {name}")
            kind = 1 if path.suffix.lower() == ".bas" else 3 if path.suffix.lower() == ".frm" else 2
            raw = path.read_text(encoding=h.encoding)
            snap = Snapshot.parse(name, kind, path.relative_to(self.root).as_posix(), raw)
            if not snap.header and path.with_suffix(".header").exists():
                header_path = self.path(path.with_suffix(".header").relative_to(self.root).as_posix())
                header = header_path.read_text(encoding=h.encoding)
                snap = replace(snap, header=normalize(header))
            result[name] = snap
        return result

    def align(self, files: Snapshot | None, office: Snapshot | None) -> Snapshot | None:
        if files is None or office is None:
            return files
        # .cls alone cannot distinguish a host module from an ordinary class.
        if files.kind == 2 and office.kind == 100:
            files = replace(files, kind=100)
        if not files.header and not self.handler.save_headers and not self.handler.in_file_headers:
            files = replace(files, header=office.header)
        return files

    def write(self, snapshot: Snapshot) -> None:
        h = self.handler
        raw = snapshot.raw
        # Keep the original member attributes in the working file, never in Lines.
        _, code = h.component_handler.split_vba_content(raw)
        content = raw if h.in_file_headers else code + "\n"
        path = self.path(snapshot.path)
        if h.save_headers:
            atomic_bytes(path.with_suffix(".header"), (snapshot.header + "\n").encode(h.encoding))
        atomic_bytes(path, content.encode(h.encoding))

    def apply(self, snapshot: Snapshot, existing: Snapshot | None) -> None:
        h = self.handler
        if h.app_name == "Access":
            # New VBIDE modules can still need an interactive Access Save As.
            # Do not create them or claim they have been persisted automatically.
            if existing is None:
                raise VBAError(
                    "Save new standard modules in Access first; automatic Access module creation is excluded."
                )
            try:
                h.app.CurrentProject.AllModules.Item(snapshot.name)
            except Exception as exc:
                raise VBAError("Save this module in Access before importing source changes.") from exc
        if snapshot.attributes or (existing and existing.attributes):
            raise VBAError("Import would strip procedure/member attributes; this component requires review.")
        if existing and (snapshot.header != existing.header or snapshot.path != existing.path):
            raise VBAError("Header changes and component moves require review.")
        components = h.get_vba_project().VBComponents
        if existing is None:
            if snapshot.kind != 1 or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", snapshot.name):
                raise VBAError("Only unambiguous new standard modules are supported.")
            expected = f'Attribute VB_Name = "{snapshot.name}"'
            if snapshot.header and snapshot.header != expected:
                raise VBAError("New module has unsupported header attributes.")
            component = components.Add(1)
            component.Name = snapshot.name
        else:
            component = components(snapshot.name)
        if snapshot.kind == 100:
            h._update_document_module(snapshot.name, snapshot.body, components)
        elif snapshot.kind == 1:
            h._update_module_content(component, snapshot.body)
        else:
            raise VBAError("Only standard and Excel/Word document code is supported.")

    def save(self, name: str) -> None:
        if self.handler.app_name != "Access":
            self.handler.doc.Save()
        else:
            # acModule=5. Save the named existing module, not the active object.
            self.handler.app.DoCmd.Save(5, name)

    def metadata(self, baselines: dict[str, dict]) -> None:
        if getattr(self.handler, "sync_save_metadata", False):
            self.handler._save_metadata(
                {
                    name: {
                        "encoding": self.handler.encoding,
                        "type": "Standard Module" if value["kind"] == 1 else "Document Module",
                    }
                    for name, value in baselines.items()
                }
            )


class SyncSession:
    """Persisted baselines, conflict artifacts, and verified component operations."""

    def __init__(self, adapter: OfficeAdapter, *, resolve=None, component=None):
        self.adapter = adapter
        self.directory = adapter.state_dir
        self.state_path = self.directory / "state.json"
        self.resolve = resolve
        self.component = component
        self.state = {
            "version": STATE_VERSION,
            "identity": adapter.identity(),
            "baseline": {},
            "paused": {},
            "pending": {},
        }
        self.cache: dict[str, Snapshot] = {}
        self.settling: dict[str, tuple[object, int]] = {}
        self.read_failed = False

    def persist(self):
        atomic_json(self.state_path, self.state)

    def archive(self, name, reason, baseline, office, files, category="conflicts") -> str:
        component_id = hashlib.sha256(name.encode()).hexdigest()[:16]
        path = self.directory / category / component_id / uuid.uuid4().hex
        manifest = {
            "component": name,
            "reason": reason,
            "status": "paused" if category == "conflicts" else "pending",
            "snapshots": {},
        }
        for label, snapshot in (("baseline", baseline), ("office", office), ("filesystem", files)):
            atomic_json(path / f"{label}.json", asdict(snapshot) if snapshot else {})
            raw = snapshot.raw if snapshot else "<absent>"
            manifest["snapshots"][label] = {
                "path": snapshot.path if snapshot else None,
                "sha256": hashlib.sha256(raw.encode("utf-8")).hexdigest(),
            }
            atomic_bytes(path / f"{label}.txt", raw.encode("utf-8"))
        atomic_json(path / "manifest.json", manifest)
        return str(path)

    def pause(self, name, reason, baseline, office, files):
        if name not in self.state["paused"]:
            path = self.archive(name, reason, baseline, office, files)
            self.state["paused"][name] = {"reason": reason, "artifacts": path}
            self.persist()
        path = self.state["paused"][name]["artifacts"]
        logger.warning(
            'PAUSED %s: %s. Review %s; restart with --sync-resolve office|files --sync-component "%s" after review.',
            name,
            reason,
            path,
            name,
        )

    def adopt(self, snapshot):
        self.state["baseline"][snapshot.name] = asdict(snapshot)
        self.state["pending"].pop(snapshot.name, None)
        self.state["paused"].pop(snapshot.name, None)
        self.cache[snapshot.name] = snapshot
        self.persist()
        self.adapter.metadata(self.state["baseline"])

    def baseline(self, name):
        value = self.state["baseline"].get(name)
        return Snapshot.load(value) if value else None

    def startup(self):
        if self.state_path.exists():
            try:
                loaded = json.loads(self.state_path.read_text(encoding="utf-8"))
                if loaded["version"] != STATE_VERSION or loaded["identity"] != self.adapter.identity():
                    raise ValueError("identity/settings/schema mismatch")
                for key in ("baseline", "paused", "pending"):
                    if not isinstance(loaded[key], dict):
                        raise ValueError(f"invalid {key}")
                self.state = loaded
            except (KeyError, TypeError, ValueError) as exc:
                raise VBAError(
                    f"Cannot reuse {self.state_path}: {exc}. Preserve and review state before reinitializing."
                ) from exc
        inventory = self.adapter.inventory()
        files = self.adapter.files()
        fresh_empty = not self.state_path.exists() and not files
        # Baselines of components excluded by --include/--exclude are left untouched.
        names = (set(inventory) | set(files) | set(self.state["baseline"])) - self.adapter.excluded
        if self.component and self.component not in names:
            raise VBAError(f"Unknown component for resolution: {self.component}")
        for name in sorted(names):
            baseline = self.baseline(name)
            office = self.adapter.export(name, inventory[name][0]) if name in inventory else None
            files_snapshot = self.adapter.align(files.get(name), office or baseline)
            if office:
                self.cache[name] = office
            if self.resolve and self.component == name:
                self.archive(name, "explicit resolution", baseline, office, files_snapshot, "operations")
                if self.resolve == "office" and office is not None:
                    self.transfer("export", name, baseline, office, files_snapshot)
                elif self.resolve == "files" and files_snapshot is not None:
                    self.transfer("import", name, baseline, office, files_snapshot)
                else:
                    self.pause(
                        name, "Selected side is absent; deletion is not supported", baseline, office, files_snapshot
                    )
            elif name in self.state["pending"] or name in self.state["paused"]:
                self.pause(name, "Unresolved conflict or interrupted operation", baseline, office, files_snapshot)
            elif fresh_empty and office is not None:
                self.transfer("export", name, baseline, office, files_snapshot)
            else:
                action = plan(baseline, office, files_snapshot, startup=True)
                if action == "adopt":
                    self.adopt(office)
                elif action == "import":
                    # Saved source can be imported after settling in the loop.
                    continue
                else:
                    self.pause(
                        name,
                        "Startup divergence or missing component; choose authority explicitly",
                        baseline,
                        office,
                        files_snapshot,
                    )
        self.persist()

    def transfer(self, direction, name, baseline, office, files):
        # Recheck both sides immediately before mutation, including hidden attributes.
        inventory = self.adapter.inventory()
        current = self.adapter.export(name, inventory[name][0]) if name in inventory else None
        disk = self.adapter.align(self.adapter.files().get(name), current or office or files)
        if current != office or disk != files:
            if current is not None:
                self.cache[name] = current
            self.settling.pop(name, None)
            return
        if direction == "import":
            if files.attributes or (office and office.attributes):
                self.pause(name, "Import would lose procedure/member attributes", baseline, office, files)
                return
            if office and (files.header != office.header or files.path != office.path or files.kind != office.kind):
                self.pause(name, "Header/type/path changes require review", baseline, office, files)
                return
        elif files and office.path != files.path:
            self.pause(name, "Component moves require review", baseline, office, files)
            return
        artifacts = self.archive(name, direction, baseline, office, files, "operations")
        self.state["pending"][name] = {"direction": direction, "artifacts": artifacts, "phase": "apply"}
        self.persist()
        try:
            if direction == "export":
                self.adapter.write(office)
                expected = office
            else:
                self.adapter.apply(files, office)
                expected = self.adapter.export(name, files.kind)
                desired = files
                if office is None and not desired.header:
                    desired = replace(desired, header=expected.header)
                if expected != desired:
                    raise VBAError("Post-import export differs from source; baseline was not advanced")
                self.state["pending"][name].update(phase="save", expected=asdict(expected))
                self.persist()
                if not self.save(name):
                    return
            disk = self.adapter.align(self.adapter.files().get(name), expected)
            actual = self.adapter.export(name, expected.kind)
            if disk != expected or actual != expected:
                raise VBAError("Post-write verification failed; baseline was not advanced")
            self.adopt(actual)
            logger.info("%s %s", "Exported" if direction == "export" else "Imported", name)
        except (DocumentClosedError, RPCError):
            raise
        except Exception as exc:
            self.pause(name, f"Operation failed: {exc}", baseline, office, files)

    def save(self, name) -> bool:
        self.state["last_save"] = {"component": name, "attempted": time.time(), "outcome": "unknown"}
        self.persist()
        try:
            self.adapter.save(name)
        except Exception as exc:
            self.state["last_save"]["outcome"] = "failed"
            self.persist()
            logger.warning("Persistence failed for %s; will retry saving without reimport: %s", name, exc)
            return False
        self.state["last_save"]["outcome"] = "succeeded"
        self.persist()
        return True

    def scan(self):
        # A failed inventory/file read must reset settling, never signal absence.
        try:
            inventory = self.adapter.inventory()
            files = self.adapter.files()
        except (DocumentClosedError, RPCError):
            raise
        except Exception as exc:
            self.settling.clear()
            self.read_failed = True
            logger.warning("Synchronization suspended (read/busy error): %s", exc)
            return
        self.read_failed = False
        for name in sorted(set(inventory) | set(files) | set(self.state["baseline"])):
            if name in self.state["paused"] or name in self.adapter.excluded:
                continue
            baseline = self.baseline(name)
            office = self.cache.get(name) if name in inventory else None
            disk = self.adapter.align(files.get(name), office or baseline)
            signature = (inventory.get(name), disk)
            previous, count = self.settling.get(name, (None, 0))
            count = count + 1 if previous == signature else 1
            self.settling[name] = (signature, count)
            if count < SETTLED_SCANS:
                continue
            try:
                if name in inventory and (office is None or (office.kind, office.body) != inventory[name]):
                    office = self.adapter.export(name, inventory[name][0])
                    self.cache[name] = office
                    disk = self.adapter.align(files.get(name), office)
                pending = self.state["pending"].get(name)
                if pending:
                    expected = Snapshot.load(pending["expected"]) if pending.get("expected") else None
                    actual = self.adapter.export(name, inventory[name][0]) if name in inventory else None
                    if pending["phase"] == "save" and actual == expected and disk == expected:
                        if self.save(name):
                            self.adopt(expected)
                    else:
                        self.pause(name, "Pending operation changed; review before continuing", baseline, office, disk)
                    continue
                action = plan(baseline, office, disk)
                if action == "adopt":
                    if baseline != office:
                        self.adopt(office)
                elif baseline is None and office is None and disk and disk.kind == 1:
                    self.transfer("import", name, baseline, office, disk)
                elif baseline is None and office and disk is None and office.kind == 1:
                    self.transfer("export", name, baseline, office, disk)
                elif action == "conflict":
                    self.pause(name, "Both sides changed or a component is missing", baseline, office, disk)
                else:
                    self.transfer(action, name, baseline, office, disk)
            except (DocumentClosedError, RPCError):
                raise
            except Exception as exc:
                self.settling.pop(name, None)
                self.read_failed = True
                logger.warning("Retrying %s after read/busy error: %s", name, exc)

    def summary(self):
        logger.info(
            "Synchronization status: %s paused, %s pending", len(self.state["paused"]), len(self.state["pending"])
        )

    def run(self) -> None:
        document_id = hashlib.sha256(self.adapter.identity()["document"].encode()).hexdigest()
        with ExitStack() as stack:
            stack.enter_context(SessionLock(Path(tempfile.gettempdir()) / "vba-edit-locks" / f"{document_id}.bin"))
            stack.enter_context(SessionLock(self.directory / "session.bin"))
            last_scan = time.monotonic()
            last_check = last_scan
            last_summary = last_scan
            try:
                self.startup()
                if getattr(self.adapter.handler, "open_folder", False):
                    os.startfile(str(self.adapter.root))
                for _changes in watch(
                    self.adapter.root,
                    recursive=self.adapter.handler.use_rubberduck_folders,
                    rust_timeout=INTERVAL_MS,
                    yield_on_timeout=True,
                    watch_filter=lambda _change, path: not Path(path).resolve().is_relative_to(self.directory),
                ):
                    now = time.monotonic()
                    if now - last_check >= 5:
                        if not self.adapter.handler.is_document_open():
                            raise DocumentClosedError(self.adapter.handler.document_type)
                        last_check = now
                    if not _changes or now - last_scan >= INTERVAL_MS / 1000:
                        self.scan()
                        last_scan = now
                    if now - last_summary >= 30:
                        self.summary()
                        last_summary = now
            except KeyboardInterrupt:
                logger.info("Stopping bidirectional editing; Office remains open.")
            finally:
                self.summary()
            if self.state["paused"] or self.state["pending"] or self.read_failed:
                raise VBAError("Edit session ended with unresolved synchronization work; review .vba-edit/state.json.")


def run_bidirectional(handler, *, save_metadata=False, resolve=None, component=None):
    if handler.doc is None:
        handler.open_document()
    handler.sync_save_metadata = save_metadata or any(int(c.Type) == 3 for c in handler.get_vba_project().VBComponents)
    adapter = OfficeAdapter(handler)
    logger.info(
        "Bidirectional edit: %s (%s), %s; interval 2.5s, %s identical scans; non-Access imports save the document.",
        handler.doc_path,
        handler.app_name,
        adapter.project_name,
        SETTLED_SCANS,
    )
    logger.info(
        "Source exports may include unsaved VBE changes. Classes/forms/references and attribute-only changes are excluded."
    )
    module_filter = getattr(handler, "module_filter", None)
    if module_filter:
        logger.info("Module filter active (%s)", module_filter.describe())
    SyncSession(adapter, resolve=resolve, component=component).run()

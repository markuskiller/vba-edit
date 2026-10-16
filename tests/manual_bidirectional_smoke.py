"""Live smoke check using a new Office instance and disposable documents only.

Run separately from pytest (whose global fixtures terminate Office processes):
    python -m tests.manual_bidirectional_smoke --host excel --directory _workbench/spike-excel

No macros are executed and only this check's document/application are closed.
This verifies code synchronization, not the full Stage 0 performance/break-mode matrix.
"""

import argparse
import json
import statistics
import shutil
import time
from pathlib import Path

import pythoncom
import win32com.client

from vba_edit.bidirectional import OfficeAdapter, SyncSession
from vba_edit.office_vba import AccessVBAHandler, ExcelVBAHandler, PowerPointVBAHandler, WordVBAHandler


def smoke(host: str, directory: Path, seed_database: Path | None = None) -> dict:
    if host == "access" and seed_database is None:
        raise ValueError("Access requires --seed-database with a saved SmokeCode module to avoid Save As dialogs")
    directory = directory.resolve()
    directory.mkdir(parents=True, exist_ok=True)
    result = {"host": host, "checks": [], "status": "failed"}
    names = {"excel": "Excel", "word": "Word", "powerpoint": "PowerPoint", "access": "Access"}
    classes = {
        "excel": ExcelVBAHandler,
        "word": WordVBAHandler,
        "powerpoint": PowerPointVBAHandler,
        "access": AccessVBAHandler,
    }
    formats = {"excel": (".xlsm", 52), "word": (".docm", 13), "powerpoint": (".pptm", 25), "access": (".accdb", None)}
    extension, file_format = formats[host]
    document_path = directory / ("smoke" + extension)
    if document_path.exists():
        raise ValueError("Use a new directory; the smoke check never overwrites an existing document")
    application = document = None
    pythoncom.CoInitialize()
    try:
        application = win32com.client.DispatchEx(names[host] + ".Application")
        application.AutomationSecurity = 3  # Disable macros in any seed document.
        result["version"] = str(application.Version)
        if host == "excel":
            application.DisplayAlerts = False
            document = application.Workbooks.Add()
            document.SaveAs(str(document_path), FileFormat=file_format)
        elif host == "word":
            application.DisplayAlerts = 0
            document = application.Documents.Add()
            document.SaveAs2(str(document_path), FileFormat=file_format)
        elif host == "powerpoint":
            document = application.Presentations.Add(WithWindow=False)
            document.SaveAs(str(document_path), FileFormat=file_format)
        else:
            if seed_database:
                shutil.copy2(seed_database, document_path)
                application.OpenCurrentDatabase(str(document_path))
            else:
                application.NewCurrentDatabase(str(document_path))
            document = application.CurrentDb()
        project = application.VBE.ActiveVBProject if host == "access" else document.VBProject
        if host == "access" and seed_database:
            component = project.VBComponents("SmokeCode")
            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
        else:
            component = project.VBComponents.Add(1)
            component.Name = "SmokeCode"
        component.CodeModule.AddFromString('Sub Original()\n    Debug.Print "initial"\nEnd Sub')
        handler = classes[host](str(document_path), str(directory / "source"), in_file_headers=True)
        handler.app = application
        handler.doc = document
        handler.sync_save_metadata = True
        session = SyncSession(OfficeAdapter(handler))
        session.startup()
        initial = session.baseline("SmokeCode")
        assert initial and not session.state["paused"], session.state["paused"]
        result["checks"].append("startup exports including empty host modules")
        durations = []
        for _ in range(20):
            start = time.perf_counter()
            session.scan()
            durations.append((time.perf_counter() - start) * 1000)
        result["scan_ms"] = {"median": statistics.median(durations), "max": max(durations)}
        assert not session.state["paused"]
        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
        component.CodeModule.AddFromString('Sub Original()\n    Debug.Print "office"\nEnd Sub')
        session.scan()
        session.scan()
        code_path = session.adapter.path(session.baseline("SmokeCode").path)
        assert '"office"' in code_path.read_text(encoding=handler.encoding)
        result["checks"].append("Office to source with unchanged line count")
        code_path.write_text(
            code_path.read_text(encoding=handler.encoding).replace('"office"', '"source"'), encoding=handler.encoding
        )
        session.scan()
        session.scan()
        assert '"source"' in component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
        assert not session.state["paused"], session.state["paused"]
        result["checks"].append("source to Office with export readback and save")
        if host == "access":
            application.CloseCurrentDatabase()
            application.OpenCurrentDatabase(str(document_path))
            project = application.VBE.ActiveVBProject
            component = project.VBComponents("SmokeCode")
            handler.doc = application.CurrentDb()
            assert '"source"' in component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
            result["checks"].append("Access saved module survives actual database close and reopen")
        if host in {"excel", "word"}:
            host_component = next(c for c in project.VBComponents if int(c.Type) == 100)
            host_snapshot = session.baseline(host_component.Name)
            host_path = session.adapter.path(host_snapshot.path)
            host_path.write_text(host_snapshot.header + "\nSub SmokeHost()\nEnd Sub\n", encoding=handler.encoding)
            session.scan()
            session.scan()
            assert "SmokeHost" in host_component.CodeModule.Lines(1, host_component.CodeModule.CountOfLines)
            assert host_component.Name not in session.state["paused"], session.state["paused"]
            result["checks"].append("host module updated in place")
        old = session.baseline("SmokeCode")
        code_path.write_text(
            code_path.read_text(encoding=handler.encoding).replace('"source"', '"file conflict"'),
            encoding=handler.encoding,
        )
        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
        component.CodeModule.AddFromString('Sub Original()\n    Debug.Print "office conflict"\nEnd Sub')
        session.scan()
        session.scan()
        assert "SmokeCode" in session.state["paused"]
        assert session.baseline("SmokeCode") == old
        result["checks"].append("conflict preserves baseline and both sides")
        resolved = SyncSession(session.adapter, resolve="office", component="SmokeCode")
        resolved.startup()
        assert "SmokeCode" not in resolved.state["paused"], resolved.state["paused"]
        result["checks"].append("explicit Office authority resolution")
        # Reopening an older saved document is represented by restoring its body.
        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
        component.CodeModule.AddFromString(initial.body)
        restarted = SyncSession(session.adapter)
        restarted.startup()
        assert "SmokeCode" in restarted.state["paused"]
        assert '"office conflict"' in code_path.read_text(encoding=handler.encoding)
        result["checks"].append("startup divergence preserves newer source")
        result["status"] = "passed"
    except Exception as exc:
        result["error"] = repr(exc)
    finally:
        try:
            if document is not None:
                if host == "access":
                    application.DoCmd.Save(5, "SmokeCode")
                    application.CloseCurrentDatabase()
                elif host == "powerpoint":
                    document.Saved = True
                    document.Close()
                else:
                    document.Close(SaveChanges=False)
            if application is not None:
                application.Quit()
        finally:
            pythoncom.CoUninitialize()
            (directory / "result.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--host", choices=("excel", "word", "powerpoint", "access"), required=True)
    parser.add_argument("--directory", type=Path, required=True)
    parser.add_argument(
        "--seed-database", type=Path, help="Copy a disposable Access fixture with a saved SmokeCode module"
    )
    args = parser.parse_args()
    outcome = smoke(args.host, args.directory, args.seed_database)
    print(json.dumps(outcome, indent=2))
    raise SystemExit(0 if outcome["status"] == "passed" else 1)

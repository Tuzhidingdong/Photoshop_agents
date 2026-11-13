"""Safe Photoshop connection probes and read-only environment inspection."""
from __future__ import annotations

import csv
import importlib
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


def _photoshop_progids() -> list[str]:
    """Return likely Photoshop Application ProgIDs, newest first."""
    values: list[str] = []
    if os.name == "nt":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Photoshop.Application\CurVer") as key:
                cur_ver, _ = winreg.QueryValueEx(key, None)
                if cur_ver:
                    values.append(str(cur_ver))
        except OSError:
            pass
    values.extend([
        "Photoshop.Application",
        "Photoshop.Application.200",
        "Photoshop.Application.190",
        "Photoshop.Application.180",
        "Photoshop.Application.170",
        "Photoshop.Application.160",
    ])
    return list(dict.fromkeys(values))


def _get_active_win32com_application() -> tuple[Any, str]:
    """Attach to an already-running instance using late-bound COM; never activate one."""
    if os.name != "nt":
        raise OSError("win32com is available only on Windows.")
    pythoncom = importlib.import_module("pythoncom")
    dynamic = importlib.import_module("win32com.client.dynamic")
    failures: list[str] = []
    for progid in _photoshop_progids():
        try:
            clsid = pythoncom.CLSIDFromProgID(progid)
            raw_dispatch = pythoncom.GetActiveObject(clsid)
            return dynamic.Dispatch(raw_dispatch), progid
        except Exception as exc:
            failures.append(f"{progid}: {type(exc).__name__}: {exc}")
    # photoshop-python-api 0.25.0 exposes a CLSID resolver for installations
    # that do not publish versioned ProgIDs. Reuse that resolver, but still attach
    # through GetActiveObject so this path cannot activate a new Photoshop process.
    try:
        clsid_module = importlib.import_module("photoshop.api._clsid")
        class_id = clsid_module.resolve_photoshop_class_id("Application")
        if class_id:
            raw_dispatch = pythoncom.GetActiveObject(class_id)
            return dynamic.Dispatch(raw_dispatch), "CLSID:" + str(class_id)
    except Exception as exc:
        failures.append(f"Application CLSID: {type(exc).__name__}: {exc}")

    raise RuntimeError(
        "No running Photoshop COM object was found through Application ProgIDs or CLSID. "
        + " | ".join(failures[-5:])
    )


def _win32com_probe_worker(connection) -> None:
    """Child-process entry point; a timeout can terminate this worker, not Photoshop."""
    initialized = False
    try:
        pythoncom = importlib.import_module("pythoncom")
        pythoncom.CoInitialize()
        initialized = True
        app, progid = _get_active_win32com_application()
        value = app.doJavaScript('String(app.version) + String.fromCharCode(9) + String(app.documents.length);')
        fields = str(value).split(chr(9), 1)
        connection.send({
            "ok": True,
            "progid": progid,
            "photoshop_version": fields[0],
            "open_document_count": int(fields[1]) if len(fields) > 1 else None,
        })
    except BaseException as exc:
        try:
            connection.send({"ok": False, "error_type": type(exc).__name__, "error": str(exc)})
        except Exception:
            pass
    finally:
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        try:
            connection.close()
        except Exception:
            pass


def probe_win32com(timeout_seconds: float = 4.0) -> dict[str, Any]:
    """Probe a live Photoshop COM object in an isolated process with a hard timeout."""
    started = time.perf_counter()
    if os.name != "nt":
        return {
            "ok": False, "timed_out": False,
            "duration_ms": round((time.perf_counter() - started) * 1000, 1),
            "error_type": "UnsupportedPlatform",
            "error": "Photoshop COM automation is supported on Windows only.",
        }
    try:
        import multiprocessing
        context = multiprocessing.get_context("spawn")
        receiver, sender = context.Pipe(duplex=False)
        process = context.Process(target=_win32com_probe_worker, args=(sender,), daemon=True)
        process.start()
        sender.close()
        if receiver.poll(max(0.1, float(timeout_seconds))):
            try:
                result = receiver.recv()
            except EOFError:
                result = {
                    "ok": False, "timed_out": False,
                    "error_type": "ProbeProcessExited",
                    "error": "The COM probe process exited without returning a result.",
                }
            process.join(timeout=0.5)
            if process.is_alive():
                process.terminate()
                process.join(timeout=1.0)
        else:
            process.terminate()
            process.join(timeout=1.0)
            result = {
                "ok": False, "timed_out": True,
                "error_type": "ConnectionTimeout",
                "error": f"Win32COM probe exceeded {timeout_seconds:.1f} seconds; only the probe process was terminated.",
            }
        receiver.close()
    except Exception as exc:
        result = {
            "ok": False, "timed_out": False,
            "error_type": type(exc).__name__, "error": str(exc),
        }
    result.setdefault("timed_out", False)
    result["duration_ms"] = round((time.perf_counter() - started) * 1000, 1)
    return result


def connect_win32com_application() -> tuple[Any, str]:
    """Return a late-bound dispatch to an already-running Photoshop instance."""
    pythoncom = importlib.import_module("pythoncom")
    pythoncom.CoInitialize()
    return _get_active_win32com_application()


def probe_photoshop_python_api(timeout_seconds: float = 8.0) -> dict[str, Any]:
    """Test photoshop-python-api in an isolated process with a hard timeout."""
    started = time.perf_counter()
    script = (
        "import json\n"
        "try:\n"
        " import photoshop.api as api\n"
        " app=api.Application()\n"
        " value=app.doJavaScript('String(app.version) + String.fromCharCode(9) + String(app.documents.length);')\n"
        " fields=str(value).split(chr(9),1)\n"
        " result={'ok':True,'photoshop_version':fields[0],"
        "'open_document_count':int(fields[1]) if len(fields)>1 else None}\n"
        "except BaseException as exc:\n"
        " result={'ok':False,'error_type':type(exc).__name__,'error':str(exc)}\n"
        "print('SCI_PS_API_PROBE=' + json.dumps(result,ensure_ascii=False))\n"
    )
    try:
        completed = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, text=True,
            timeout=max(0.5, float(timeout_seconds)),
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0,
        )
        marker = "SCI_PS_API_PROBE="
        payload = next(
            (line[len(marker):] for line in reversed(completed.stdout.splitlines()) if line.startswith(marker)),
            None,
        )
        if payload is not None:
            result = json.loads(payload)
            result["return_code"] = completed.returncode
        else:
            result = {
                "ok": False, "timed_out": False,
                "error_type": "ProbeOutputMissing",
                "error": (completed.stderr or completed.stdout or "The API probe returned no structured result.")[-3000:],
                "return_code": completed.returncode,
            }
    except subprocess.TimeoutExpired:
        result = {
            "ok": False, "timed_out": True,
            "error_type": "ConnectionTimeout",
            "error": f"photoshop-python-api probe exceeded {timeout_seconds:.1f} seconds; only the probe process was terminated.",
        }
    except Exception as exc:
        result = {
            "ok": False, "timed_out": False,
            "error_type": type(exc).__name__, "error": str(exc),
        }
    result["duration_ms"] = round((time.perf_counter() - started) * 1000, 1)
    return result


def parse_tasklist_csv(output: str) -> list[dict[str, Any]]:
    """Parse tasklist CSV output without relying on window titles."""
    processes: list[dict[str, Any]] = []
    for row in csv.reader(output.splitlines()):
        if len(row) < 2 or row[0].strip().lower() != "photoshop.exe":
            continue
        try:
            pid = int(row[1].replace(",", "").strip())
        except ValueError:
            continue
        processes.append({
            "name": "Photoshop.exe",
            "pid": pid,
            "session_name": row[2].strip() if len(row) > 2 else None,
            "session_number": row[3].strip() if len(row) > 3 else None,
            "memory_usage": row[4].strip() if len(row) > 4 else None,
        })
    return processes


def inspect_photoshop_processes(timeout_seconds: float = 5.0) -> dict[str, Any]:
    """List Photoshop processes; window title is intentionally not a health criterion."""
    if os.name != "nt":
        return {"available": False, "processes": [], "error": "Windows process inspection is unavailable."}
    powershell = (
        "$p = Get-CimInstance Win32_Process -Filter \"Name='Photoshop.exe'\" | "
        "Select-Object ProcessId,ExecutablePath; "
        "if ($null -eq $p) { '[]' } else { ConvertTo-Json -InputObject @($p) -Compress }"
    )
    try:
        completed = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", powershell],
            capture_output=True, text=True,
            timeout=max(0.5, float(timeout_seconds)), check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if completed.returncode == 0:
            parsed = json.loads(completed.stdout.strip() or "[]")
            if isinstance(parsed, dict):
                parsed = [parsed]
            return {"available": True, "processes": parsed or [], "source": "Win32_Process"}
    except Exception as exc:
        powershell_error = f"{type(exc).__name__}: {exc}"
    else:
        powershell_error = (completed.stderr or f"PowerShell exit code {completed.returncode}").strip()
    try:
        fallback = subprocess.run(
            ["tasklist.exe", "/FI", "IMAGENAME eq Photoshop.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True,
            timeout=max(0.5, float(timeout_seconds)), check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return {
            "available": fallback.returncode == 0,
            "processes": parse_tasklist_csv(fallback.stdout),
            "source": "tasklist",
            "warning": powershell_error,
        }
    except Exception as exc:
        return {
            "available": False, "processes": [],
            "error": f"PowerShell: {powershell_error}; tasklist: {type(exc).__name__}: {exc}",
        }


def inspect_comtypes_cache() -> dict[str, Any]:
    """Report the generated-cache directory without deleting or modifying it."""
    try:
        import comtypes
        cache_dir = Path(comtypes.__file__).resolve().parent / "gen"
        if not cache_dir.exists():
            return {
                "status": "absent", "path": str(cache_dir),
                "file_count": 0, "total_bytes": 0,
                "note": "No on-disk comtypes.gen directory was found; no repair was attempted.",
            }
        files = [path for path in cache_dir.rglob("*") if path.is_file()]
        total_bytes = sum(path.stat().st_size for path in files)
        latest_mtime = max((path.stat().st_mtime for path in files), default=None)
        return {
            "status": "present", "path": str(cache_dir),
            "file_count": len(files), "total_bytes": total_bytes,
            "latest_file_mtime": latest_mtime,
            "note": "Presence alone does not establish corruption; cache was inspected read-only.",
        }
    except ImportError:
        return {"status": "not-installed", "path": None, "file_count": 0, "total_bytes": 0}
    except Exception as exc:
        return {"status": "inspection-error", "error_type": type(exc).__name__, "error": str(exc)}

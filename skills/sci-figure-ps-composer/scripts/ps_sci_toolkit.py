#!/usr/bin/env python
"""Photoshop COM/JSX bridge for sci-figure-ps-composer.

This module deliberately keeps Photoshop as the native document engine.
It provides connection, JSX execution, preference locking, and diagnostics.
"""

from __future__ import annotations

import argparse
import importlib.metadata
import json
import os
import platform
import sys
from pathlib import Path
from typing import Any, Dict, Optional


PROTOCOL_PREFIX = "SCI_PS_"


class PhotoshopBridgeError(RuntimeError):
    """A normalized failure raised by the Photoshop bridge."""

    def __init__(
        self,
        operation: str,
        message: str,
        *,
        exception_type: Optional[str] = None,
        jsx_preview: Optional[str] = None,
    ) -> None:
        super().__init__(message)
        self.operation = operation
        self.exception_type = exception_type
        self.jsx_preview = jsx_preview

    def to_dict(self) -> Dict[str, Any]:
        return {
            "operation": self.operation,
            "exception_type": self.exception_type,
            "message": str(self),
            "jsx_preview": self.jsx_preview,
        }


def _package_version(distribution: str) -> Optional[str]:
    try:
        return importlib.metadata.version(distribution)
    except importlib.metadata.PackageNotFoundError:
        return None


def _clean_jsx_preview(jsx: str, limit: int = 500) -> str:
    preview = " ".join(jsx.strip().split())
    if len(preview) > limit:
        return preview[:limit] + "..."
    return preview


def _parse_protocol(
    result: Any,
    expected_marker: str,
    minimum_fields: int,
    operation: str,
    jsx: str,
) -> list[str]:
    raw = "" if result is None else str(result).strip()
    fields = raw.split("\t")

    if len(fields) < minimum_fields or fields[0] != expected_marker:
        raise PhotoshopBridgeError(
            operation,
            "Photoshop returned an invalid bridge response: " + repr(raw),
            exception_type="InvalidBridgeResponse",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    if len(fields) > 1 and fields[1] == "ERROR":
        name = fields[2] if len(fields) > 2 else "ExtendScriptError"
        message = fields[3] if len(fields) > 3 else "Unknown JSX error"
        line = fields[4] if len(fields) > 4 else ""
        raise PhotoshopBridgeError(
            operation,
            "ExtendScript " + name + ": " + message +
            ((" (line " + line + ")") if line else ""),
            exception_type="ExtendScriptError",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    if len(fields) > 1 and fields[1] != "OK":
        raise PhotoshopBridgeError(
            operation,
            "Photoshop bridge reported an unexpected status: " + repr(raw),
            exception_type="InvalidBridgeStatus",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    return fields


class PhotoshopBridge:
    """Thin, synchronous bridge to a running desktop Photoshop instance."""

    def __init__(self, app: Any = None) -> None:
        self.app = app if app is not None else self._connect_application()

    @staticmethod
    def _connect_application() -> Any:
        if os.name != "nt":
            raise PhotoshopBridgeError(
                "connect",
                "Photoshop COM automation is supported by this toolkit on Windows only.",
                exception_type="UnsupportedPlatform",
            )

        try:
            import photoshop.api as photoshop_api
        except Exception as exc:
            raise PhotoshopBridgeError(
                "connect",
                "Could not import photoshop.api. Install photoshop-python-api "
                "into this Python interpreter. Original error: " + str(exc),
                exception_type=type(exc).__name__,
            ) from exc

        try:
            return photoshop_api.Application()
        except Exception as exc:
            raise PhotoshopBridgeError(
                "connect",
                "Could not connect to Photoshop through COM. Start desktop "
                "Photoshop, verify its COM registration, and retry. Original "
                "error: " + str(exc),
                exception_type=type(exc).__name__,
            ) from exc

    def execute_jsx(self, jsx: str) -> str:
        """Execute ExtendScript through Application.doJavaScript()."""
        if not isinstance(jsx, str) or not jsx.strip():
            raise PhotoshopBridgeError(
                "execute_jsx",
                "JSX source must be a non-empty string.",
                exception_type="InvalidJSXInput",
            )

        try:
            result = self.app.doJavaScript(jsx)
        except Exception as exc:
            raise PhotoshopBridgeError(
                "execute_jsx",
                "Photoshop rejected or failed to execute JSX. Original error: " + str(exc),
                exception_type=type(exc).__name__,
                jsx_preview=_clean_jsx_preview(jsx),
            ) from exc

        if result is None:
            return ""
        if isinstance(result, bytes):
            return result.decode("utf-8", errors="replace")
        return str(result)

    def get_status(self) -> Dict[str, Any]:
        """Read Photoshop version and current document state without modifying it."""
        jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    try {
        var count = app.documents.length;
        var activeName = count > 0 ? app.activeDocument.name : "";
        return [
            "SCI_PS_STATUS",
            "OK",
            clean(app.version),
            String(count),
            clean(activeName)
        ].join("\t");
    } catch (e) {
        return [
            "SCI_PS_STATUS",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean(e.line)
        ].join("\t");
    }
})();
"""
        raw = self.execute_jsx(jsx)
        fields = _parse_protocol(raw, "SCI_PS_STATUS", 5, "status", jsx)

        return {
            "connected": True,
            "photoshop_version": fields[2],
            "open_document_count": int(fields[3]),
            "active_document": fields[4] or None,
        }

    def lock_units(self) -> Dict[str, Any]:
        """Set Photoshop ruler units to Pixels and type units to Points."""
        jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    try {
        var previousRuler = clean(app.preferences.rulerUnits);
        var previousType = clean(app.preferences.typeUnits);

        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var currentRuler = clean(app.preferences.rulerUnits);
        var currentType = clean(app.preferences.typeUnits);
        var rulerOK = app.preferences.rulerUnits == Units.PIXELS;
        var typeOK = app.preferences.typeUnits == TypeUnits.POINTS;

        return [
            "SCI_PS_UNITS",
            "OK",
            previousRuler,
            previousType,
            currentRuler,
            currentType,
            String(rulerOK),
            String(typeOK)
        ].join("\t");
    } catch (e) {
        return [
            "SCI_PS_UNITS",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean(e.line)
        ].join("\t");
    }
})();
"""
        raw = self.execute_jsx(jsx)
        fields = _parse_protocol(raw, "SCI_PS_UNITS", 8, "lock_units", jsx)

        ruler_ok = fields[6].lower() == "true"
        type_ok = fields[7].lower() == "true"

        result = {
            "previous_ruler_units": fields[2],
            "previous_type_units": fields[3],
            "ruler_units": fields[4],
            "type_units": fields[5],
            "ruler_units_pixels": ruler_ok,
            "type_units_points": type_ok,
            "passed": ruler_ok and type_ok,
        }

        if not result["passed"]:
            raise PhotoshopBridgeError(
                "lock_units",
                "Photoshop did not retain the requested Pixels/Points units: "
                + repr(result),
                exception_type="UnitVerificationError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        return result

    def test_jsx_exception_capture(self) -> Dict[str, Any]:
        """Verify that a controlled ExtendScript exception can be caught and returned."""
        jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    var caught = false;
    var message = "";
    try {
        throw new Error("SCI_EXPECTED_TEST_EXCEPTION");
    } catch (e) {
        caught = String(e.message).indexOf("SCI_EXPECTED_TEST_EXCEPTION") >= 0;
        message = clean(e.message);
    }
    return [
        "SCI_PS_EXCEPTION",
        "OK",
        String(caught),
        message
    ].join("\t");
})();
"""
        raw = self.execute_jsx(jsx)
        fields = _parse_protocol(raw, "SCI_PS_EXCEPTION", 4, "exception_test", jsx)
        passed = fields[2].lower() == "true"

        if not passed:
            raise PhotoshopBridgeError(
                "exception_test",
                "The controlled ExtendScript exception was not captured as expected.",
                exception_type="ExceptionCaptureError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        return {
            "passed": passed,
            "exception_message": fields[3],
        }

    def diagnose(self) -> Dict[str, Any]:
        """Run non-destructive Photoshop/JSX diagnostics and return a report."""
        report: Dict[str, Any] = {
            "python_executable": sys.executable,
            "python_version": platform.python_version(),
            "photoshop_python_api_version": _package_version("photoshop-python-api"),
            "comtypes_version": _package_version("comtypes"),
            "photoshop": None,
            "unit_lock": None,
            "jsx_exception_capture": None,
            "passed": False,
        }

        report["photoshop"] = self.get_status()
        report["unit_lock"] = self.lock_units()
        report["jsx_exception_capture"] = self.test_jsx_exception_capture()
        report["passed"] = (
            report["photoshop"]["connected"]
            and report["unit_lock"]["passed"]
            and report["jsx_exception_capture"]["passed"]
        )
        return report


def _print_human_report(action: str, result: Dict[str, Any]) -> None:
    print("SCI FIGURE PHOTOSHOP BRIDGE")
    print("Action: " + action)

    if action == "diagnose":
        print("Python: " + str(result.get("python_executable")))
        print("Python version: " + str(result.get("python_version")))
        print("photoshop-python-api: " + str(result.get("photoshop_python_api_version")))
        print("comtypes: " + str(result.get("comtypes_version")))

        photoshop = result.get("photoshop") or {}
        print("COM connection: " + str(photoshop.get("connected", False)))
        print("Photoshop version: " + str(photoshop.get("photoshop_version")))
        print("Open documents: " + str(photoshop.get("open_document_count")))

        units = result.get("unit_lock") or {}
        print("Ruler units Pixels: " + str(units.get("ruler_units_pixels")))
        print("Type units Points: " + str(units.get("type_units_points")))

        exception_test = result.get("jsx_exception_capture") or {}
        print("JSX exception capture: " + str(exception_test.get("passed")))

    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))

    print("RESULT: " + ("PASS" if result.get("passed", True) else "FAILED"))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Windows Photoshop COM and ExtendScript bridge."
    )
    parser.add_argument(
        "--action",
        choices=("diagnose", "status", "lock-units", "exec-jsx"),
        default="diagnose",
        help="Bridge operation. Default: diagnose.",
    )
    parser.add_argument(
        "--jsx-file",
        type=Path,
        help="JSX file to execute when --action exec-jsx is selected.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print machine-readable JSON.",
    )
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if args.action == "exec-jsx":
        if args.jsx_file is None:
            print("ERROR: --jsx-file is required for --action exec-jsx.", file=sys.stderr)
            return 2
        if not args.jsx_file.is_file():
            print("ERROR: JSX file does not exist: " + str(args.jsx_file), file=sys.stderr)
            return 2

    try:
        bridge = PhotoshopBridge()

        if args.action == "diagnose":
            result = bridge.diagnose()
        elif args.action == "status":
            result = bridge.get_status()
            result["passed"] = True
        elif args.action == "lock-units":
            result = bridge.lock_units()
        else:
            jsx = args.jsx_file.read_text(encoding="utf-8-sig")
            result = {
                "jsx_file": str(args.jsx_file.resolve()),
                "result": bridge.execute_jsx(jsx),
                "passed": True,
            }

        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            _print_human_report(args.action, result)

        return 0 if result.get("passed", True) else 1

    except PhotoshopBridgeError as exc:
        error = exc.to_dict()
        if args.json:
            print(json.dumps({"passed": False, "error": error}, ensure_ascii=False, indent=2))
        else:
            print("BRIDGE ERROR")
            print("Operation: " + str(error["operation"]))
            print("Exception: " + str(error["exception_type"]))
            print("Message: " + str(error["message"]))
            if error.get("jsx_preview"):
                print("JSX preview: " + str(error["jsx_preview"]))
        return 1
    except Exception as exc:
        if args.json:
            print(json.dumps({
                "passed": False,
                "error": {
                    "operation": args.action,
                    "exception_type": type(exc).__name__,
                    "message": str(exc),
                },
            }, ensure_ascii=False, indent=2))
        else:
            print("UNEXPECTED ERROR: " + type(exc).__name__ + ": " + str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

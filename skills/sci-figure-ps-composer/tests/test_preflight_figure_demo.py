#!/usr/bin/env python
"""Real-Photoshop Stage 7 publication preflight integration demo.

Inspects an editable Stage 5 PSD without saving it, validates the Stage 6
export bundle, and checks the known synthetic Stage 4 scale calibration.
Expected typography/PPI warnings are reported but do not fail this non-strict
preflight; callers can use --strict to treat warnings as failures.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, preflight_figure


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-psd", type=Path, required=True)
    parser.add_argument("--export-dir", type=Path)
    parser.add_argument("--export-basename")
    parser.add_argument("--calibration-json", type=Path)
    parser.add_argument("--min-panel-ppi", type=float, default=300.0)
    parser.add_argument("--strict", action="store_true")
    args = parser.parse_args()

    input_psd = args.input_psd.expanduser().resolve(strict=True)
    export_dir = args.export_dir.expanduser().resolve(strict=True) if args.export_dir else None
    source_hash_before = sha256(input_psd)

    # This synthetic demo's Stage 4 test added 10 µm using 0.2 µm/source px.
    # Real figures must supply their own calibration values rather than reuse
    # this fixture manifest.
    calibrations = (
        json.loads(args.calibration_json.read_text(encoding="utf-8-sig"))
        if args.calibration_json
        else {
            "a": {
                "physical_length": 10.0,
                "pixel_size": 0.2,
                "unit": "µm",
                "tolerance_px": 2.0,
            }
        }
    )

    result = preflight_figure(
        input_psd,
        expected_dpi=300,
        expected_width_cm=17.6,
        min_panel_ppi=args.min_panel_ppi,
        calibrations=calibrations,
        export_dir=export_dir,
        export_basename=args.export_basename,
        strict=args.strict,
        bridge=PhotoshopBridge(),
    )

    if not result["passed"]:
        raise RuntimeError(
            "Publication preflight found " + str(result["failure_count"]) +
            " hard/strict failure(s). Inspect the checks report."
        )
    if sha256(input_psd) != source_hash_before:
        raise RuntimeError("The Stage 5 source PSD was modified by preflight.")

    checks = {item["id"]: item for item in result["checks"]}
    required_checks = (
        "canvas.geometry",
        "canvas.dpi",
        "canvas.color_mode",
        "structure.panel_groups",
        "structure.smart_objects",
        "panel_a.scale_calibration_width",
        "panel_a.scale_calibration_label",
        "source.integrity",
    )
    for check_id in required_checks:
        if check_id not in checks or checks[check_id]["status"] != "PASS":
            raise RuntimeError("Required preflight check did not pass: " + check_id)

    if export_dir is not None:
        for check_id in ("exports.editable_psd", "exports.tiff", "exports.png"):
            if check_id not in checks or checks[check_id]["status"] != "PASS":
                raise RuntimeError("Export-bundle preflight did not pass: " + check_id)

    report = {
        "passed": True,
        "input_psd": str(input_psd),
        "input_sha256_unchanged": True,
        "document": result["document"],
        "canvas_width_cm": result["canvas_width_cm"],
        "panel_count": len(result["panels"]),
        "text_layer_count": len(result["text_layers"]),
        "warning_count": result["warning_count"],
        "failure_count": result["failure_count"],
        "checks": result["checks"],
        "panels": result["panels"],
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("STAGE7_PREFLIGHT_FIGURE_REAL_PHOTOSHOP_TEST=PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE7_PREFLIGHT_FIGURE_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

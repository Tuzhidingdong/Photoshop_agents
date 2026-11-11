#!/usr/bin/env python
"""Real-Photoshop Stage 9 visual layer-index integration demo."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, inspect_psd, preview_psd_layers


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-psd", type=Path, required=True)
    parser.add_argument("--preview-dir", type=Path, required=True)
    parser.add_argument("--max-previews", type=int, default=200)
    parser.add_argument(
        "--allow-partial",
        action="store_true",
        help="Allow a deliberately partial visual index; default requires every non-background node.",
    )
    parser.add_argument("--thumbnail-width-px", type=int, default=512)
    parser.add_argument("--execution-mode", choices=("quiet", "interactive"), default="quiet")
    args = parser.parse_args()

    source = args.input_psd.expanduser().resolve(strict=True)
    before = sha256(source)
    bridge = PhotoshopBridge(execution_mode=args.execution_mode)

    inspection = inspect_psd(source, bridge=bridge)
    candidates = [
        layer for layer in inspection["layers"]
        if not layer["is_background"]
    ]
    if not candidates:
        raise RuntimeError("The PSD has no non-background layers to preview.")

    selected = candidates[: min(len(candidates), args.max_previews)]
    if not args.allow_partial and len(selected) != len(candidates):
        raise RuntimeError(
            "Full visual coverage is required by default: "
            + str(len(candidates)) + " non-background nodes exist, but --max-previews "
            + str(args.max_previews) + " would preview only " + str(len(selected))
            + ". Increase --max-previews or explicitly pass --allow-partial."
        )
    result = preview_psd_layers(
        source,
        args.preview_dir,
        node_keys=[layer["node_key"] for layer in selected],
        max_previews=args.max_previews,
        thumbnail_width_px=args.thumbnail_width_px,
        bridge=bridge,
    )

    if not result["passed"]:
        raise RuntimeError("No layer preview was rendered successfully.")
    composite = Path(result["composite_preview"])
    contact_sheet = Path(result["contact_sheet"]["path"])
    report_path = Path(result["report_path"])

    for output in (composite, contact_sheet, report_path):
        if not output.is_file() or output.stat().st_size <= 0:
            raise RuntimeError("Required visual index output is missing or empty: " + str(output))

    successful = [
        preview for preview in result["previews"]
        if preview["status"] == "OK"
    ]
    if not successful:
        raise RuntimeError("No individual layer preview succeeded.")
    if not args.allow_partial:
        expected_keys = {layer["node_key"] for layer in candidates}
        actual_keys = {preview["node_key"] for preview in result["previews"]}
        if actual_keys != expected_keys:
            raise RuntimeError("Full visual coverage failed: preview node keys do not match the PSD.")
        if result["failed_count"] != 0:
            raise RuntimeError(
                "Full visual coverage failed: "
                + str(result["failed_count"]) + " individual preview(s) failed."
            )

    after = sha256(source)
    if before != after or not result["source_sha256_unchanged"]:
        raise RuntimeError("Visual layer preview generation modified the source PSD.")

    report = {
        "passed": True,
        "input_psd": str(source),
        "layer_tree_sha256": inspection["layer_tree_sha256"],
        "selected_layer_count": len(selected),
        "eligible_non_background_node_count": len(candidates),
        "full_layer_coverage": len(selected) == len(candidates) and result["failed_count"] == 0,
        "successful_preview_count": result["successful_count"],
        "failed_preview_count": result["failed_count"],
        "composite_preview": str(composite),
        "layer_contact_sheet": str(contact_sheet),
        "preview_report": str(report_path),
        "source_sha256_unchanged": True,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("STAGE9_PREVIEW_PSD_LAYERS_REAL_PHOTOSHOP_TEST=PASS")
    print("STAGE9_FULL_LAYER_COVERAGE=" + ("PASS" if report["full_layer_coverage"] else "PARTIAL"))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE9_PREVIEW_PSD_LAYERS_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

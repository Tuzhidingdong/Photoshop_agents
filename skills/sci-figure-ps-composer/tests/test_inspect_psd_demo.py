#!/usr/bin/env python
"""Real-Photoshop read-only Stage 8 PSD inspection integration demo."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, inspect_psd


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-psd", type=Path, required=True)
    parser.add_argument("--report-path", type=Path)
    parser.add_argument("--preview-dir", type=Path)
    parser.add_argument("--preview-max-layers", type=int, default=12)
    parser.add_argument("--preview-max-dimension", type=int, default=1400)
    parser.add_argument("--execution-mode", choices=("quiet", "interactive"), default="quiet")
    args = parser.parse_args()

    source = args.input_psd.expanduser().resolve(strict=True)
    before = sha256(source)
    result = inspect_psd(
        source,
        report_path=args.report_path,
        preview_dir=args.preview_dir,
        preview_max_layers=args.preview_max_layers,
        preview_max_dimension=args.preview_max_dimension,
        bridge=PhotoshopBridge(execution_mode=args.execution_mode),
    )
    after = sha256(source)

    if before != after or not result["source_sha256_unchanged"]:
        raise RuntimeError("Read-only inspection modified the source PSD.")
    if result["layer_count"] != len(result["layers"]):
        raise RuntimeError("Inspection layer count does not match the layer tree.")
    if result["layer_count"] < 1:
        raise RuntimeError("The source PSD has no inspectable layers.")
    if args.preview_dir:
        previews = result["visual_previews"]
        composite = next(
            (item for item in previews if item["node_key"] == "__composite__" and item["status"] == "PASS"),
            None,
        )
        if composite is None or not Path(composite["path"]).is_file():
            raise RuntimeError("Composite visual preview was not generated.")
        if not any(item["status"] == "PASS" for item in previews if item["node_key"] != "__composite__"):
            raise RuntimeError("No individual layer preview could be generated.")
    if len({layer["node_key"] for layer in result["layers"]}) != result["layer_count"]:
        raise RuntimeError("Inspection returned duplicate node keys.")

    report = {
        "passed": True,
        "input_psd": str(source),
        "source_sha256_unchanged": True,
        "document": result["document"],
        "layer_count": result["layer_count"],
        "group_count": result["group_count"],
        "text_layer_count": result["text_layer_count"],
        "smart_object_count": result["smart_object_count"],
        "layer_tree_sha256": result["layer_tree_sha256"],
        "report_path": result.get("report_path"),
        "visual_previews": result.get("visual_previews", []),
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("STAGE8_INSPECT_PSD_REAL_PHOTOSHOP_TEST=PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE8_INSPECT_PSD_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

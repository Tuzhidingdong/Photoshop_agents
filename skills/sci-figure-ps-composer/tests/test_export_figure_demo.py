#!/usr/bin/env python
"""Real-Photoshop Stage 6 publication export integration demo.

The test exports a byte-identical editable PSD copy plus flattened RGB LZW
TIFF and RGB PNG from a separate Photoshop duplicate. It reopens all three
outputs in Photoshop and inspects raster metadata without re-encoding them.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, export_figure


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input-psd",
        type=Path,
        required=True,
        help="Stage 5 editable PSD; it is never modified.",
    )
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--basename")
    parser.add_argument("--expected-dpi", type=int, choices=(300, 600))
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_psd = args.input_psd.expanduser().resolve(strict=True)
    output_dir = args.output_dir.expanduser().resolve() if args.output_dir else None
    source_hash_before = sha256(input_psd)

    result = export_figure(
        input_psd,
        output_dir,
        basename=args.basename,
        expected_dpi=args.expected_dpi,
        overwrite=args.overwrite,
        bridge=PhotoshopBridge(),
    )

    editable_psd = Path(result["editable_psd"]["path"])
    tiff_path = Path(result["tiff"]["path"])
    png_path = Path(result["png"]["path"])

    for path in (editable_psd, tiff_path, png_path):
        if not path.is_file() or path.stat().st_size <= 0:
            raise RuntimeError("Missing or empty export: " + str(path))

    if sha256(input_psd) != source_hash_before:
        raise RuntimeError("The supplied Stage 5 PSD was modified.")
    if sha256(editable_psd) != source_hash_before:
        raise RuntimeError("The exported editable PSD is not byte-identical to its input.")

    if result["mode"] != "RGB":
        raise RuntimeError("Exported source mode is not RGB.")
    if result["dpi"] not in (300.0, 600.0):
        raise RuntimeError("Expected a 300 or 600 DPI source PSD.")
    if result["tiff"]["compression_tag"] != 5:
        raise RuntimeError("TIFF compression tag is not LZW (5).")
    if result["tiff"]["mode"] != "RGB" or result["png"]["mode"] != "RGB":
        raise RuntimeError("TIFF and PNG must both be RGB.")
    if result["tiff"]["width_px"] != result["width_px"] or result["png"]["width_px"] != result["width_px"]:
        raise RuntimeError("Raster export width differs from the PSD canvas.")
    if result["tiff"]["height_px"] != result["height_px"] or result["png"]["height_px"] != result["height_px"]:
        raise RuntimeError("Raster export height differs from the PSD canvas.")

    report = {
        "passed": True,
        "input_psd": str(input_psd),
        "input_sha256_unchanged": True,
        "output_dir": str(editable_psd.parent),
        "editable_psd": result["editable_psd"],
        "tiff": result["tiff"],
        "png": result["png"],
        "source_document": {
            "width_px": result["width_px"],
            "height_px": result["height_px"],
            "dpi": result["dpi"],
            "mode": result["mode"],
            "panel_group_count": result["panel_group_count"],
            "smart_object_count": result["smart_object_count"],
            "text_layer_count": result["text_layer_count"],
        },
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("STAGE6_EXPORT_FIGURE_REAL_PHOTOSHOP_TEST=PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE6_EXPORT_FIGURE_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

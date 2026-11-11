#!/usr/bin/env python
"""Real-Photoshop Stage 2 integration demo for compose_grid.

Creates four synthetic raster panels, composes them in Photoshop as embedded
Smart Objects, saves a PSD, reopens it in Photoshop, and checks its structure.
This test must run on Windows with desktop Photoshop open.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

from PIL import Image, ImageDraw

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, compose_grid


def create_demo_panels(output_dir: Path) -> list[Path]:
    """Create four RGB panels with a white border for Auto-Trim validation."""
    colors = [
        (65, 125, 190),
        (190, 100, 65),
        (75, 155, 105),
        (145, 100, 175),
    ]
    paths: list[Path] = []

    for index, color in enumerate(colors):
        image = Image.new("RGB", (800, 600), "white")
        draw = ImageDraw.Draw(image)

        # Opaque interior with deliberately broad pure-white outer margin.
        draw.rectangle((48, 48, 751, 551), fill=color)
        draw.ellipse((170, 145, 390, 365), outline="white", width=12)
        draw.line((430, 130, 660, 470), fill="black", width=9)
        draw.rectangle((465, 185, 650, 335), outline="white", width=8)
        draw.text((72, 70), f"TEST PANEL {index + 1}", fill="white")

        path = output_dir / f"panel_{chr(ord('a') + index)}.png"
        image.save(path, dpi=(300, 300))
        paths.append(path)

    return paths


def verify_saved_psd(bridge: PhotoshopBridge, psd_path: Path) -> dict:
    """Reopen the saved PSD in Photoshop and inspect dimensions and layer types."""
    jsx_path = json.dumps(str(psd_path.resolve()), ensure_ascii=True)
    jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    var verificationDoc = null;
    var previousDialogs = app.displayDialogs;
    try {
        app.displayDialogs = DialogModes.NO;
        var file = new File(__PSD_PATH__);
        if (!file.exists) {
            throw new Error("PSD file does not exist: " + file.fsName);
        }

        verificationDoc = app.open(file);

        var width = verificationDoc.width.as("px");
        var height = verificationDoc.height.as("px");
        var resolution = verificationDoc.resolution;
        var isRGB = verificationDoc.mode == DocumentMode.RGB;
        var groupCount = verificationDoc.layerSets.length;
        var smartObjectCount = 0;
        var names = [];

        for (var i = 0; i < verificationDoc.layerSets.length; i++) {
            var group = verificationDoc.layerSets[i];
            names.push(clean(group.name));

            for (var j = 0; j < group.artLayers.length; j++) {
                if (group.artLayers[j].kind == LayerKind.SMARTOBJECT) {
                    smartObjectCount++;
                }
            }
        }

        verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
        verificationDoc = null;
        return [
            "SCI_PS_VERIFY",
            "OK",
            String(width),
            String(height),
            String(resolution),
            String(isRGB),
            String(groupCount),
            String(smartObjectCount),
            names.join(",")
        ].join("\t");

    } catch (e) {
        if (verificationDoc !== null) {
            try {
                verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeError) {}
        }
        return [
            "SCI_PS_VERIFY",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean(e.line)
        ].join("\t");
    } finally {
        app.displayDialogs = previousDialogs;
    }
})();
""".replace("__PSD_PATH__", jsx_path)

    fields = bridge.execute_jsx(jsx).rstrip("\r\n").split("\t")
    if len(fields) < 9 or fields[0] != "SCI_PS_VERIFY" or fields[1] != "OK":
        raise RuntimeError("Photoshop PSD verification failed: " + repr(fields))

    return {
        "width_px": int(fields[2]),
        "height_px": int(fields[3]),
        "dpi": float(fields[4]),
        "rgb": fields[5].lower() == "true",
        "panel_group_count": int(fields[6]),
        "smart_object_count": int(fields[7]),
        "group_names": fields[8].split(",") if fields[8] else [],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Directory for demo input images and the generated PSD.",
    )
    args = parser.parse_args()

    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir
        else Path.home() / "Documents" / "sci_figure_ps_composer_demo"
        / time.strftime("%Y%m%d_%H%M%S")
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    print("Creating four synthetic panels in:", output_dir)
    images = create_demo_panels(output_dir)
    original_sizes = {path: path.stat().st_size for path in images}
    psd_path = output_dir / "Figure_Grid_Stage2.psd"

    bridge = PhotoshopBridge()
    result = compose_grid(
        images,
        psd_path,
        preset="double",
        rows=2,
        cols=2,
        dpi=300,
        margin_cm=0.3,
        gap_cm=0.2,
        align="center",
        auto_trim=True,
        overwrite=False,
        bridge=bridge,
    )

    if not psd_path.is_file() or psd_path.stat().st_size <= 0:
        raise RuntimeError("Photoshop did not create a non-empty PSD.")
    if result["width_px"] != 2079:
        raise RuntimeError("Expected 2079 px canvas width; got " + str(result["width_px"]))
    if result["panel_count"] != 4:
        raise RuntimeError("Expected four composed panels.")
    if result["trimmed_panel_count"] != 4:
        raise RuntimeError(
            "Expected four white-bordered Smart Objects to be trimmed; got "
            + str(result["trimmed_panel_count"])
        )
    if any(path.stat().st_size != size for path, size in original_sizes.items()):
        raise RuntimeError("One or more source images were modified.")

    verification = verify_saved_psd(bridge, psd_path)

    if verification["width_px"] != 2079:
        raise RuntimeError("Reopened PSD width does not match 2079 px.")
    if verification["height_px"] != result["height_px"]:
        raise RuntimeError("Reopened PSD height does not match the composed document.")
    if abs(verification["dpi"] - 300) > 0.01:
        raise RuntimeError("Reopened PSD resolution is not 300 DPI.")
    if not verification["rgb"]:
        raise RuntimeError("Reopened PSD is not RGB.")
    if verification["panel_group_count"] != 4:
        raise RuntimeError("Reopened PSD does not contain four panel groups.")
    if verification["smart_object_count"] != 4:
        raise RuntimeError("Reopened PSD does not contain four Smart Objects.")

    report = {
        "passed": True,
        "output_dir": str(output_dir),
        "psd_path": str(psd_path),
        "psd_bytes": psd_path.stat().st_size,
        "composition": result,
        "reopened_psd_verification": verification,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("STAGE2_COMPOSE_GRID_REAL_PHOTOSHOP_TEST=PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE2_COMPOSE_GRID_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

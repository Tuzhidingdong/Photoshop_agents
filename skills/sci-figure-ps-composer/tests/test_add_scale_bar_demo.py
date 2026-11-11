#!/usr/bin/env python
"""Real-Photoshop Stage 4 integration demo for editable scale bars.

The physical scale is calculated from the selected panel's embedded Smart
Object dimensions after trimming and its actual on-canvas transform.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, add_scale_bar


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def verify_scale_bar_psd(bridge: PhotoshopBridge, psd_path: Path) -> dict:
    """Reopen the output and verify bar, outline, text, and physical geometry."""
    jsx_path = json.dumps(str(psd_path.resolve()), ensure_ascii=True)
    jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function px(value) {
        return value.as("px");
    }

    var verificationDoc = null;
    var smartDoc = null;
    var previousDialogs = app.displayDialogs;
    try {
        app.displayDialogs = DialogModes.NO;
        var file = new File(__PSD_PATH__);
        if (!file.exists) {
            throw new Error("Scale-bar PSD does not exist: " + file.fsName);
        }

        verificationDoc = app.open(file);
        var targetGroup = null;
        var panelGroupCount = 0;
        var scaleLabelCount = 0;
        for (var i = 0; i < verificationDoc.layerSets.length; i++) {
            var group = verificationDoc.layerSets[i];
            if (group.name.indexOf("Group_Panel_") === 0) {
                panelGroupCount++;
            }
            if (group.name === "Group_Panel_a") {
                targetGroup = group;
            }
        }
        if (panelGroupCount !== 4) {
            throw new Error("Expected four panel groups; got " + panelGroupCount);
        }
        if (targetGroup === null) {
            throw new Error("Group_Panel_a is missing.");
        }

        var smartLayer = null;
        var barLayer = null;
        var outlineLayer = null;
        var scaleTextLayer = null;
        var originalPanelLabel = null;
        for (var j = 0; j < targetGroup.artLayers.length; j++) {
            var layer = targetGroup.artLayers[j];
            if (layer.name.indexOf("Image_Panel_") === 0 &&
                layer.kind == LayerKind.SMARTOBJECT) {
                smartLayer = layer;
            } else if (layer.name === "Scale_Bar_a") {
                barLayer = layer;
            } else if (layer.name === "Scale_Bar_Outline_a") {
                outlineLayer = layer;
            } else if (layer.name === "Scale_Label_a") {
                scaleTextLayer = layer;
                scaleLabelCount++;
            } else if (layer.name === "Panel_Label_a") {
                originalPanelLabel = layer;
            }
        }

        if (smartLayer === null || barLayer === null ||
            outlineLayer === null || scaleTextLayer === null) {
            throw new Error("One or more expected editable scale-bar layers are missing.");
        }
        if (barLayer.kind !== LayerKind.NORMAL) {
            throw new Error("Scale rectangle is not a separate normal layer.");
        }
        if (outlineLayer.kind !== LayerKind.NORMAL) {
            throw new Error("Scale outline is not a separate normal layer.");
        }
        if (scaleTextLayer.kind !== LayerKind.TEXT) {
            throw new Error("Scale label is not an editable Photoshop text layer.");
        }
        if (originalPanelLabel === null || originalPanelLabel.kind !== LayerKind.TEXT) {
            throw new Error("The existing panel label was not preserved.");
        }

        var textItem = scaleTextLayer.textItem;
        var textContents = String(textItem.contents);
        var fontName = String(textItem.font);
        var fontSize = textItem.size.as("pt");
        var isBlack = (
            Math.round(textItem.color.rgb.red) === 0 &&
            Math.round(textItem.color.rgb.green) === 0 &&
            Math.round(textItem.color.rgb.blue) === 0
        );
        if (textContents !== "10 µm") {
            throw new Error("Expected scale label '10 µm'; got " + textContents);
        }
        if (fontName !== "ArialMT") {
            throw new Error("Expected ArialMT; got " + fontName);
        }
        if (Math.abs(fontSize - 8.0) > 0.05) {
            throw new Error("Expected 8 pt scale label; got " + fontSize);
        }
        if (!isBlack) {
            throw new Error("Expected a black scale label.");
        }

        var imageBounds = smartLayer.bounds;
        var imageLeft = px(imageBounds[0]);
        var imageTop = px(imageBounds[1]);
        var imageRight = px(imageBounds[2]);
        var imageBottom = px(imageBounds[3]);

        var barBounds = barLayer.bounds;
        var barLeft = px(barBounds[0]);
        var barTop = px(barBounds[1]);
        var barRight = px(barBounds[2]);
        var barBottom = px(barBounds[3]);
        var barWidth = barRight - barLeft;
        var barHeight = barBottom - barTop;

        var outlineBounds = outlineLayer.bounds;
        var outlineLeft = px(outlineBounds[0]);
        var outlineTop = px(outlineBounds[1]);
        var outlineRight = px(outlineBounds[2]);
        var outlineBottom = px(outlineBounds[3]);
        var outlineWidth = outlineRight - outlineLeft;

        if (!(barWidth > 1 && barHeight > 1)) {
            throw new Error("Scale rectangle has invalid dimensions.");
        }
        if (!(outlineWidth > barWidth && outlineBottom - outlineTop > barHeight)) {
            throw new Error("Contrasting outline does not extend beyond the bar.");
        }
        if (barRight > imageRight || barBottom > imageBottom ||
            barLeft < imageLeft || barTop < imageTop) {
            throw new Error("Scale bar is outside the panel image bounds.");
        }
        if ((imageRight - barRight) < 10 || (imageBottom - barBottom) < 10) {
            throw new Error("Scale bar is not inset from the bottom-right panel edge.");
        }

        var labelBounds = scaleTextLayer.bounds;
        var labelLeft = px(labelBounds[0]);
        var labelTop = px(labelBounds[1]);
        var labelRight = px(labelBounds[2]);
        var labelBottom = px(labelBounds[3]);
        if (labelLeft < imageLeft || labelRight > imageRight ||
            labelTop < imageTop || labelBottom > barTop) {
            throw new Error("Scale label is not positioned above the bar within the panel.");
        }

        verificationDoc.activeLayer = smartLayer;
        executeAction(
            stringIDToTypeID("placedLayerEditContents"),
            new ActionDescriptor(),
            DialogModes.NO
        );
        smartDoc = app.activeDocument;
        var smartWidth = px(smartDoc.width);
        var smartHeight = px(smartDoc.height);
        smartDoc.close(SaveOptions.DONOTSAVECHANGES);
        smartDoc = null;
        app.activeDocument = verificationDoc;

        var displayedWidth = imageRight - imageLeft;
        var displayedPerSourcePixel = displayedWidth / smartWidth;
        var expectedWidth = (10.0 / 0.2) * displayedPerSourcePixel;
        if (Math.abs(barWidth - expectedWidth) > 2.0) {
            throw new Error(
                "Physical scale-bar width mismatch: got " + barWidth +
                ", expected approximately " + expectedWidth
            );
        }

        // Other panel groups must not receive scale layers.
        for (var groupIndex = 0; groupIndex < verificationDoc.layerSets.length; groupIndex++) {
            var otherGroup = verificationDoc.layerSets[groupIndex];
            if (otherGroup.name.indexOf("Group_Panel_") !== 0 ||
                otherGroup.name === "Group_Panel_a") {
                continue;
            }
            for (var otherIndex = 0; otherIndex < otherGroup.artLayers.length; otherIndex++) {
                if (otherGroup.artLayers[otherIndex].name.indexOf("Scale_") === 0) {
                    throw new Error("A scale layer was added to an unselected panel.");
                }
            }
        }

        verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
        verificationDoc = null;

        return [
            "SCI_PS_VERIFY_SCALE_BAR",
            "OK",
            String(panelGroupCount),
            String(scaleLabelCount),
            String(smartWidth),
            String(smartHeight),
            String(displayedPerSourcePixel),
            String(barWidth),
            String(expectedWidth),
            fontName,
            String(fontSize),
            textContents
        ].join("\t");

    } catch (e) {
        if (smartDoc !== null) {
            try {
                smartDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (smartCloseError) {}
        }
        if (verificationDoc !== null) {
            try {
                verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeError) {}
        }
        return [
            "SCI_PS_VERIFY_SCALE_BAR",
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
    if (
        len(fields) < 12
        or fields[0] != "SCI_PS_VERIFY_SCALE_BAR"
        or fields[1] != "OK"
    ):
        raise RuntimeError("Photoshop scale-bar verification failed: " + repr(fields))

    return {
        "panel_group_count": int(fields[2]),
        "scale_label_count": int(fields[3]),
        "smart_object_width_px": int(float(fields[4])),
        "smart_object_height_px": int(float(fields[5])),
        "displayed_pixels_per_source_pixel": float(fields[6]),
        "bar_width_px": float(fields[7]),
        "expected_bar_width_px": float(fields[8]),
        "font_postscript": fields[9],
        "font_size_pt": float(fields[10]),
        "label": fields[11],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input-psd",
        type=Path,
        required=True,
        help="Stage 3 labeled PSD; the input is never modified.",
    )
    parser.add_argument(
        "--output-psd",
        type=Path,
        help="Output path. Defaults to <input-stem>_Stage4_ScaleBar.psd.",
    )
    parser.add_argument("--panel-id", default="a")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_psd = args.input_psd.expanduser().resolve(strict=True)
    output_psd = (
        args.output_psd.expanduser().resolve()
        if args.output_psd
        else input_psd.with_name(input_psd.stem + "_Stage4_ScaleBar.psd")
    )
    source_hash_before = sha256(input_psd)

    bridge = PhotoshopBridge()
    result = add_scale_bar(
        input_psd,
        output_psd,
        panel_id=args.panel_id,
        mode="physical",
        physical_length=10.0,
        pixel_size=0.2,
        unit="µm",
        color_hex="#000000",
        outline_hex="#FFFFFF",
        bar_height_pt=2.0,
        outline_width_pt=0.5,
        label_size_pt=8.0,
        label_gap_pt=1.5,
        inset_cm=0.2,
        bridge=bridge,
        overwrite=args.overwrite,
    )

    if not output_psd.is_file() or output_psd.stat().st_size <= 0:
        raise RuntimeError("Photoshop did not create a non-empty scale-bar PSD.")

    verification = verify_scale_bar_psd(bridge, output_psd)
    if sha256(input_psd) != source_hash_before:
        raise RuntimeError("The source PSD was modified; this must never happen.")
    if verification["panel_group_count"] != 4:
        raise RuntimeError("Expected four panel groups.")
    if verification["scale_label_count"] != 1:
        raise RuntimeError("Expected exactly one scale label on the selected panel.")
    if abs(
        verification["bar_width_px"] - verification["expected_bar_width_px"]
    ) > 2.0:
        raise RuntimeError("Physical scale-bar calibration did not verify.")

    report = {
        "passed": True,
        "input_psd": str(input_psd),
        "input_sha256_unchanged": True,
        "output_psd": str(output_psd),
        "output_bytes": output_psd.stat().st_size,
        "operation": result,
        "reopened_psd_verification": verification,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("STAGE4_SCALE_BAR_REAL_PHOTOSHOP_TEST=PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE4_SCALE_BAR_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

#!/usr/bin/env python
"""Real-Photoshop Stage 3 integration demo for editable panel labels.

Run on Windows with desktop Photoshop open, using a PSD created by the
Stage 2 compose_grid integration demo.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, add_panel_labels


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def verify_labeled_psd(bridge: PhotoshopBridge, psd_path: Path) -> dict:
    """Reopen the PSD and inspect native text layers and their formatting."""
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
    var previousDialogs = app.displayDialogs;
    try {
        app.displayDialogs = DialogModes.NO;
        var file = new File(__PSD_PATH__);
        if (!file.exists) {
            throw new Error("Labeled PSD does not exist: " + file.fsName);
        }

        verificationDoc = app.open(file);
        var groupCount = 0;
        var labelCount = 0;
        var editableTextCount = 0;
        var labels = [];
        var checks = [];

        for (var i = 0; i < verificationDoc.layerSets.length; i++) {
            var group = verificationDoc.layerSets[i];
            if (group.name.indexOf("Group_Panel_") !== 0) {
                continue;
            }
            groupCount++;

            var smartLayer = null;
            var panelLabel = null;
            for (var j = 0; j < group.artLayers.length; j++) {
                var layer = group.artLayers[j];
                if (layer.name.indexOf("Image_Panel_") === 0 &&
                    layer.kind == LayerKind.SMARTOBJECT) {
                    smartLayer = layer;
                }
                if (layer.name.indexOf("Panel_Label_") === 0) {
                    panelLabel = layer;
                }
            }

            if (panelLabel === null) {
                throw new Error("No panel label layer found in " + group.name);
            }
            if (smartLayer === null) {
                throw new Error("No embedded Smart Object found in " + group.name);
            }
            if (panelLabel.kind !== LayerKind.TEXT) {
                throw new Error("Panel label is not an editable text layer: " + group.name);
            }

            var item = panelLabel.textItem;
            var labelContents = String(item.contents);
            var fontName = String(item.font);
            var fontSize = item.size.as("pt");
            var isBlack = (
                Math.round(item.color.rgb.red) === 0 &&
                Math.round(item.color.rgb.green) === 0 &&
                Math.round(item.color.rgb.blue) === 0
            );
            var textPosition = item.position;
            var textX = px(textPosition[0]);
            var textY = px(textPosition[1]);
            var imageBounds = smartLayer.bounds;
            var imageLeft = px(imageBounds[0]);
            var imageTop = px(imageBounds[1]);
            var imageRight = px(imageBounds[2]);
            var imageBottom = px(imageBounds[3]);

            if (fontName !== "Arial-BoldMT") {
                throw new Error(
                    "Expected Arial-BoldMT in " + group.name + "; got " + fontName
                );
            }
            if (Math.abs(fontSize - 10.0) > 0.05) {
                throw new Error(
                    "Expected 10 pt panel label in " + group.name + "; got " + fontSize
                );
            }
            if (!isBlack) {
                throw new Error("Panel label is not black: " + group.name);
            }
            if (!(textX >= imageLeft && textX < imageRight &&
                  textY > imageTop && textY < imageBottom)) {
                throw new Error("Panel label position is outside its panel image: " + group.name);
            }

            labelCount++;
            editableTextCount++;
            labels.push(group.name.substring("Group_Panel_".length) + "=" + labelContents);
            checks.push(
                group.name + ":font=" + fontName +
                ",size=" + fontSize +
                ",black=" + isBlack +
                ",editable=" + (panelLabel.kind == LayerKind.TEXT)
            );
        }

        if (groupCount !== 4 || labelCount !== 4 || editableTextCount !== 4) {
            throw new Error(
                "Expected 4 panel groups and 4 editable labels; got groups=" +
                groupCount + ", labels=" + labelCount
            );
        }

        verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
        verificationDoc = null;

        return [
            "SCI_PS_VERIFY_LABELS",
            "OK",
            String(groupCount),
            String(labelCount),
            String(editableTextCount),
            labels.join(","),
            checks.join(";")
        ].join("\t");

    } catch (e) {
        if (verificationDoc !== null) {
            try {
                verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeError) {}
        }
        return [
            "SCI_PS_VERIFY_LABELS",
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
    if len(fields) < 7 or fields[0] != "SCI_PS_VERIFY_LABELS" or fields[1] != "OK":
        raise RuntimeError("Photoshop label verification failed: " + repr(fields))

    labels = fields[5].split(",") if fields[5] else []
    return {
        "panel_group_count": int(fields[2]),
        "label_layer_count": int(fields[3]),
        "editable_text_layer_count": int(fields[4]),
        "labels": labels,
        "format_checks": fields[6].split(";") if fields[6] else [],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input-psd",
        type=Path,
        required=True,
        help="Stage 2 PSD to label; this source file is never overwritten.",
    )
    parser.add_argument(
        "--output-psd",
        type=Path,
        help="Output path. Defaults to <input-stem>_Stage3_Labeled.psd.",
    )
    parser.add_argument(
        "--style",
        choices=("lowercase", "uppercase", "parenthesized"),
        default="lowercase",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_psd = args.input_psd.expanduser().resolve(strict=True)
    output_psd = (
        args.output_psd.expanduser().resolve()
        if args.output_psd
        else input_psd.with_name(input_psd.stem + "_Stage3_Labeled.psd")
    )
    source_hash_before = sha256(input_psd)

    bridge = PhotoshopBridge()
    result = add_panel_labels(
        input_psd,
        output_psd,
        style=args.style,
        size_pt=10.0,
        font_postscript="Arial-BoldMT",
        inset_cm=0.1,
        color_hex="#000000",
        overwrite=args.overwrite,
        bridge=bridge,
    )

    if not output_psd.is_file() or output_psd.stat().st_size <= 0:
        raise RuntimeError("Photoshop did not create a non-empty labeled PSD.")

    verification = verify_labeled_psd(bridge, output_psd)
    if sha256(input_psd) != source_hash_before:
        raise RuntimeError("The source PSD was modified; this must never happen.")
    if verification["panel_group_count"] != 4:
        raise RuntimeError("Expected four panel groups.")
    if verification["label_layer_count"] != 4:
        raise RuntimeError("Expected four panel label layers.")
    if verification["editable_text_layer_count"] != 4:
        raise RuntimeError("Expected four editable text layers.")

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
    print("STAGE3_PANEL_LABELS_REAL_PHOTOSHOP_TEST=PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE3_PANEL_LABELS_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

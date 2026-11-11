#!/usr/bin/env python
"""Real-Photoshop Stage 5 integration demo for recursive text normalization.

The demo creates a separate fixture PSD containing a text layer nested two
groups deep, normalizes every text layer into another PSD, then reopens that
output and verifies font, point size, and color across the full hierarchy.
The supplied input PSD is never saved or modified.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, unify_fonts


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def create_nested_text_fixture(
    bridge: PhotoshopBridge,
    source_psd: Path,
    fixture_psd: Path,
    *,
    overwrite: bool,
) -> dict:
    """Create a copy with one text layer two LayerSets deep, without saving source."""
    source_js = json.dumps(str(source_psd.resolve()), ensure_ascii=True)
    fixture_js = json.dumps(str(fixture_psd.resolve()), ensure_ascii=True)
    overwrite_js = "true" if overwrite else "false"
    jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function px(value) {
        return value.as("px");
    }
    function makeColor(red, green, blue) {
        var color = new SolidColor();
        color.rgb.red = red;
        color.rgb.green = green;
        color.rgb.blue = blue;
        return color;
    }

    var previousDialogs = app.displayDialogs;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var doc = null;
    var response = "";
    try {
        app.displayDialogs = DialogModes.NO;
        var sourceFile = new File(__SOURCE_PSD__);
        if (!sourceFile.exists) {
            throw new Error("Stage 4 input PSD does not exist.");
        }
        var fixtureFile = new File(__FIXTURE_PSD__);
        if (fixtureFile.exists) {
            if (!__OVERWRITE_FIXTURE__) {
                throw new Error("Fixture already exists; pass --overwrite to replace it.");
            }
            if (!fixtureFile.remove()) {
                throw new Error("Could not remove existing fixture PSD.");
            }
        }

        doc = app.open(sourceFile);
        var panelGroup = null;
        var panelGroupCount = 0;
        for (var i = 0; i < doc.layerSets.length; i++) {
            var group = doc.layerSets[i];
            if (group.name.indexOf("Group_Panel_") === 0) {
                panelGroupCount++;
            }
            if (group.name === "Group_Panel_b") {
                panelGroup = group;
            }
        }
        if (panelGroupCount !== 4 || panelGroup === null) {
            throw new Error("Expected four panel groups including Group_Panel_b.");
        }

        var smartLayer = null;
        for (var j = 0; j < panelGroup.artLayers.length; j++) {
            var candidate = panelGroup.artLayers[j];
            if (candidate.name.indexOf("Image_Panel_") === 0 &&
                candidate.kind == LayerKind.SMARTOBJECT) {
                smartLayer = candidate;
                break;
            }
        }
        if (smartLayer === null) {
            throw new Error("Panel b Smart Object is missing.");
        }

        var imageBounds = smartLayer.bounds;
        var imageLeft = px(imageBounds[0]);
        var imageTop = px(imageBounds[1]);

        var outer = panelGroup.layerSets.add();
        outer.name = "Stage5_Test_Outer";
        var inner = outer.layerSets.add();
        inner.name = "Stage5_Test_Inner";

        var nestedText = inner.artLayers.add();
        nestedText.kind = LayerKind.TEXT;
        nestedText.name = "Stage5_Nested_Text";
        var textItem = nestedText.textItem;
        textItem.contents = "Nested Font Test";
        textItem.font = "ArialMT";
        textItem.size = new UnitValue(6, "pt");
        textItem.color = makeColor(0, 0, 255);
        textItem.position = [
            new UnitValue(imageLeft + 24, "px"),
            new UnitValue(imageTop + 70, "px")
        ];

        var psdOptions = new PhotoshopSaveOptions();
        psdOptions.embedColorProfile = true;
        psdOptions.maximizeCompatibility = true;
        doc.saveAs(fixtureFile, psdOptions, true, Extension.LOWERCASE);

        if (!fixtureFile.exists || fixtureFile.length <= 0) {
            throw new Error("Photoshop did not save the nested-text fixture.");
        }

        response = [
            "SCI_PS_STAGE5_FIXTURE",
            "OK",
            String(panelGroupCount),
            fixtureFile.fsName,
            nestedText.name
        ].join("\t");
    } catch (e) {
        response = [
            "SCI_PS_STAGE5_FIXTURE",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean((e.line === undefined) ? "" : e.line)
        ].join("\t");
    } finally {
        if (doc !== null) {
            try {
                doc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeError) {}
        }
        if (previousDoc !== null) {
            try {
                app.activeDocument = previousDoc;
            } catch (restoreError) {}
        }
        app.displayDialogs = previousDialogs;
    }
    return response;
})();
""".replace("__SOURCE_PSD__", source_js)
    jsx = jsx.replace("__FIXTURE_PSD__", fixture_js)
    jsx = jsx.replace("__OVERWRITE_FIXTURE__", overwrite_js)

    fields = bridge.execute_jsx(jsx).rstrip("\r\n").split("\t")
    if len(fields) < 5 or fields[0] != "SCI_PS_STAGE5_FIXTURE" or fields[1] != "OK":
        raise RuntimeError("Could not create nested-text fixture: " + repr(fields))
    return {
        "panel_group_count": int(fields[2]),
        "fixture_psd": fields[3],
        "nested_text_layer": fields[4],
    }


def verify_unified_fonts_psd(
    bridge: PhotoshopBridge,
    psd_path: Path,
    *,
    expected_font: str,
    expected_size_pt: float,
) -> dict:
    """Reopen and inspect text layers in every nested group."""
    jsx_path = json.dumps(str(psd_path.resolve()), ensure_ascii=True)
    font_js = json.dumps(expected_font, ensure_ascii=True)
    size_js = repr(float(expected_size_pt))
    jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function isBlack(color) {
        return (
            Math.round(color.rgb.red) === 0 &&
            Math.round(color.rgb.green) === 0 &&
            Math.round(color.rgb.blue) === 0
        );
    }

    var verificationDoc = null;
    var previousDialogs = app.displayDialogs;
    var textCount = 0;
    var panelGroupCount = 0;
    var panelLabelCount = 0;
    var scaleLabelCount = 0;
    var nestedTextFound = false;
    var checks = [];

    function inspectContainer(container, parentPath) {
        for (var i = 0; i < container.artLayers.length; i++) {
            var layer = container.artLayers[i];
            if (layer.kind !== LayerKind.TEXT) {
                continue;
            }

            textCount++;
            var textItem = layer.textItem;
            var fontName = String(textItem.font);
            var fontSize = textItem.size.as("pt");
            var black = isBlack(textItem.color);
            var layerPath = parentPath + "/" + layer.name;

            if (fontName !== __EXPECTED_FONT__) {
                throw new Error("Unexpected font in " + layerPath + ": " + fontName);
            }
            if (Math.abs(fontSize - __EXPECTED_SIZE_PT__) > 0.05) {
                throw new Error("Unexpected point size in " + layerPath + ": " + fontSize);
            }
            if (!black) {
                throw new Error("Text layer is not black: " + layerPath);
            }

            if (layer.name.indexOf("Panel_Label_") === 0) {
                panelLabelCount++;
            }
            if (layer.name.indexOf("Scale_Label_") === 0) {
                scaleLabelCount++;
            }
            if (layer.name === "Stage5_Nested_Text" &&
                layerPath.indexOf("Stage5_Test_Outer/Stage5_Test_Inner") >= 0) {
                nestedTextFound = true;
            }

            checks.push(
                layerPath + ":font=" + fontName +
                ",size=" + fontSize.toFixed(3) +
                ",black=" + black
            );
        }

        for (var groupIndex = 0; groupIndex < container.layerSets.length; groupIndex++) {
            var group = container.layerSets[groupIndex];
            inspectContainer(group, parentPath + "/" + group.name);
        }
    }

    try {
        app.displayDialogs = DialogModes.NO;
        var file = new File(__PSD_PATH__);
        if (!file.exists) {
            throw new Error("Font-normalized PSD does not exist.");
        }

        verificationDoc = app.open(file);
        for (var i = 0; i < verificationDoc.layerSets.length; i++) {
            if (verificationDoc.layerSets[i].name.indexOf("Group_Panel_") === 0) {
                panelGroupCount++;
            }
        }
        inspectContainer(verificationDoc, verificationDoc.name);

        if (panelGroupCount !== 4) {
            throw new Error("Expected four panel groups; got " + panelGroupCount);
        }
        if (textCount < 6) {
            throw new Error("Expected at least six text layers; got " + textCount);
        }
        if (panelLabelCount !== 4) {
            throw new Error("Expected four preserved panel labels; got " + panelLabelCount);
        }
        if (scaleLabelCount !== 1) {
            throw new Error("Expected one preserved scale label; got " + scaleLabelCount);
        }
        if (!nestedTextFound) {
            throw new Error("The text layer two groups deep was not normalized.");
        }

        verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
        verificationDoc = null;

        return [
            "SCI_PS_VERIFY_UNIFY_FONTS",
            "OK",
            String(panelGroupCount),
            String(textCount),
            String(panelLabelCount),
            String(scaleLabelCount),
            String(nestedTextFound),
            checks.join(";")
        ].join("\t");
    } catch (e) {
        if (verificationDoc !== null) {
            try {
                verificationDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeError) {}
        }
        return [
            "SCI_PS_VERIFY_UNIFY_FONTS",
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
    jsx = jsx.replace("__EXPECTED_FONT__", font_js)
    jsx = jsx.replace("__EXPECTED_SIZE_PT__", size_js)

    fields = bridge.execute_jsx(jsx).rstrip("\r\n").split("\t")
    if (
        len(fields) < 8
        or fields[0] != "SCI_PS_VERIFY_UNIFY_FONTS"
        or fields[1] != "OK"
    ):
        raise RuntimeError("Photoshop font verification failed: " + repr(fields))
    return {
        "panel_group_count": int(fields[2]),
        "text_layer_count": int(fields[3]),
        "panel_label_count": int(fields[4]),
        "scale_label_count": int(fields[5]),
        "nested_text_layer_verified": fields[6].lower() == "true",
        "format_checks": fields[7].split(";") if fields[7] else [],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-psd", type=Path, required=True)
    parser.add_argument("--fixture-psd", type=Path)
    parser.add_argument("--output-psd", type=Path)
    parser.add_argument("--font-postscript", default="ArialMT")
    parser.add_argument("--font-size-pt", type=float, default=9.0)
    parser.add_argument("--font-color", default="#000000")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_psd = args.input_psd.expanduser().resolve(strict=True)
    fixture_psd = (
        args.fixture_psd.expanduser().resolve()
        if args.fixture_psd
        else input_psd.with_name(input_psd.stem + "_Stage5_FontsFixture.psd")
    )
    output_psd = (
        args.output_psd.expanduser().resolve()
        if args.output_psd
        else input_psd.with_name(input_psd.stem + "_Stage5_FontsUnified.psd")
    )

    input_hash_before = sha256(input_psd)
    bridge = PhotoshopBridge()

    fixture_result = create_nested_text_fixture(
        bridge,
        input_psd,
        fixture_psd,
        overwrite=args.overwrite,
    )
    fixture_hash_before = sha256(fixture_psd)

    operation = unify_fonts(
        fixture_psd,
        output_psd,
        font_postscript=args.font_postscript,
        size_pt=args.font_size_pt,
        color_hex=args.font_color,
        overwrite=args.overwrite,
        bridge=bridge,
    )

    if not output_psd.is_file() or output_psd.stat().st_size <= 0:
        raise RuntimeError("Photoshop did not create a non-empty font-normalized PSD.")

    verification = verify_unified_fonts_psd(
        bridge,
        output_psd,
        expected_font=args.font_postscript,
        expected_size_pt=args.font_size_pt,
    )

    if sha256(input_psd) != input_hash_before:
        raise RuntimeError("The supplied Stage 4 PSD was modified; this must never happen.")
    if sha256(fixture_psd) != fixture_hash_before:
        raise RuntimeError("The nested-text fixture PSD was modified; this must never happen.")
    if verification["panel_group_count"] != 4:
        raise RuntimeError("Expected four panel groups.")
    if verification["text_layer_count"] < 6:
        raise RuntimeError("Expected at least six normalized text layers.")
    if verification["panel_label_count"] != 4 or verification["scale_label_count"] != 1:
        raise RuntimeError("Existing panel or scale labels were not preserved.")
    if not verification["nested_text_layer_verified"]:
        raise RuntimeError("Recursive normalization of nested text did not verify.")

    report = {
        "passed": True,
        "input_psd": str(input_psd),
        "input_sha256_unchanged": True,
        "fixture_psd": str(fixture_psd),
        "fixture_sha256_unchanged": True,
        "output_psd": str(output_psd),
        "output_bytes": output_psd.stat().st_size,
        "fixture": fixture_result,
        "operation": operation,
        "reopened_psd_verification": verification,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("STAGE5_UNIFY_FONTS_REAL_PHOTOSHOP_TEST=PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE5_UNIFY_FONTS_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

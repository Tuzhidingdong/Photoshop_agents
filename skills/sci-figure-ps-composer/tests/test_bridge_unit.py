from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from ps_sci_toolkit import (
    PhotoshopBridge,
    PhotoshopBridgeError,
    calculate_grid_layout,
    cm_to_pixels,
    compose_grid,
    _detect_trim_steps,
    add_panel_labels,
    add_scale_bar,
    _panel_label_text,
    _physical_scale_width_px,
    unify_fonts,
    export_figure,
    _inspect_export_raster,
    preflight_figure,
    inspect_psd,
    _validate_organize_plan,
    apply_organize_plan,
    _parse_organize_verify_records,
    _build_apply_organize_plan_jsx,
    _build_preview_psd_layers_jsx,
    validate_psd_classification,
    create_organize_plan_draft,
    _parse_psd_preview_records,
    preview_psd_layers,
)


class FakePhotoshopApplication:
    def __init__(self, responses=None, error=None):
        self.responses = list(responses or [])
        self.error = error
        self.calls = []

    def doJavaScript(self, jsx):
        self.calls.append(jsx)
        if self.error is not None:
            raise self.error
        if not self.responses:
            return ""
        return self.responses.pop(0)


class PhotoshopBridgeUnitTests(unittest.TestCase):
    def test_execute_jsx_passes_source_to_photoshop(self):
        app = FakePhotoshopApplication(["42"])
        bridge = PhotoshopBridge(app=app)

        result = bridge.execute_jsx("1 + 1;")

        self.assertEqual(result, "42")
        self.assertEqual(app.calls, ["1 + 1;"])

    def test_interactive_execution_mode_allows_photoshop_dialogs(self):
        app = FakePhotoshopApplication(["ok"])
        bridge = PhotoshopBridge(app=app, execution_mode="interactive")

        self.assertEqual(bridge.execute_jsx("app.displayDialogs = DialogModes.NO;"), "ok")
        self.assertIn("DialogModes.ALL", app.calls[0])
        self.assertNotIn("DialogModes.NO", app.calls[0])

    def test_execution_mode_rejects_unknown_values(self):
        with self.assertRaises(ValueError):
            PhotoshopBridge(app=FakePhotoshopApplication(), execution_mode="headless")

    def test_execute_jsx_rejects_empty_source(self):
        bridge = PhotoshopBridge(app=FakePhotoshopApplication())

        with self.assertRaises(PhotoshopBridgeError) as caught:
            bridge.execute_jsx("   ")

        self.assertEqual(caught.exception.operation, "execute_jsx")
        self.assertEqual(caught.exception.exception_type, "InvalidJSXInput")

    def test_execute_jsx_wraps_com_or_host_exception(self):
        app = FakePhotoshopApplication(error=OSError("COM server unavailable"))
        bridge = PhotoshopBridge(app=app)

        with self.assertRaises(PhotoshopBridgeError) as caught:
            bridge.execute_jsx("app.version;")

        self.assertEqual(caught.exception.operation, "execute_jsx")
        self.assertEqual(caught.exception.exception_type, "OSError")
        self.assertIn("COM server unavailable", str(caught.exception))
        self.assertIsNotNone(caught.exception.jsx_preview)

    def test_status_protocol_is_parsed(self):
        app = FakePhotoshopApplication([
            "SCI_PS_STATUS\tOK\t27.9.1\t0\t"
        ])
        bridge = PhotoshopBridge(app=app)

        result = bridge.get_status()

        self.assertTrue(result["connected"])
        self.assertEqual(result["photoshop_version"], "27.9.1")
        self.assertEqual(result["open_document_count"], 0)
        self.assertIsNone(result["active_document"])

    def test_lock_units_protocol_is_parsed_and_verified(self):
        app = FakePhotoshopApplication([
            "SCI_PS_UNITS\tOK\tUnits.INCHES\tTypeUnits.PICAS\t"
            "Units.PIXELS\tTypeUnits.POINTS\ttrue\ttrue"
        ])
        bridge = PhotoshopBridge(app=app)

        result = bridge.lock_units()

        self.assertTrue(result["passed"])
        self.assertEqual(result["previous_ruler_units"], "Units.INCHES")
        self.assertEqual(result["previous_type_units"], "TypeUnits.PICAS")
        self.assertEqual(result["ruler_units"], "Units.PIXELS")
        self.assertEqual(result["type_units"], "TypeUnits.POINTS")

    def test_lock_units_fails_if_readback_is_wrong(self):
        app = FakePhotoshopApplication([
            "SCI_PS_UNITS\tOK\tUnits.INCHES\tTypeUnits.PICAS\t"
            "Units.INCHES\tTypeUnits.POINTS\tfalse\ttrue"
        ])
        bridge = PhotoshopBridge(app=app)

        with self.assertRaises(PhotoshopBridgeError) as caught:
            bridge.lock_units()

        self.assertEqual(caught.exception.operation, "lock_units")
        self.assertEqual(caught.exception.exception_type, "UnitVerificationError")

    def test_jsx_exception_capture_protocol_is_verified(self):
        app = FakePhotoshopApplication([
            "SCI_PS_EXCEPTION\tOK\ttrue\tSCI_EXPECTED_TEST_EXCEPTION"
        ])
        bridge = PhotoshopBridge(app=app)

        result = bridge.test_jsx_exception_capture()

        self.assertTrue(result["passed"])
        self.assertEqual(
            result["exception_message"],
            "SCI_EXPECTED_TEST_EXCEPTION",
        )

    def test_invalid_protocol_response_is_not_treated_as_success(self):
        app = FakePhotoshopApplication(["unexpected response"])
        bridge = PhotoshopBridge(app=app)

        with self.assertRaises(PhotoshopBridgeError) as caught:
            bridge.get_status()

        self.assertEqual(caught.exception.exception_type, "InvalidBridgeResponse")

    def test_short_jsx_error_response_is_preserved(self):
        app = FakePhotoshopApplication([
            "SCI_PS_UNITS\tERROR\tUnitError\tCannot set units\t44"
        ])
        bridge = PhotoshopBridge(app=app)

        with self.assertRaises(PhotoshopBridgeError) as caught:
            bridge.lock_units()

        self.assertEqual(caught.exception.exception_type, "ExtendScriptError")
        self.assertIn("Cannot set units", str(caught.exception))
        self.assertIn("line 44", str(caught.exception))



class FakeComposeBridge:
    def __init__(self, response):
        self.response = response
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        self.jsx_calls.append(jsx)
        return self.response


class ComposeGridUnitTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.images = []
        for index, color in enumerate(("red", "blue", "green", "orange")):
            path = self.root / f"panel_{index + 1}.png"
            Image.new("RGB", (100, 100), color=color).save(path)
            self.images.append(path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_double_column_width_is_2079_pixels_at_300_dpi(self):
        self.assertEqual(cm_to_pixels(17.6, 300), 2079)

    def test_layout_derives_square_2_by_2_height(self):
        layout = calculate_grid_layout(
            [(100, 100)] * 4,
            width_cm=17.6,
            rows=2,
            cols=2,
            dpi=300,
            margin_cm=0.3,
            gap_cm=0.2,
        )

        self.assertEqual(layout["width_px"], 2079)
        self.assertEqual(layout["height_px"], 2079)
        self.assertEqual(len(layout["panels"]), 4)
        self.assertEqual(
            [panel["panel"] for panel in layout["panels"]],
            ["a", "b", "c", "d"],
        )
        self.assertLess(layout["panels"][0]["x"], layout["panels"][1]["x"])
        self.assertLess(layout["panels"][0]["y"], layout["panels"][2]["y"])

    def test_custom_height_is_respected(self):
        layout = calculate_grid_layout(
            [(160, 90)] * 4,
            width_cm=17.6,
            height_cm=10.0,
            rows=2,
            cols=2,
            dpi=300,
        )
        self.assertEqual(layout["height_px"], 1181)

    def test_layout_rejects_wrong_number_of_images(self):
        with self.assertRaises(ValueError):
            calculate_grid_layout(
                [(100, 100)] * 3,
                width_cm=17.6,
                rows=2,
                cols=2,
            )

    def test_compose_grid_builds_native_smart_object_jsx(self):
        layout = calculate_grid_layout(
            [(100, 100)] * 4,
            width_cm=17.6,
            rows=2,
            cols=2,
            dpi=300,
            margin_cm=0.3,
            gap_cm=0.2,
        )
        bridge = FakeComposeBridge(
            "SCI_PS_COMPOSE\tOK\tDemo\t2079\t2079\t300\t4\t0\t"
        )

        original_sizes = [path.stat().st_size for path in self.images]
        result = compose_grid(
            self.images,
            preset="double",
            rows=2,
            cols=2,
            dpi=300,
            margin_cm=0.3,
            gap_cm=0.2,
            auto_trim=True,
            bridge=bridge,
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["width_px"], layout["width_px"])
        self.assertEqual(result["height_px"], layout["height_px"])
        self.assertEqual(result["panel_count"], 4)
        self.assertTrue(result["auto_trim"])
        self.assertIn('executeAction(charIDToTypeID("Plc ")', bridge.jsx_calls[0])
        self.assertIn('placedLayerEditContents', bridge.jsx_calls[0])
        self.assertIn('if (trimStep === "transparent")', bridge.jsx_calls[0])
        self.assertIn('smartDoc.trim(TrimType.TOPLEFT, true, true, true, true)', bridge.jsx_calls[0])
        self.assertIn('smartDoc.trim(TrimType.TRANSPARENT, true, true, true, true)', bridge.jsx_calls[0])
        self.assertIn('panel.trim_steps.length > 0', bridge.jsx_calls[0])
        self.assertIn('doc.close(SaveOptions.DONOTSAVECHANGES)', bridge.jsx_calls[0])
        self.assertIn('Group_Panel_a', bridge.jsx_calls[0])
        self.assertEqual(original_sizes, [path.stat().st_size for path in self.images])

    def test_trim_preflight_detects_opaque_white_margin_without_alpha_trim(self):
        path = self.root / "white_margin.png"
        image = Image.new("RGB", (100, 80), "white")
        draw = __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(image)
        draw.rectangle((10, 8, 89, 71), fill=(100, 120, 140))
        image.save(path)

        self.assertEqual(_detect_trim_steps(path), ["top_left"])

    def test_trim_preflight_detects_transparency_and_skips_unneeded_white_trim(self):
        path = self.root / "transparent_margin.png"
        image = Image.new("RGBA", (100, 80), (0, 0, 0, 0))
        draw = __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(image)
        draw.rectangle((10, 8, 89, 71), fill=(100, 120, 140, 255))
        image.save(path)

        self.assertEqual(_detect_trim_steps(path), ["transparent"])

    def test_trim_preflight_skips_solid_image_without_trim_border(self):
        path = self.root / "solid.png"
        Image.new("RGB", (100, 80), (100, 120, 140)).save(path)

        self.assertEqual(_detect_trim_steps(path), [])

    def test_compose_grid_rejects_existing_output_without_overwrite(self):
        output = self.root / "existing.psd"
        output.write_bytes(b"keep")
        bridge = FakeComposeBridge("unused")

        with self.assertRaises(FileExistsError):
            compose_grid(
                self.images,
                output,
                preset="double",
                rows=2,
                cols=2,
                bridge=bridge,
            )

        self.assertEqual(output.read_bytes(), b"keep")
        self.assertEqual(bridge.jsx_calls, [])


class FakeLabelBridge:
    def __init__(self, output_path, response):
        self.output_path = Path(output_path)
        self.response = response
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        self.jsx_calls.append(jsx)
        self.output_path.write_bytes(b"mock-labeled-psd")
        return self.response


class PanelLabelUnitTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Figure_Grid.psd"
        self.output_psd = self.root / "Figure_Grid_Labeled.psd"
        self.input_psd.write_bytes(b"source-psd")

    def tearDown(self):
        self.temp_dir.cleanup()

    def _bridge(self):
        response = (
            "SCI_PS_LABELS\tOK\tFigure_Grid_Labeled\t4\t4\tlowercase\t"
            "Group_Panel_d=d,Group_Panel_c=c,Group_Panel_b=b,Group_Panel_a=a\t"
            + str(self.output_psd)
        )
        return FakeLabelBridge(self.output_psd, response)

    def test_panel_label_styles(self):
        self.assertEqual(_panel_label_text("a", "lowercase"), "a")
        self.assertEqual(_panel_label_text("b", "uppercase"), "B")
        self.assertEqual(_panel_label_text("aa", "parenthesized"), "(aa)")
        with self.assertRaises(ValueError):
            _panel_label_text("panel-a", "lowercase")

    def test_add_panel_labels_builds_native_editable_text_layers(self):
        source_before = self.input_psd.read_bytes()
        bridge = self._bridge()

        result = add_panel_labels(
            self.input_psd,
            self.output_psd,
            style="lowercase",
            size_pt=10,
            inset_cm=0.1,
            bridge=bridge,
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["panel_group_count"], 4)
        self.assertEqual(result["labels_created"], 4)
        self.assertEqual(result["style"], "lowercase")
        self.assertEqual(result["font_postscript"], "Arial-BoldMT")
        self.assertEqual(result["font_size_pt"], 10)
        self.assertEqual(result["color_hex"], "#000000")
        self.assertEqual(self.input_psd.read_bytes(), source_before)
        self.assertEqual(self.output_psd.read_bytes(), b"mock-labeled-psd")

        jsx = bridge.jsx_calls[0]
        self.assertIn("LayerKind.TEXT", jsx)
        self.assertIn("Panel_Label_", jsx)
        self.assertIn('new UnitValue(cfg.size_pt, "pt")', jsx)
        self.assertIn("Arial-BoldMT", jsx)
        self.assertIn("textItem.position", jsx)
        self.assertIn("doc.saveAs(outputFile", jsx)
        self.assertIn("doc.close(SaveOptions.DONOTSAVECHANGES)", jsx)

    def test_add_panel_labels_supports_uppercase_and_parenthesized(self):
        for style, expected in (
            ("uppercase", "uppercase"),
            ("parenthesized", "parenthesized"),
        ):
            with self.subTest(style=style):
                output = self.root / (style + ".psd")
                response = (
                    "SCI_PS_LABELS\tOK\tLabeled\t4\t4\t" + expected +
                    "\tGroup_Panel_a=x\t" + str(output)
                )
                bridge = FakeLabelBridge(output, response)
                result = add_panel_labels(
                    self.input_psd,
                    output,
                    style=style,
                    bridge=bridge,
                )
                self.assertEqual(result["style"], style)

    def test_add_panel_labels_rejects_source_overwrite(self):
        bridge = self._bridge()
        with self.assertRaises(ValueError):
            add_panel_labels(
                self.input_psd,
                self.input_psd,
                overwrite=True,
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])

    def test_add_panel_labels_rejects_existing_output_without_overwrite(self):
        self.output_psd.write_bytes(b"keep")
        bridge = self._bridge()
        with self.assertRaises(FileExistsError):
            add_panel_labels(
                self.input_psd,
                self.output_psd,
                bridge=bridge,
            )
        self.assertEqual(self.output_psd.read_bytes(), b"keep")
        self.assertEqual(bridge.jsx_calls, [])

    def test_add_panel_labels_validates_size_and_color(self):
        bridge = self._bridge()
        with self.assertRaises(ValueError):
            add_panel_labels(
                self.input_psd,
                self.output_psd,
                size_pt=9,
                bridge=bridge,
            )
        with self.assertRaises(ValueError):
            add_panel_labels(
                self.input_psd,
                self.output_psd,
                color_hex="black",
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])



class FakeScaleBarBridge:
    def __init__(self, output_path, response):
        self.output_path = Path(output_path)
        self.response = response
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        self.jsx_calls.append(jsx)
        self.output_path.write_bytes(b"mock-scale-bar-psd")
        return self.response


class ScaleBarUnitTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Figure_Labeled.psd"
        self.output_psd = self.root / "Figure_ScaleBar.psd"
        self.input_psd.write_bytes(b"source-psd")

    def tearDown(self):
        self.temp_dir.cleanup()

    def _bridge(self, output=None, mode="pixels", width="120", label="100 px"):
        output = Path(output or self.output_psd)
        response = (
            "SCI_PS_SCALE_BAR\tOK\tFigure_ScaleBar\ta\t" + mode + "\t" +
            width + "\t" + label + "\t1000\t800\t0.75\ttrue\t" + str(output)
        )
        return FakeScaleBarBridge(output, response)

    def test_physical_scale_width_accounts_for_smart_object_transform(self):
        # 10 µm / 0.2 µm per source pixel = 50 source pixels.
        # A 1000 px Smart Object displayed at 1500 px gives a 1.5x transform.
        self.assertAlmostEqual(
            _physical_scale_width_px(10, 0.2, 1500, 1000),
            75.0,
        )

    def test_physical_scale_width_rejects_invalid_dimensions(self):
        with self.assertRaises(ValueError):
            _physical_scale_width_px(10, 0.2, 1500, 0)

    def test_add_scale_bar_pixel_mode_builds_separate_editable_layers(self):
        source_before = self.input_psd.read_bytes()
        bridge = self._bridge(mode="pixels", width="120")
        result = add_scale_bar(
            self.input_psd,
            self.output_psd,
            panel_id="a",
            mode="pixels",
            bar_width_px=120,
            outline_hex="#FFFFFF",
            bridge=bridge,
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["panel_id"], "a")
        self.assertEqual(result["mode"], "pixels")
        self.assertEqual(result["bar_width_px"], 120)
        self.assertEqual(result["label"], "100 px")
        self.assertTrue(result["outline"])
        self.assertEqual(self.input_psd.read_bytes(), source_before)
        self.assertEqual(self.output_psd.read_bytes(), b"mock-scale-bar-psd")

        jsx = bridge.jsx_calls[0]
        self.assertIn("Scale_Bar_Outline_", jsx)
        self.assertIn("Scale_Bar_", jsx)
        self.assertIn("Scale_Label_", jsx)
        self.assertIn("LayerKind.TEXT", jsx)
        self.assertIn("parentDoc.selection.fill", jsx)
        self.assertIn("Justification.RIGHT", jsx)
        self.assertIn("labelTargetRight = innerRight", jsx)
        self.assertIn("labelTargetBottom = barTop - labelGapPx", jsx)
        self.assertIn("textLayer.translate(", jsx)
        self.assertIn("labelBottom > barTop - labelGapPx + 0.5", jsx)
        self.assertIn("placedLayerEditContents", jsx)
        self.assertIn("displayedImageWidth / sourceWidthPx", jsx)
        self.assertIn("doc.close(SaveOptions.DONOTSAVECHANGES)", jsx)

    def test_add_scale_bar_physical_mode_uses_calibrated_label(self):
        bridge = self._bridge(mode="physical", width="75", label="10 µm")
        result = add_scale_bar(
            self.input_psd,
            self.output_psd,
            mode="physical",
            physical_length=10,
            pixel_size=0.2,
            unit="µm",
            color_hex="#FFFFFF",
            outline_hex="#000000",
            bridge=bridge,
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["mode"], "physical")
        self.assertEqual(result["label"], "10 µm")
        self.assertEqual(result["smart_object_width_px"], 1000)
        self.assertEqual(result["smart_object_height_px"], 800)
        self.assertEqual(result["displayed_pixels_per_source_pixel"], 0.75)
        self.assertTrue(result["outline"])
        self.assertIn('"label":"10 \\u00b5m"', bridge.jsx_calls[0])

    def test_add_scale_bar_rejects_source_overwrite(self):
        bridge = self._bridge()
        with self.assertRaises(ValueError):
            add_scale_bar(
                self.input_psd,
                self.input_psd,
                bar_width_px=100,
                overwrite=True,
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])

    def test_add_scale_bar_rejects_existing_output_without_overwrite(self):
        self.output_psd.write_bytes(b"keep")
        bridge = self._bridge()
        with self.assertRaises(FileExistsError):
            add_scale_bar(
                self.input_psd,
                self.output_psd,
                bar_width_px=100,
                bridge=bridge,
            )
        self.assertEqual(self.output_psd.read_bytes(), b"keep")
        self.assertEqual(bridge.jsx_calls, [])

    def test_add_scale_bar_validates_mode_calibration_and_contrast(self):
        bridge = self._bridge()
        with self.assertRaises(ValueError):
            add_scale_bar(self.input_psd, self.output_psd, mode="pixels", bridge=bridge)
        with self.assertRaises(ValueError):
            add_scale_bar(
                self.input_psd,
                self.output_psd,
                mode="physical",
                physical_length=10,
                pixel_size=0.2,
                bridge=bridge,
            )
        with self.assertRaises(ValueError):
            add_scale_bar(
                self.input_psd,
                self.output_psd,
                bar_width_px=100,
                color_hex="#FF0000",
                bridge=bridge,
            )
        with self.assertRaises(ValueError):
            add_scale_bar(
                self.input_psd,
                self.output_psd,
                bar_width_px=100,
                color_hex="#000000",
                outline_hex="#000000",
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])


class FakeUnifyFontsBridge:
    def __init__(self, output_path, response):
        self.output_path = Path(output_path)
        self.response = response
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        self.jsx_calls.append(jsx)
        self.output_path.write_bytes(b"mock-font-normalized-psd")
        return self.response


class UnifyFontsUnitTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Figure_ScaleBar.psd"
        self.output_psd = self.root / "Figure_FontsUnified.psd"
        self.input_psd.write_bytes(b"source-psd")

    def tearDown(self):
        self.temp_dir.cleanup()

    def _bridge(
        self,
        *,
        output=None,
        font="Arial-BoldMT",
        size="9",
        color="#000000",
        text_count="6",
        changed_count="6",
    ):
        output = Path(output or self.output_psd)
        response = (
            "SCI_PS_UNIFY_FONTS\tOK\tFigure_FontsUnified\t" +
            text_count + "\t" + changed_count + "\t" +
            font + "\t" + size + "\t" + color + "\t" +
            str(output) + "\tGroup_Panel_a/Panel_Label_a;Group_Panel_b/Nested/Test"
        )
        return FakeUnifyFontsBridge(output, response)

    def test_unify_fonts_recursively_targets_native_text_layers(self):
        source_before = self.input_psd.read_bytes()
        bridge = self._bridge()
        result = unify_fonts(
            self.input_psd,
            self.output_psd,
            font_postscript="Arial-BoldMT",
            size_pt=9,
            color_hex="#000000",
            bridge=bridge,
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["text_layer_count"], 6)
        self.assertEqual(result["changed_layer_count"], 6)
        self.assertEqual(result["font_postscript"], "Arial-BoldMT")
        self.assertEqual(result["font_size_pt"], 9.0)
        self.assertEqual(result["color_hex"], "#000000")
        self.assertEqual(len(result["layer_paths"]), 2)
        self.assertEqual(self.input_psd.read_bytes(), source_before)
        self.assertEqual(self.output_psd.read_bytes(), b"mock-font-normalized-psd")

        jsx = bridge.jsx_calls[0]
        self.assertIn("function processContainer(container, parentPath)", jsx)
        self.assertIn("container.artLayers", jsx)
        self.assertIn("container.layerSets", jsx)
        self.assertIn("LayerKind.TEXT", jsx)
        self.assertIn('new UnitValue(cfg.size_pt, "pt")', jsx)
        self.assertIn("textItem.font = cfg.font_postscript", jsx)
        self.assertIn("textItem.color = makeColor(cfg.color_rgb)", jsx)
        self.assertIn("doc.saveAs(outputFile", jsx)
        self.assertIn("doc.close(SaveOptions.DONOTSAVECHANGES)", jsx)

    def test_unify_fonts_can_selectively_change_only_one_attribute(self):
        bridge = self._bridge(font="ArialMT", size="", color="")
        result = unify_fonts(
            self.input_psd,
            self.output_psd,
            font_postscript="ArialMT",
            bridge=bridge,
        )
        self.assertEqual(result["font_postscript"], "ArialMT")
        self.assertIsNone(result["font_size_pt"])
        self.assertIsNone(result["color_hex"])
        self.assertIn('"size_pt":null', bridge.jsx_calls[0])
        self.assertIn('"color_rgb":null', bridge.jsx_calls[0])

    def test_unify_fonts_rejects_no_requested_attributes(self):
        bridge = self._bridge()
        with self.assertRaises(ValueError):
            unify_fonts(self.input_psd, self.output_psd, bridge=bridge)
        self.assertEqual(bridge.jsx_calls, [])

    def test_unify_fonts_rejects_source_overwrite(self):
        bridge = self._bridge()
        with self.assertRaises(ValueError):
            unify_fonts(
                self.input_psd,
                self.input_psd,
                font_postscript="ArialMT",
                overwrite=True,
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])

    def test_unify_fonts_rejects_existing_output_without_overwrite(self):
        self.output_psd.write_bytes(b"keep")
        bridge = self._bridge()
        with self.assertRaises(FileExistsError):
            unify_fonts(
                self.input_psd,
                self.output_psd,
                size_pt=9,
                bridge=bridge,
            )
        self.assertEqual(self.output_psd.read_bytes(), b"keep")
        self.assertEqual(bridge.jsx_calls, [])

    def test_unify_fonts_validates_font_size_and_color(self):
        bridge = self._bridge()
        with self.assertRaises(ValueError):
            unify_fonts(
                self.input_psd,
                self.output_psd,
                size_pt=0,
                bridge=bridge,
            )
        with self.assertRaises(ValueError):
            unify_fonts(
                self.input_psd,
                self.output_psd,
                color_hex="black",
                bridge=bridge,
            )
        with self.assertRaises(ValueError):
            unify_fonts(
                self.input_psd,
                self.output_psd,
                font_postscript="Arial\tMT",
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])

    def test_unify_fonts_rejects_incomplete_photoshop_response(self):
        bridge = self._bridge(text_count="6", changed_count="5")
        with self.assertRaises(PhotoshopBridgeError) as caught:
            unify_fonts(
                self.input_psd,
                self.output_psd,
                size_pt=9,
                bridge=bridge,
            )
        self.assertEqual(caught.exception.exception_type, "TextLayerCountMismatch")




class FakeExportFigureBridge:
    def __init__(self, *, width=40, height=30, dpi=300.0, mode="RGB"):
        self.width = width
        self.height = height
        self.dpi = dpi
        self.mode = mode
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        import json
        import re
        from PIL import Image

        self.jsx_calls.append(jsx)
        match = re.search(r"var cfg = (\{.*\});", jsx)
        if match is None:
            raise AssertionError("Could not read export config from generated JSX.")
        config = json.loads(match.group(1))
        tiff_path = Path(config["tiff_temp_path"])
        png_path = Path(config["png_temp_path"])

        Image.new("RGB", (self.width, self.height), "white").save(
            tiff_path,
            format="TIFF",
            compression="tiff_lzw",
            dpi=(self.dpi, self.dpi),
        )
        Image.new("RGB", (self.width, self.height), "white").save(
            png_path,
            format="PNG",
            dpi=(self.dpi, self.dpi),
        )

        return (
            "SCI_PS_EXPORT_FIGURE\tOK\tFigure_Export\t" +
            str(self.width) + "\t" + str(self.height) + "\t" +
            str(self.dpi) + "\t" + self.mode + "\t4\t4\t6\t" +
            str(self.width) + "\t" + str(self.height) + "\t" +
            str(self.dpi) + "\tRGB\t" +
            str(self.width) + "\t" + str(self.height) + "\t" +
            str(self.dpi) + "\tRGB\t" +
            str(tiff_path) + "\t" + str(png_path)
        )


class ExportFigureUnitTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Figure_Stage5.psd"
        self.output_dir = self.root / "exports"
        self.input_psd.write_bytes(b"editable-source-psd")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_export_figure_creates_editable_psd_lzw_tiff_and_rgb_png(self):
        from PIL import Image

        source_before = self.input_psd.read_bytes()
        bridge = FakeExportFigureBridge(width=40, height=30, dpi=300)
        result = export_figure(
            self.input_psd,
            self.output_dir,
            basename="Figure_1",
            expected_dpi=300,
            bridge=bridge,
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["width_px"], 40)
        self.assertEqual(result["height_px"], 30)
        self.assertEqual(result["dpi"], 300)
        self.assertEqual(result["mode"], "RGB")
        self.assertTrue(result["editable_psd"]["sha256_matches_input"])
        self.assertTrue(result["input_sha256_unchanged"])
        self.assertEqual(result["panel_group_count"], 4)
        self.assertEqual(result["smart_object_count"], 4)
        self.assertEqual(result["text_layer_count"], 6)

        psd_output = Path(result["editable_psd"]["path"])
        tiff_output = Path(result["tiff"]["path"])
        png_output = Path(result["png"]["path"])
        self.assertEqual(psd_output.read_bytes(), source_before)
        self.assertEqual(self.input_psd.read_bytes(), source_before)
        self.assertEqual(result["tiff"]["compression_tag"], 5)
        self.assertEqual(result["tiff"]["mode"], "RGB")
        self.assertEqual(result["png"]["format"], "PNG")
        self.assertEqual(result["png"]["mode"], "RGB")
        self.assertAlmostEqual(result["png"]["dpi_x"], 300, delta=0.5)

        with Image.open(tiff_output) as image:
            self.assertEqual(image.tag_v2.get(259), 5)
        with Image.open(png_output) as image:
            self.assertEqual(image.mode, "RGB")

        self.assertEqual(len(bridge.jsx_calls), 1)
        jsx = bridge.jsx_calls[0]
        import json
        config_start = jsx.index("var cfg = ") + len("var cfg = ")
        config_end = jsx.index(";", config_start)
        config = json.loads(jsx[config_start:config_end])
        self.assertTrue(Path(config["psd_temp_path"]).name.startswith(".sci_ps_"))
        self.assertTrue(Path(config["tiff_temp_path"]).name.startswith(".sci_tif_"))
        self.assertTrue(Path(config["png_temp_path"]).name.startswith(".sci_png_"))
        self.assertLess(len(Path(config["psd_temp_path"]).name), 64)
        self.assertIn("TIFFEncoding.TIFFLZW", jsx)
        self.assertIn("new TiffSaveOptions()", jsx)
        self.assertIn("new PNGSaveOptions()", jsx)
        self.assertIn("sourceDoc.duplicate", jsx)
        self.assertIn("tiffDoc.flatten()", jsx)
        self.assertIn("pngDoc.flatten()", jsx)
        self.assertIn("SaveOptions.DONOTSAVECHANGES", jsx)

    def test_export_figure_rejects_existing_bundle_without_overwrite(self):
        output_dir = self.output_dir
        output_dir.mkdir()
        existing = output_dir / "Figure_Stage5.tif"
        existing.write_bytes(b"keep")
        bridge = FakeExportFigureBridge()

        with self.assertRaises(FileExistsError):
            export_figure(
                self.input_psd,
                output_dir,
                bridge=bridge,
            )

        self.assertEqual(existing.read_bytes(), b"keep")
        self.assertEqual(bridge.jsx_calls, [])

    def test_export_figure_rejects_source_psd_as_bundle_output(self):
        bridge = FakeExportFigureBridge()
        with self.assertRaises(ValueError):
            export_figure(
                self.input_psd,
                self.root,
                basename=self.input_psd.stem,
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])

    def test_export_figure_rejects_invalid_dpi_and_basename(self):
        bridge = FakeExportFigureBridge()
        with self.assertRaises(ValueError):
            export_figure(
                self.input_psd,
                self.output_dir,
                expected_dpi=450,
                bridge=bridge,
            )
        with self.assertRaises(ValueError):
            export_figure(
                self.input_psd,
                self.output_dir,
                basename="../bad",
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])

    def test_export_figure_rejects_wrong_source_dpi(self):
        bridge = FakeExportFigureBridge(dpi=300)
        with self.assertRaises(PhotoshopBridgeError) as caught:
            export_figure(
                self.input_psd,
                self.output_dir,
                expected_dpi=600,
                bridge=bridge,
            )
        self.assertEqual(caught.exception.operation, "export_figure")
        self.assertEqual(caught.exception.exception_type, "DpiMismatch")

    def test_raster_metadata_inspection_rejects_non_lzw_tiff(self):
        from PIL import Image

        path = self.root / "uncompressed.tif"
        Image.new("RGB", (20, 10), "white").save(
            path,
            format="TIFF",
            compression="raw",
            dpi=(300, 300),
        )
        with self.assertRaises(PhotoshopBridgeError) as caught:
            _inspect_export_raster(
                path,
                expected_format="TIFF",
                expected_size=(20, 10),
                expected_dpi=300,
            )
        self.assertEqual(caught.exception.operation, "export_figure")
        self.assertIn("compression tag", str(caught.exception))



    def test_export_figure_uses_short_temp_names_for_long_scientific_basename(self):
        long_basename = (
            "Figure_Grid_Stage2_Stage3_Labeled_Stage4_ScaleBar_Stage5_FontsUnified"
        )
        bridge = FakeExportFigureBridge(width=40, height=30, dpi=300)
        result = export_figure(
            self.input_psd,
            self.output_dir,
            basename=long_basename,
            expected_dpi=300,
            bridge=bridge,
        )

        import json
        jsx = bridge.jsx_calls[0]
        config_start = jsx.index("var cfg = ") + len("var cfg = ")
        config_end = jsx.index(";", config_start)
        config = json.loads(jsx[config_start:config_end])
        for key in ("psd_temp_path", "tiff_temp_path", "png_temp_path"):
            temp_path = Path(config[key])
            self.assertLess(len(str(temp_path)), 240)
            self.assertNotIn(long_basename, temp_path.name)
        self.assertEqual(Path(result["editable_psd"]["path"]).name, long_basename + ".psd")



class FakePreflightBridge:
    def __init__(self, response):
        self.response = response
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        self.jsx_calls.append(jsx)
        return self.response


def _encode_preflight_record(values):
    from urllib.parse import quote
    return "|".join(quote(str(value), safe="-_.!~*'()") for value in values)


def _make_preflight_response(*, missing_first_smart_object=False):
    rows = [
        _encode_preflight_record([
            "DOC", "Figure_Stage6", 2079, 1583, 300, "RGB",
            "sRGB IEC61966-2.1", 4, 3 if missing_first_smart_object else 4, 5,
        ])
    ]
    positions = {
        "a": (35, 35, 1027, 779),
        "b": (1052, 35, 2044, 779),
        "c": (35, 804, 1027, 1548),
        "d": (1052, 804, 2044, 1548),
    }
    for panel_id, (left, top, right, bottom) in positions.items():
        label_left = left + 8
        label_top = top + 8
        bar_values = (
            [1, 72, 8, 900, 700, 972, 708, 1, 898, 698, 974, 710, 1,
             "10 µm", "ArialMT", 8, "#000000", 930, 670, 972, 680]
            if panel_id == "a"
            else [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", 0, "", 0, 0, 0, 0]
        )
        row = [
            "PANEL",
            panel_id,
            "Group_Panel_" + panel_id,
            "" if missing_first_smart_object and panel_id == "a" else "Image_Panel_" + panel_id,
            0 if missing_first_smart_object and panel_id == "a" else 704,
            0 if missing_first_smart_object and panel_id == "a" else 504,
            left, top, right, bottom,
            704 / (right - left) * 300 if not (missing_first_smart_object and panel_id == "a") else 0,
            504 / (bottom - top) * 300 if not (missing_first_smart_object and panel_id == "a") else 0,
            1, panel_id, "Arial-BoldMT", 10, "#000000",
            label_left, label_top, label_left + 14, label_top + 15,
        ] + bar_values
        rows.append(_encode_preflight_record(row))
        rows.append(_encode_preflight_record([
            "TEXT",
            "Figure_Stage6/Group_Panel_" + panel_id + "/Panel_Label_" + panel_id,
            "Panel_Label_" + panel_id,
            panel_id,
            "Arial-BoldMT",
            10,
            "#000000",
            label_left,
            label_top,
            label_left + 14,
            label_top + 15,
        ]))
    rows.append(_encode_preflight_record([
        "TEXT",
        "Figure_Stage6/Group_Panel_a/Scale_Label_a",
        "Scale_Label_a",
        "10 µm",
        "ArialMT",
        8,
        "#000000",
        930, 670, 972, 680,
    ]))
    return "SCI_PS_PREFLIGHT\tOK\t" + "\n".join(rows)



class FakeInspectBridge:
    def __init__(self, response, callback=None):
        self.response = response
        self.callback = callback
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        self.jsx_calls.append(jsx)
        if self.callback is not None:
            self.callback(jsx)
        return self.response


def _encode_inspection_record(values):
    from urllib.parse import quote
    return "|".join(quote(str(value), safe="-_.!~*'()") for value in values)


def _make_inspection_response():
    rows = [
        _encode_inspection_record(["DOC", "Legacy Figure", 1200, 800, 300, "RGB", "sRGB IEC61966-2.1"]),
        _encode_inspection_record(["COUNT", 5]),
        _encode_inspection_record([
            "LAYER", "root/L0", "root", 0, 0, "ArtLayer", "Title",
            "LayerKind.TEXT", True, 100, "BlendMode.NORMAL", False, False,
            False, False, False, 300, 40, 900, 90, "A figure title",
            "ArialMT", 16, "#000000",
            "Legacy Figure/Title",
        ]),
        _encode_inspection_record([
            "LAYER", "root/G1", "root", 1, 0, "LayerSet", "Old Group",
            "LayerSet", True, 100, "BlendMode.PASS_THROUGH", False, False,
            False, False, False, 100, 100, 700, 700, "", "", "", "",
            "Legacy Figure/Old Group",
        ]),
        _encode_inspection_record([
            "LAYER", "root/G1/L0", "root/G1", 0, 1, "ArtLayer", "Caption copy",
            "LayerKind.TEXT", True, 100, "BlendMode.NORMAL", False, False,
            False, False, False, 120, 120, 400, 170, "Electrochemical data",
            "ArialMT", 8, "#000000",
            "Legacy Figure/Old Group/Caption copy",
        ]),
        _encode_inspection_record([
            "LAYER", "root/L2", "root", 2, 0, "ArtLayer", "Micrograph",
            "LayerKind.SMARTOBJECT", True, 100, "BlendMode.NORMAL", False,
            False, False, False, False, 720, 100, 1180, 700, "", "", "", "",
            "Legacy Figure/Micrograph",
        ]),
        _encode_inspection_record([
            "LAYER", "root/L3", "root", 3, 0, "ArtLayer", "Background",
            "LayerKind.NORMAL", True, 100, "BlendMode.NORMAL", True, True, True,
            True, False, 0, 0, 1200, 800, "", "", "", "",
            "Legacy Figure/Background",
        ]),
    ]
    # The mixed layer collection is front-to-back: the locked Background
    # layer remains at the final index.
    return "SCI_PS_INSPECT_PSD\tOK\t" + "\n".join(rows)

class InspectPsdUnitTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Legacy.psd"
        self.input_psd.write_bytes(b"unchanged-legacy-psd")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_inspect_psd_returns_ordered_nested_layer_tree(self):
        bridge = FakeInspectBridge(_make_inspection_response())
        result = inspect_psd(self.input_psd, bridge=bridge)

        self.assertTrue(result["passed"])
        self.assertEqual(result["layer_count"], 5)
        self.assertEqual(result["group_count"], 1)
        self.assertEqual(result["text_layer_count"], 2)
        self.assertEqual(result["smart_object_count"], 1)
        self.assertEqual(result["layers"][1]["node_key"], "root/G1")
        self.assertEqual(result["layers"][2]["parent_key"], "root/G1")
        self.assertEqual(result["layers"][2]["text"]["contents"], "Electrochemical data")
        self.assertEqual(result["layers"][3]["kind"], "LayerKind.SMARTOBJECT")
        self.assertEqual(len(result["layer_tree_sha256"]), 64)
        self.assertTrue(result["source_sha256_unchanged"])
        self.assertEqual(self.input_psd.read_bytes(), b"unchanged-legacy-psd")
        self.assertIn("container.layers", bridge.jsx_calls[0])
        self.assertIn("app.preferences.rulerUnits = previousRulerUnits", bridge.jsx_calls[0])
        self.assertIn("document_label", bridge.jsx_calls[0])
        self.assertIn("Legacy.psd", bridge.jsx_calls[0])
        self.assertIn("sci_insp_", bridge.jsx_calls[0])
        self.assertNotIn(str(self.input_psd), bridge.jsx_calls[0])
        self.assertEqual(result["layers"][-1]["name"], "Background")
        self.assertTrue(result["layers"][-1]["is_background"])

    def test_inspect_psd_can_write_a_new_json_report(self):
        bridge = FakeInspectBridge(_make_inspection_response())
        report_path = self.root / "reports" / "Legacy_inspection.json"

        result = inspect_psd(self.input_psd, report_path=report_path, bridge=bridge)

        self.assertEqual(result["report_path"], str(report_path))
        saved = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertEqual(saved["source_sha256"], result["source_sha256"])
        self.assertEqual(saved["layer_count"], 5)
        self.assertEqual(saved["report_path"], str(report_path))

    def test_inspect_psd_does_not_overwrite_an_existing_report(self):
        report_path = self.root / "existing.json"
        report_path.write_text("keep", encoding="utf-8")
        bridge = FakeInspectBridge(_make_inspection_response())

        with self.assertRaises(FileExistsError):
            inspect_psd(self.input_psd, report_path=report_path, bridge=bridge)

        self.assertEqual(report_path.read_text(encoding="utf-8"), "keep")
        self.assertEqual(bridge.jsx_calls, [])

    def test_inspect_psd_rejects_malformed_layer_record_count(self):
        response = (
            "SCI_PS_INSPECT_PSD\tOK\t"
            + _encode_inspection_record(["DOC", "Bad", 100, 100, 300, "RGB", ""])
            + "\n"
            + _encode_inspection_record(["COUNT", 2])
        )
        bridge = FakeInspectBridge(response)
        with self.assertRaises(ValueError):
            inspect_psd(self.input_psd, bridge=bridge)




class FakePreviewBridge:
    def __init__(self, inspection_response=None):
        self.inspection_response = inspection_response or _make_inspection_response()
        self.jsx_calls = []

    def execute_jsx(self, jsx):
        self.jsx_calls.append(jsx)
        if "SCI_PS_INSPECT_PSD" in jsx:
            return self.inspection_response

        import re
        match = re.search(r"var cfg = (\{.*\});", jsx)
        if not match:
            raise AssertionError("Preview JSX configuration was not embedded.")
        config = json.loads(match.group(1))
        from PIL import Image

        Image.new("RGBA", (128, 96), (245, 245, 245, 255)).save(
            Path(config["output_dir"]) / "composite.png",
            format="PNG",
        )
        records = []
        for item in config["items"]:
            preview_path = Path(config["output_dir"]) / item["file_name"]
            Image.new("RGBA", (64, 48), (40, 90, 140, 220)).save(preview_path, format="PNG")
            records.append(_encode_inspection_record([
                "PREVIEW", item["node_key"], item["file_name"], "OK", 64, 48, ""
            ]))
        return (
            "SCI_PS_PREVIEW_PSD\tOK\t"
            + str(len(config["items"]))
            + "\t"
            + "\n".join(records)
        )


class PreviewPsdLayersTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Legacy.psd"
        self.input_psd.write_bytes(b"unchanged-legacy-psd")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_preview_psd_layers_creates_contact_sheet_and_manifest(self):
        bridge = FakePreviewBridge()
        output_dir = self.root / "layer_previews"

        result = preview_psd_layers(
            self.input_psd,
            output_dir,
            bridge=bridge,
            thumbnail_width_px=256,
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["preview_count"], 4)
        self.assertEqual(result["successful_count"], 4)
        self.assertEqual(result["failed_count"], 0)
        self.assertEqual(result["skipped_backgrounds"], ["root/L3"])
        self.assertEqual(result["inspection_layer_count"], 5)
        self.assertEqual(result["all_node_count"], 5)
        self.assertEqual(result["eligible_non_background_node_count"], 4)
        self.assertEqual(result["previewed_non_background_node_count"], 4)
        self.assertTrue(result["full_visual_coverage"])
        self.assertEqual(result["missing_preview_node_keys"], [])
        self.assertEqual(result["failed_preview_node_keys"], [])
        self.assertTrue(Path(result["contact_sheet"]["path"]).is_file())
        self.assertTrue(Path(result["composite_preview"]).is_file())
        self.assertTrue(Path(result["report_path"]).is_file())
        self.assertEqual(self.input_psd.read_bytes(), b"unchanged-legacy-psd")
        self.assertTrue(result["source_sha256_unchanged"])
        self.assertIn("restoreGroupChildren", bridge.jsx_calls[-1])
        self.assertIn('sourceDoc.duplicate("SCI_PSD_COMPOSITE_PREVIEW"', bridge.jsx_calls[-1])
        self.assertIn('sourceDoc.duplicate(\n                    "SCI_Layer_Preview_"', bridge.jsx_calls[-1])
        self.assertTrue(Path(result["composite_preview"]).is_file())
        self.assertIn("ResampleMethod.BICUBICSHARPER", bridge.jsx_calls[-1])
        self.assertIn("SaveOptions.DONOTSAVECHANGES", bridge.jsx_calls[-1])

    def test_preview_psd_layers_can_select_specific_node_keys(self):
        bridge = FakePreviewBridge()
        output_dir = self.root / "selected_previews"

        result = preview_psd_layers(
            self.input_psd,
            output_dir,
            node_keys=["root/G1/L0", "root/L2"],
            bridge=bridge,
        )

        self.assertEqual(result["preview_count"], 2)
        self.assertEqual(
            [item["node_key"] for item in result["previews"]],
            ["root/G1/L0", "root/L2"],
        )
        self.assertFalse(result["full_visual_coverage"])
        self.assertEqual(
            result["missing_preview_node_keys"],
            ["root/G1", "root/L0"],
        )
        self.assertEqual(result["previewed_non_background_node_count"], 2)

    def test_preview_psd_layers_rejects_unknown_or_duplicate_node_keys(self):
        for keys in (["root/not-found"], ["root/L2", "root/L2"]):
            bridge = FakePreviewBridge()
            output_dir = self.root / ("bad_" + str(len(list(self.root.iterdir()))))
            with self.assertRaises(ValueError):
                preview_psd_layers(
                    self.input_psd,
                    output_dir,
                    node_keys=keys,
                    bridge=bridge,
                )

    def test_preview_psd_layers_rejects_existing_output_directory(self):
        output_dir = self.root / "existing"
        output_dir.mkdir()
        bridge = FakePreviewBridge()
        with self.assertRaises(FileExistsError):
            preview_psd_layers(self.input_psd, output_dir, bridge=bridge)
        self.assertEqual(bridge.jsx_calls, [])

    def test_preview_jsx_contains_isolation_and_safe_duplicate_workflow(self):
        jsx = _build_preview_psd_layers_jsx({
            "input_path": "C:/Temp/source.psd",
            "output_dir": "C:/Temp/previews",
            "thumbnail_width_px": 512,
            "items": [
                {
                    "sequence": 1,
                    "node_key": "root/G0/L1",
                    "stack_index": 1,
                    "file_name": "layer_001.png",
                }
            ],
        })
        self.assertIn("hideAll(previewDoc)", jsx)
        self.assertIn("setAncestorsVisible(item, previewMap)", jsx)
        self.assertIn("restoreGroupChildren(sourceLayer, previewLayer)", jsx)
        self.assertIn("Boolean(sourceLayer.grouped)", jsx)
        self.assertIn("PNGSaveOptions", jsx)
        self.assertIn("previewDoc.resizeImage", jsx)
        self.assertIn("previousRulerUnits", jsx)
        self.assertIn("previousTypeUnits", jsx)

    def test_preview_record_parser_rejects_duplicate_or_unknown_status(self):
        good = _encode_inspection_record([
            "PREVIEW", "root/L0", "layer_001.png", "OK", 64, 48, ""
        ])
        duplicate = good + "\n" + good
        with self.assertRaisesRegex(ValueError, "Duplicate PSD preview node key"):
            _parse_psd_preview_records(duplicate)

        bad_status = _encode_inspection_record([
            "PREVIEW", "root/L0", "layer_001.png", "MAYBE", 64, 48, ""
        ])
        with self.assertRaisesRegex(ValueError, "Unknown PSD preview status"):
            _parse_psd_preview_records(bad_status)




def _make_valid_classification(inspection, use_visual_evidence=False):
    classifications = []
    for layer in inspection["layers"]:
        if layer["typename"] == "LayerSet":
            role = "panel_group" if layer["name"].startswith("Group_Panel_") else "unknown"
            evidence = [{
                "type": "layer_name",
                "observation": layer["name"],
            }]
        elif layer["kind"] == "LayerKind.TEXT":
            role = "title" if layer["name"] == "Title" else "annotation"
            evidence = [{
                "type": "text_content",
                "observation": layer["text"]["contents"],
            }]
        elif layer["kind"] == "LayerKind.SMARTOBJECT":
            role = "panel_image"
            if use_visual_evidence:
                evidence = [{
                    "type": "visual_preview",
                    "preview_node_key": layer["node_key"],
                    "observation": "The isolated preview shows the embedded panel image.",
                }]
            else:
                evidence = [{
                    "type": "layer_type",
                    "observation": layer["kind"],
                }]
        else:
            role = "unknown"
            evidence = [{
                "type": "layer_type",
                "observation": layer["kind"],
            }]
        classifications.append({
            "node_key": layer["node_key"],
            "role": role,
            "confidence": 0.92 if role != "unknown" else 0.35,
            "rationale": "Classification is grounded in the recorded layer evidence.",
            "evidence": evidence,
        })
    return {
        "schema_version": 1,
        "source_sha256": inspection["source_sha256"],
        "layer_tree_sha256": inspection["layer_tree_sha256"],
        "classifications": classifications,
    }


class PsdSemanticClassificationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Legacy.psd"
        self.input_psd.write_bytes(b"unchanged-legacy-psd")
        self.inspection = inspect_psd(
            self.input_psd,
            bridge=FakeInspectBridge(_make_inspection_response()),
        )
        self.inspection_path = self.root / "inspection.json"
        self.inspection_path.write_text(
            json.dumps(self.inspection, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.classification = _make_valid_classification(self.inspection)
        self.classification_path = self.root / "classification.json"
        self.classification_path.write_text(
            json.dumps(self.classification, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_layer_name_and_layer_type_evidence_must_match_inspection(self):
        wrong_name = _make_valid_classification(self.inspection)
        named_item = next(
            item for item in wrong_name["classifications"]
            if item["evidence"][0]["type"] == "layer_name"
        )
        named_item["evidence"][0]["observation"] = "Invented group name"
        with self.assertRaisesRegex(ValueError, "exact inspected layer name"):
            validate_psd_classification(self.input_psd, self.inspection_path, wrong_name)

        wrong_type = _make_valid_classification(self.inspection)
        smart = next(
            item for item in wrong_type["classifications"]
            if item["role"] == "panel_image"
        )
        smart["evidence"] = [{"type": "layer_type", "observation": "LayerKind.TEXT"}]
        with self.assertRaisesRegex(ValueError, "does not match"):
            validate_psd_classification(self.input_psd, self.inspection_path, wrong_type)

    def test_text_and_geometry_evidence_must_match_inspected_values(self):
        wrong_text = _make_valid_classification(self.inspection)
        text_item = next(
            item for item in wrong_text["classifications"]
            if item["role"] == "title"
        )
        text_item["evidence"][0]["observation"] = "Text that is not in the layer"
        with self.assertRaisesRegex(ValueError, "quote text actually present"):
            validate_psd_classification(self.input_psd, self.inspection_path, wrong_text)

        wrong_geometry = _make_valid_classification(self.inspection)
        group = next(
            item for item in wrong_geometry["classifications"]
            if item["role"] == "unknown"
        )
        group["evidence"] = [{
            "type": "geometry",
            "observation": "The bounds identify the proposed panel.",
            "bounds_px": [0, 0, 1, 1],
        }]
        with self.assertRaisesRegex(ValueError, "bounds_px do not match"):
            validate_psd_classification(self.input_psd, self.inspection_path, wrong_geometry)

    def test_validate_classification_checks_every_node_and_evidence(self):
        result = validate_psd_classification(
            self.input_psd,
            self.inspection_path,
            self.classification_path,
        )
        self.assertTrue(result["passed"])
        self.assertEqual(result["layer_count"], 5)
        self.assertEqual(result["classified_count"], 5)
        self.assertEqual(result["role_counts"]["unknown"], 2)
        self.assertEqual(result["role_counts"]["title"], 1)
        self.assertEqual(len(result["classification_sha256"]), 64)
        self.assertTrue(result["source_sha256"] == self.inspection["source_sha256"])

    def test_validate_classification_checks_visual_evidence_against_preview_manifest(self):
        preview_dir = self.root / "previews"
        preview_dir.mkdir()
        image_path = preview_dir / "layer_001.png"
        image_path.write_bytes(b"fake-png")
        composite_path = preview_dir / "composite.png"
        composite_path.write_bytes(b"fake-composite")
        contact_path = preview_dir / "layer_contact_sheet.png"
        contact_path.write_bytes(b"fake-contact")
        layer = next(x for x in self.inspection["layers"] if x["kind"] == "LayerKind.SMARTOBJECT")
        preview_report = {
            "schema_version": 1,
            "source_sha256": self.inspection["source_sha256"],
            "layer_tree_sha256": self.inspection["layer_tree_sha256"],
            "preview_dir": str(preview_dir),
            "composite_preview": str(composite_path),
            "contact_sheet": {"path": str(contact_path)},
            "previews": [{
                "node_key": layer["node_key"],
                "file_name": "layer_001.png",
                "status": "OK",
            }],
        }
        preview_path = self.root / "preview_report.json"
        preview_path.write_text(json.dumps(preview_report), encoding="utf-8")
        visual_classification = _make_valid_classification(
            self.inspection, use_visual_evidence=True
        )
        result = validate_psd_classification(
            self.input_psd,
            self.inspection_path,
            visual_classification,
            preview_report=preview_path,
        )
        self.assertTrue(result["passed"])

    def test_validate_classification_rejects_stale_hashes_and_missing_nodes(self):
        stale = json.loads(json.dumps(self.classification))
        stale["source_sha256"] = "0" * 64
        stale_path = self.root / "stale.json"
        stale_path.write_text(json.dumps(stale), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "source_sha256"):
            validate_psd_classification(self.input_psd, self.inspection_path, stale_path)

        missing = json.loads(json.dumps(self.classification))
        missing["classifications"].pop()
        missing_path = self.root / "missing.json"
        missing_path.write_text(json.dumps(missing), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "every inspected node"):
            validate_psd_classification(self.input_psd, self.inspection_path, missing_path)

    def test_validate_classification_rejects_unknown_nodes_duplicate_roles_and_bad_confidence(self):
        cases = []
        unknown = json.loads(json.dumps(self.classification))
        unknown["classifications"][0]["node_key"] = "root/not-real"
        cases.append((unknown, "unknown node_key"))
        duplicate = json.loads(json.dumps(self.classification))
        duplicate["classifications"][1]["node_key"] = duplicate["classifications"][0]["node_key"]
        cases.append((duplicate, "Duplicate classification"))
        bad_confidence = json.loads(json.dumps(self.classification))
        bad_confidence["classifications"][0]["confidence"] = 1.5
        cases.append((bad_confidence, "confidence"))
        for index, (item, message) in enumerate(cases):
            path = self.root / ("invalid_" + str(index) + ".json")
            path.write_text(json.dumps(item), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, message):
                validate_psd_classification(self.input_psd, self.inspection_path, path)

    def test_visual_evidence_requires_a_successful_matching_preview(self):
        item = _make_valid_classification(self.inspection, use_visual_evidence=True)
        with self.assertRaisesRegex(ValueError, "no successful preview"):
            validate_psd_classification(
                self.input_psd,
                self.inspection_path,
                item,
                preview_report={
                    "schema_version": 1,
                    "source_sha256": self.inspection["source_sha256"],
                    "layer_tree_sha256": self.inspection["layer_tree_sha256"],
                    "preview_dir": str(self.root),
                    "previews": [],
                },
            )

    def test_text_evidence_must_reference_a_native_text_layer(self):
        item = json.loads(json.dumps(self.classification))
        smart = next(x for x in item["classifications"] if x["role"] == "panel_image")
        smart["evidence"] = [{
            "type": "text_content",
            "observation": "This is not a text layer.",
        }]
        with self.assertRaisesRegex(ValueError, "native text layer"):
            validate_psd_classification(self.input_psd, self.inspection_path, item)

    def test_plan_draft_requires_valid_classification_and_operations_and_is_unapproved(self):
        output = self.root / "organize_plan_draft.json"
        operations = [{
            "action": "create_group",
            "key": "review_group",
            "name": "Review_Group",
            "parent": "root",
        }]
        result = create_organize_plan_draft(
            self.input_psd,
            self.inspection_path,
            self.classification_path,
            operations,
            output,
        )
        self.assertTrue(result["passed"])
        self.assertTrue(result["requires_user_approval"])
        self.assertFalse(result["approved"])
        saved = json.loads(output.read_text(encoding="utf-8"))
        self.assertTrue(saved["requires_user_approval"])
        self.assertFalse(saved["approved"])
        self.assertEqual(saved["operations"][0]["action"], "create_group")
        self.assertEqual(saved["source_sha256"], self.inspection["source_sha256"])

    def test_plan_draft_rejects_unsafe_operations_and_existing_output(self):
        output = self.root / "unsafe_plan.json"
        unsafe = [{
            "action": "delete_layer",
            "target": "root/G1",
        }]
        with self.assertRaisesRegex(ValueError, "Unsupported organizer action"):
            create_organize_plan_draft(
                self.input_psd,
                self.inspection_path,
                self.classification_path,
                unsafe,
                output,
            )
        output.write_text("preserve", encoding="utf-8")
        with self.assertRaises(FileExistsError):
            create_organize_plan_draft(
                self.input_psd,
                self.inspection_path,
                self.classification_path,
                [{"action": "create_group", "key": "g", "name": "G", "parent": "root"}],
                output,
            )
        self.assertEqual(output.read_text(encoding="utf-8"), "preserve")



class OrganizePlanValidationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Legacy.psd"
        self.input_psd.write_bytes(b"unchanged-legacy-psd")
        self.inspection = inspect_psd(
            self.input_psd,
            bridge=FakeInspectBridge(_make_inspection_response()),
        )

    def tearDown(self):
        self.temp_dir.cleanup()

    def plan(self, operations):
        return {
            "schema_version": 1,
            "source_sha256": self.inspection["source_sha256"],
            "layer_tree_sha256": self.inspection["layer_tree_sha256"],
            "operations": operations,
        }

    def test_valid_plan_can_create_group_rename_move_reorder_and_set_visibility(self):
        plan = self.plan([
            {"action": "create_group", "key": "panel_a", "name": "Panel A", "parent": "root"},
            {
                "action": "rename_layer",
                "target": "root/G1",
                "expected_name": "Old Group",
                "new_name": "Legacy Source Group",
            },
            {
                "action": "move_layer",
                "target": "root/G1/L0",
                "expected_name": "Caption copy",
                "destination": "panel_a",
                "position": "end",
            },
            {
                "action": "reorder_layer",
                "target": "root/L2",
                "expected_name": "Micrograph",
                "relative_to": "root/L0",
                "relative_expected_name": "Title",
                "position": "before",
            },
            {
                "action": "set_visibility",
                "target": "root/L2",
                "expected_name": "Micrograph",
                "visible": False,
            },
        ])

        normalized = _validate_organize_plan(plan, self.inspection)

        self.assertEqual(len(normalized["operations"]), 5)
        self.assertIn("panel_a", normalized["created_groups"])
        identities = {item["identity"] for item in normalized["verify_refs"]}
        self.assertIn("group:panel_a", identities)
        self.assertIn("source:root/G1/L0", identities)

    def test_plan_rejects_stale_source_hash_or_layer_tree(self):
        stale_hash = self.plan([
            {"action": "rename_layer", "target": "root/G1", "expected_name": "Old Group", "new_name": "New"}
        ])
        stale_hash["source_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValueError, "source_sha256"):
            _validate_organize_plan(stale_hash, self.inspection)

        stale_tree = self.plan([
            {"action": "rename_layer", "target": "root/G1", "expected_name": "Old Group", "new_name": "New"}
        ])
        stale_tree["layer_tree_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValueError, "layer_tree_sha256"):
            _validate_organize_plan(stale_tree, self.inspection)

    def test_plan_rejects_unapproved_or_destructive_operation_types(self):
        with self.assertRaises(PermissionError):
            apply_organize_plan(
                self.input_psd,
                self.plan([{"action": "delete_layer", "target": "root/G1"}]),
                self.root / "out.psd",
                approved=False,
                bridge=FakeInspectBridge(_make_inspection_response()),
            )

        with self.assertRaisesRegex(ValueError, "Unsupported organizer action"):
            _validate_organize_plan(
                self.plan([{"action": "delete_layer", "target": "root/G1"}]),
                self.inspection,
            )

    def test_plan_rejects_moving_background_locked_or_clipping_layers(self):
        plan = self.plan([
            {"action": "move_layer", "target": "root/L3", "expected_name": "Background", "destination": "root/G1"}
        ])
        with self.assertRaisesRegex(ValueError, "background"):
            _validate_organize_plan(plan, self.inspection)

        clipped = json.loads(json.dumps(self.inspection))
        layer = next(item for item in clipped["layers"] if item["node_key"] == "root/L2")
        layer["clipped_to_below"] = True
        plan = self.plan([
            {"action": "move_layer", "target": "root/L2", "expected_name": "Micrograph", "destination": "root/G1"}
        ])
        with self.assertRaisesRegex(ValueError, "clipping-mask chains"):
            _validate_organize_plan(plan, clipped)

    def test_plan_rejects_self_or_descendant_group_moves_and_sibling_mismatch(self):
        self_move = self.plan([
            {"action": "move_layer", "target": "root/G1", "expected_name": "Old Group", "destination": "root/G1"}
        ])
        with self.assertRaisesRegex(ValueError, "itself or one of its descendants"):
            _validate_organize_plan(self_move, self.inspection)

        bad_reorder = self.plan([
            {
                "action": "reorder_layer",
                "target": "root/G1/L0",
                "expected_name": "Caption copy",
                "relative_to": "root/L2",
                "relative_expected_name": "Micrograph",
                "position": "before",
            }
        ])
        with self.assertRaisesRegex(ValueError, "siblings"):
            _validate_organize_plan(bad_reorder, self.inspection)

    def test_plan_rejects_combined_group_moves_that_create_a_parent_cycle(self):
        inspection = json.loads(json.dumps(self.inspection))
        inspection["layers"].append({
            "node_key": "root/G3",
            "parent_key": "root",
            "stack_index": 4,
            "depth": 0,
            "typename": "LayerSet",
            "name": "Other Group",
            "kind": "LayerSet",
            "visible": True,
            "opacity": 100.0,
            "blend_mode": "BlendMode.PASS_THROUGH",
            "is_background": False,
            "all_locked": False,
            "position_locked": False,
            "pixels_locked": False,
            "clipped_to_below": False,
            "bounds_px": [0, 0, 10, 10],
            "text": None,
            "path": "Legacy Figure/Other Group",
        })
        inspection["layer_count"] += 1
        plan = self.plan([
            {
                "action": "move_layer",
                "target": "root/G1",
                "expected_name": "Old Group",
                "destination": "root/G3",
            },
            {
                "action": "move_layer",
                "target": "root/G3",
                "expected_name": "Other Group",
                "destination": "root/G1",
            },
        ])
        with self.assertRaisesRegex(ValueError, "parent cycle"):
            _validate_organize_plan(plan, inspection)

    def test_create_groups_must_come_before_layer_operations(self):
        plan = self.plan([
            {"action": "rename_layer", "target": "root/G1", "expected_name": "Old Group", "new_name": "Renamed"},
            {"action": "create_group", "key": "new_group", "name": "New Group", "parent": "root"},
        ])
        with self.assertRaisesRegex(ValueError, "must appear before"):
            _validate_organize_plan(plan, self.inspection)

    def test_plan_rejects_moving_a_layer_inside_a_locked_ancestor_group(self):
        inspection = json.loads(json.dumps(self.inspection))
        group = next(layer for layer in inspection["layers"] if layer["node_key"] == "root/G1")
        group["all_locked"] = True
        plan = self.plan([
            {"action": "create_group", "key": "new_group", "name": "New", "parent": "root"},
            {
                "action": "move_layer",
                "target": "root/G1/L0",
                "expected_name": "Caption copy",
                "destination": "new_group",
                "position": "end",
            },
        ])
        with self.assertRaisesRegex(ValueError, "child of a locked group"):
            _validate_organize_plan(plan, inspection)

    def test_group_parent_rejects_position_locked_groups(self):
        inspection = json.loads(json.dumps(self.inspection))
        group = next(layer for layer in inspection["layers"] if layer["node_key"] == "root/G1")
        group["position_locked"] = True
        plan = self.plan([
            {"action": "create_group", "key": "new_group", "name": "Nested", "parent": "root/G1"}
        ])
        with self.assertRaisesRegex(ValueError, "position-locked LayerSet"):
            _validate_organize_plan(plan, inspection)

    def test_apply_organize_jsx_handles_background_and_restores_preferences(self):
        jsx = _build_apply_organize_plan_jsx({
            "temp_input_path": "C:/Temp/source.psd",
            "staged_output_path": "C:/Temp/output.psd",
            "operations": [
                {"action": "create_group", "key": "g", "name": "Group", "parent": "root"}
            ],
            "verify_refs": [],
        })
        self.assertIn("bottom.isBackgroundLayer", jsx)
        self.assertIn("ElementPlacement.PLACEBEFORE", jsx)
        self.assertIn("moveToEnd(group, parent)", jsx)
        self.assertIn("app.preferences.rulerUnits = previousRulerUnits", jsx)
        self.assertIn("app.preferences.typeUnits = previousTypeUnits", jsx)
        self.assertIn("doc.saveAs(outputFile", jsx)
        self.assertIn("SaveOptions.DONOTSAVECHANGES", jsx)


    def test_organizer_verification_records_are_decoded(self):
        encoded = _encode_inspection_record([
            "VERIFY", "source:root/L0", "root/G1/L0", "root/G1", 0,
            "ArtLayer", "Caption", True,
        ])
        verified = _parse_organize_verify_records(encoded)
        self.assertEqual(verified["source:root/L0"]["name"], "Caption")
        self.assertEqual(verified["source:root/L0"]["parent_key"], "root/G1")



class PreflightFigureUnitTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.input_psd = self.root / "Figure_Stage5.psd"
        self.export_dir = self.root / "exports"
        self.input_psd.write_bytes(b"unchanged-stage5-psd")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_preflight_checks_panels_scale_calibration_and_export_bundle(self):
        bridge = FakePreflightBridge(_make_preflight_response())
        self.export_dir.mkdir()
        (self.export_dir / "Figure_Stage5.psd").write_bytes(self.input_psd.read_bytes())
        Image.new("RGB", (2079, 1583), "white").save(
            self.export_dir / "Figure_Stage5.tif",
            format="TIFF",
            compression="tiff_lzw",
            dpi=(300, 300),
        )
        Image.new("RGB", (2079, 1583), "white").save(
            self.export_dir / "Figure_Stage5.png",
            format="PNG",
            dpi=(300, 300),
        )

        result = preflight_figure(
            self.input_psd,
            expected_dpi=300,
            expected_width_cm=17.6,
            calibrations={
                "a": {
                    "physical_length": 10,
                    "pixel_size": 0.2,
                    "unit": "µm",
                    "tolerance_px": 2.0,
                }
            },
            export_dir=self.export_dir,
            bridge=bridge,
        )

        self.assertTrue(result["passed"])
        self.assertTrue(result["input_sha256_unchanged"])
        self.assertEqual(result["document"]["panel_group_count"], 4)
        self.assertEqual(len(result["panels"]), 4)
        self.assertEqual(len(result["text_layers"]), 5)
        checks = {item["id"]: item for item in result["checks"]}
        self.assertEqual(checks["canvas.dpi"]["status"], "PASS")
        self.assertEqual(checks["canvas.width"]["status"], "PASS")
        self.assertEqual(checks["panel_a.scale_calibration_width"]["status"], "PASS")
        self.assertEqual(checks["panel_a.scale_calibration_label"]["status"], "PASS")
        self.assertEqual(checks["exports.editable_psd"]["status"], "PASS")
        self.assertEqual(checks["exports.tiff"]["status"], "PASS")
        self.assertEqual(checks["exports.png"]["status"], "PASS")
        self.assertEqual(checks["panel_a.effective_ppi"]["status"], "WARN")
        self.assertGreater(result["warning_count"], 0)
        self.assertIn("placedLayerEditContents", bridge.jsx_calls[0])
        self.assertIn("SCI_PS_PREFLIGHT", bridge.jsx_calls[0])
        self.assertEqual(self.input_psd.read_bytes(), b"unchanged-stage5-psd")

    def test_preflight_jsx_restores_global_photoshop_preferences(self):
        bridge = FakePreflightBridge(_make_preflight_response())
        preflight_figure(self.input_psd, bridge=bridge)

        jsx = bridge.jsx_calls[0]
        self.assertIn("var previousRulerUnits = null;", jsx)
        self.assertIn("var previousTypeUnits = null;", jsx)
        self.assertIn("previousRulerUnits = app.preferences.rulerUnits;", jsx)
        self.assertIn("previousTypeUnits = app.preferences.typeUnits;", jsx)
        self.assertLess(
            jsx.index("previousRulerUnits = app.preferences.rulerUnits;"),
            jsx.index("app.preferences.rulerUnits = Units.PIXELS;"),
        )
        self.assertLess(
            jsx.index("previousTypeUnits = app.preferences.typeUnits;"),
            jsx.index("app.preferences.typeUnits = TypeUnits.POINTS;"),
        )
        self.assertIn("if (rulerUnitsCaptured)", jsx)
        self.assertIn("app.preferences.rulerUnits = previousRulerUnits;", jsx)
        self.assertIn("if (typeUnitsCaptured)", jsx)
        self.assertIn("app.preferences.typeUnits = previousTypeUnits;", jsx)
        self.assertIn("app.displayDialogs = previousDialogs;", jsx)


    def test_preflight_strict_promotes_publication_warnings_to_failures(self):
        bridge = FakePreflightBridge(_make_preflight_response())
        result = preflight_figure(
            self.input_psd,
            strict=True,
            bridge=bridge,
        )
        self.assertFalse(result["passed"])
        self.assertGreater(result["failure_count"], 0)
        self.assertEqual(result["warning_count"], 0)

    def test_preflight_fails_when_a_panel_smart_object_is_missing(self):
        bridge = FakePreflightBridge(
            _make_preflight_response(missing_first_smart_object=True)
        )
        result = preflight_figure(self.input_psd, bridge=bridge)
        self.assertFalse(result["passed"])
        checks = {item["id"]: item for item in result["checks"]}
        self.assertEqual(checks["structure.smart_objects"]["status"], "FAIL")
        self.assertEqual(checks["panel_a.smart_object"]["status"], "FAIL")

    def test_preflight_rejects_invalid_calibration_manifest(self):
        bridge = FakePreflightBridge(_make_preflight_response())
        with self.assertRaises(ValueError):
            preflight_figure(
                self.input_psd,
                calibrations={"a": {"physical_length": 10, "pixel_size": 0, "unit": "µm"}},
                bridge=bridge,
            )
        self.assertEqual(bridge.jsx_calls, [])

    def test_preflight_requires_a_panel_group_and_rgb_300_or_600_dpi(self):
        response = (
            "SCI_PS_PREFLIGHT\tOK\t" +
            _encode_preflight_record([
                "DOC", "BadFigure", 100, 100, 150, "CMYK", "", 0, 0, 0
            ])
        )
        bridge = FakePreflightBridge(response)
        result = preflight_figure(self.input_psd, bridge=bridge)
        self.assertFalse(result["passed"])
        checks = {item["id"]: item for item in result["checks"]}
        self.assertEqual(checks["canvas.dpi"]["status"], "FAIL")
        self.assertEqual(checks["canvas.color_mode"]["status"], "FAIL")
        self.assertEqual(checks["structure.panel_groups"]["status"], "FAIL")





if __name__ == "__main__":
    unittest.main()

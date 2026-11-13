from __future__ import annotations

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





class PhotoshopConnectionUnitTests(unittest.TestCase):
    def test_injected_application_records_backend(self):
        bridge = PhotoshopBridge(app=FakePhotoshopApplication())
        self.assertEqual(bridge.backend, "injected")
        self.assertTrue(bridge.connection_report["injected"])

    def test_silent_dialog_mode_wraps_jsx_and_restores_previous_setting(self):
        app = FakePhotoshopApplication(["result"])
        bridge = PhotoshopBridge(app=app)
        bridge.silent_dialogs = True

        self.assertEqual(bridge.execute_jsx("(function(){return 'result';})();"), "result")
        payload = app.calls[0]
        self.assertIn("app.displayDialogs=DialogModes.NO", payload)
        self.assertIn("finally{app.displayDialogs=__sciPreviousDialogs;}", payload)
        self.assertIn("return eval(", payload)

    def test_silent_dialog_mode_is_opt_in(self):
        app = FakePhotoshopApplication(["result"])
        bridge = PhotoshopBridge(app=app)

        bridge.execute_jsx("app.version;")
        self.assertEqual(app.calls, ["app.version;"])


if __name__ == "__main__":
    unittest.main()

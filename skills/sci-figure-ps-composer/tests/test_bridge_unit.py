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
    _panel_label_text,
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



if __name__ == "__main__":
    unittest.main()

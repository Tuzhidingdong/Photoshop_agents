from __future__ import annotations

import sys
import unittest
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from ps_sci_toolkit import PhotoshopBridge, PhotoshopBridgeError


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


if __name__ == "__main__":
    unittest.main()

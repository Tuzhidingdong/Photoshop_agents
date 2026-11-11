#!/usr/bin/env python
"""Photoshop COM/JSX bridge for sci-figure-ps-composer.

This module deliberately keeps Photoshop as the native document engine.
It provides connection, JSX execution, preference locking, and diagnostics.
"""

from __future__ import annotations

import argparse
import importlib.metadata
import hashlib
import json
import os
import platform
import shutil
import sys
import uuid
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import unquote


PROTOCOL_PREFIX = "SCI_PS_"


class PhotoshopBridgeError(RuntimeError):
    """A normalized failure raised by the Photoshop bridge."""

    def __init__(
        self,
        operation: str,
        message: str,
        *,
        exception_type: Optional[str] = None,
        jsx_preview: Optional[str] = None,
    ) -> None:
        super().__init__(message)
        self.operation = operation
        self.exception_type = exception_type
        self.jsx_preview = jsx_preview

    def to_dict(self) -> Dict[str, Any]:
        return {
            "operation": self.operation,
            "exception_type": self.exception_type,
            "message": str(self),
            "jsx_preview": self.jsx_preview,
        }


def _package_version(distribution: str) -> Optional[str]:
    try:
        return importlib.metadata.version(distribution)
    except importlib.metadata.PackageNotFoundError:
        return None


def _clean_jsx_preview(jsx: str, limit: int = 500) -> str:
    preview = " ".join(jsx.strip().split())
    if len(preview) > limit:
        return preview[:limit] + "..."
    return preview


def _parse_protocol(
    result: Any,
    expected_marker: str,
    minimum_fields: int,
    operation: str,
    jsx: str,
) -> list[str]:
    # Preserve trailing tab delimiters: an empty active-document name is a
    # valid final field when Photoshop has no open documents.
    raw = "" if result is None else str(result).rstrip("\r\n")
    fields = raw.split("\t")

    if not fields or fields[0] != expected_marker:
        raise PhotoshopBridgeError(
            operation,
            "Photoshop returned an invalid bridge response: " + repr(raw),
            exception_type="InvalidBridgeResponse",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    # Error responses intentionally contain fewer fields than success
    # responses, so inspect the status before enforcing the success schema.
    if len(fields) > 1 and fields[1] == "ERROR":
        name = fields[2] if len(fields) > 2 else "ExtendScriptError"
        message = fields[3] if len(fields) > 3 else "Unknown JSX error"
        line = fields[4] if len(fields) > 4 else ""
        raise PhotoshopBridgeError(
            operation,
            "ExtendScript " + name + ": " + message +
            ((" (line " + line + ")") if line else ""),
            exception_type="ExtendScriptError",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    if len(fields) < minimum_fields:
        raise PhotoshopBridgeError(
            operation,
            "Photoshop returned an incomplete bridge response: " + repr(raw),
            exception_type="InvalidBridgeResponse",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    if len(fields) > 1 and fields[1] != "OK":
        raise PhotoshopBridgeError(
            operation,
            "Photoshop bridge reported an unexpected status: " + repr(raw),
            exception_type="InvalidBridgeStatus",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    return fields


class PhotoshopBridge:
    """Thin, synchronous bridge to a running desktop Photoshop instance."""

    def __init__(
        self,
        app: Any = None,
        *,
        execution_mode: str = "quiet",
    ) -> None:
        if execution_mode not in ("quiet", "interactive"):
            raise ValueError("execution_mode must be 'quiet' or 'interactive'.")
        self.execution_mode = execution_mode
        self.app = app if app is not None else self._connect_application()

    @staticmethod
    def _connect_application() -> Any:
        if os.name != "nt":
            raise PhotoshopBridgeError(
                "connect",
                "Photoshop COM automation is supported by this toolkit on Windows only.",
                exception_type="UnsupportedPlatform",
            )

        try:
            import photoshop.api as photoshop_api
        except Exception as exc:
            raise PhotoshopBridgeError(
                "connect",
                "Could not import photoshop.api. Install photoshop-python-api "
                "into this Python interpreter. Original error: " + str(exc),
                exception_type=type(exc).__name__,
            ) from exc

        try:
            return photoshop_api.Application()
        except Exception as exc:
            raise PhotoshopBridgeError(
                "connect",
                "Could not connect to Photoshop through COM. Start desktop "
                "Photoshop, verify its COM registration, and retry. Original "
                "error: " + str(exc),
                exception_type=type(exc).__name__,
            ) from exc

    def execute_jsx(self, jsx: str) -> str:
        """Execute ExtendScript through Application.doJavaScript()."""
        if not isinstance(jsx, str) or not jsx.strip():
            raise PhotoshopBridgeError(
                "execute_jsx",
                "JSX source must be a non-empty string.",
                exception_type="InvalidJSXInput",
            )

        execution_jsx = jsx
        if self.execution_mode == "interactive":
            execution_jsx = jsx.replace("DialogModes.NO", "DialogModes.ALL")

        try:
            result = self.app.doJavaScript(execution_jsx)
        except Exception as exc:
            raise PhotoshopBridgeError(
                "execute_jsx",
                "Photoshop rejected or failed to execute JSX. Original error: " + str(exc),
                exception_type=type(exc).__name__,
                jsx_preview=_clean_jsx_preview(execution_jsx),
            ) from exc

        if result is None:
            return ""
        if isinstance(result, bytes):
            return result.decode("utf-8", errors="replace")
        return str(result)

    def get_status(self) -> Dict[str, Any]:
        """Read Photoshop version and current document state without modifying it."""
        jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    try {
        var count = app.documents.length;
        var activeName = count > 0 ? app.activeDocument.name : "";
        return [
            "SCI_PS_STATUS",
            "OK",
            clean(app.version),
            String(count),
            clean(activeName)
        ].join("\t");
    } catch (e) {
        return [
            "SCI_PS_STATUS",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean(e.line)
        ].join("\t");
    }
})();
"""
        raw = self.execute_jsx(jsx)
        fields = _parse_protocol(raw, "SCI_PS_STATUS", 5, "status", jsx)

        return {
            "connected": True,
            "photoshop_version": fields[2],
            "open_document_count": int(fields[3]),
            "active_document": fields[4] or None,
        }

    def lock_units(self) -> Dict[str, Any]:
        """Set Photoshop ruler units to Pixels and type units to Points."""
        jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    try {
        var previousRuler = clean(app.preferences.rulerUnits);
        var previousType = clean(app.preferences.typeUnits);

        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var currentRuler = clean(app.preferences.rulerUnits);
        var currentType = clean(app.preferences.typeUnits);
        var rulerOK = app.preferences.rulerUnits == Units.PIXELS;
        var typeOK = app.preferences.typeUnits == TypeUnits.POINTS;

        return [
            "SCI_PS_UNITS",
            "OK",
            previousRuler,
            previousType,
            currentRuler,
            currentType,
            String(rulerOK),
            String(typeOK)
        ].join("\t");
    } catch (e) {
        return [
            "SCI_PS_UNITS",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean(e.line)
        ].join("\t");
    }
})();
"""
        raw = self.execute_jsx(jsx)
        fields = _parse_protocol(raw, "SCI_PS_UNITS", 8, "lock_units", jsx)

        ruler_ok = fields[6].lower() == "true"
        type_ok = fields[7].lower() == "true"

        result = {
            "previous_ruler_units": fields[2],
            "previous_type_units": fields[3],
            "ruler_units": fields[4],
            "type_units": fields[5],
            "ruler_units_pixels": ruler_ok,
            "type_units_points": type_ok,
            "passed": ruler_ok and type_ok,
        }

        if not result["passed"]:
            raise PhotoshopBridgeError(
                "lock_units",
                "Photoshop did not retain the requested Pixels/Points units: "
                + repr(result),
                exception_type="UnitVerificationError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        return result

    def test_jsx_exception_capture(self) -> Dict[str, Any]:
        """Verify that a controlled ExtendScript exception can be caught and returned."""
        jsx = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    var caught = false;
    var message = "";
    try {
        throw new Error("SCI_EXPECTED_TEST_EXCEPTION");
    } catch (e) {
        caught = String(e.message).indexOf("SCI_EXPECTED_TEST_EXCEPTION") >= 0;
        message = clean(e.message);
    }
    return [
        "SCI_PS_EXCEPTION",
        "OK",
        String(caught),
        message
    ].join("\t");
})();
"""
        raw = self.execute_jsx(jsx)
        fields = _parse_protocol(raw, "SCI_PS_EXCEPTION", 4, "exception_test", jsx)
        passed = fields[2].lower() == "true"

        if not passed:
            raise PhotoshopBridgeError(
                "exception_test",
                "The controlled ExtendScript exception was not captured as expected.",
                exception_type="ExceptionCaptureError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        return {
            "passed": passed,
            "exception_message": fields[3],
        }

    def diagnose(self) -> Dict[str, Any]:
        """Run non-destructive Photoshop/JSX diagnostics and return a report."""
        report: Dict[str, Any] = {
            "python_executable": sys.executable,
            "python_version": platform.python_version(),
            "photoshop_python_api_version": _package_version("photoshop-python-api"),
            "comtypes_version": _package_version("comtypes"),
            "photoshop": None,
            "unit_lock": None,
            "jsx_exception_capture": None,
            "passed": False,
        }

        report["photoshop"] = self.get_status()
        report["unit_lock"] = self.lock_units()
        report["jsx_exception_capture"] = self.test_jsx_exception_capture()
        report["passed"] = (
            report["photoshop"]["connected"]
            and report["unit_lock"]["passed"]
            and report["jsx_exception_capture"]["passed"]
        )
        return report



# ---------------------------------------------------------------------------
# Stage 2: scientific grid composition
# ---------------------------------------------------------------------------

JOURNAL_WIDTH_PRESETS = {
    "single": 8.5,
    "half_double": 12.0,
    "double": 17.6,
}


def cm_to_pixels(value_cm: float, dpi: int) -> int:
    """Convert centimetres to integer pixels using the requested document DPI."""
    return int(round(float(value_cm) * int(dpi) / 2.54))


def _image_dimensions(image_path: Path) -> tuple[int, int]:
    """Read raster dimensions without modifying or re-encoding the source file."""
    try:
        from PIL import Image
    except ImportError as exc:
        raise PhotoshopBridgeError(
            "compose_grid",
            "Pillow is required to inspect image dimensions. Install the Skill requirements.",
            exception_type="MissingDependency",
        ) from exc

    try:
        with Image.open(str(image_path)) as image:
            width, height = image.size
    except Exception as exc:
        raise PhotoshopBridgeError(
            "compose_grid",
            "Could not read raster dimensions for " + str(image_path) + ": " + str(exc),
            exception_type=type(exc).__name__,
        ) from exc

    if width <= 0 or height <= 0:
        raise PhotoshopBridgeError(
            "compose_grid",
            "Image has invalid dimensions: " + str(image_path),
            exception_type="InvalidImageDimensions",
        )
    return int(width), int(height)


def _detect_trim_steps(image_path: Path) -> list[str]:
    """Detect only the border trims actually needed, without changing the source.

    Steps are returned in execution order: transparent pixels first, then a
    same-color border matching the resulting top-left pixel.
    """
    try:
        from PIL import Image, ImageChops
    except ImportError as exc:
        raise PhotoshopBridgeError(
            "compose_grid",
            "Pillow is required for Auto-Trim inspection.",
            exception_type="MissingDependency",
        ) from exc

    try:
        with Image.open(str(image_path)) as source:
            rgba = source.convert("RGBA")
            full_bounds = (0, 0, rgba.width, rgba.height)
            alpha_bbox = rgba.getchannel("A").getbbox()
            steps: list[str] = []

            if alpha_bbox is None:
                return []

            if alpha_bbox != full_bounds:
                steps.append("transparent")
                inspection_image = rgba.crop(alpha_bbox)
            else:
                inspection_image = rgba

            rgb = inspection_image.convert("RGB")
            corner_color = rgb.getpixel((0, 0))
            background = Image.new("RGB", rgb.size, corner_color)
            different_pixels = ImageChops.difference(rgb, background)
            different_bbox = different_pixels.getbbox()

            # Photoshop TOPLEFT trim is needed only when pixels that differ
            # from the corner-color are inset from at least one canvas edge.
            if different_bbox is not None and different_bbox != (0, 0, rgb.width, rgb.height):
                steps.append("top_left")

            return steps
    except Exception as exc:
        raise PhotoshopBridgeError(
            "compose_grid",
            "Could not inspect trim borders for " + str(image_path) + ": " + str(exc),
            exception_type=type(exc).__name__,
        ) from exc


def _panel_name(index: int) -> str:
    """Return spreadsheet-style lowercase panel names: a..z, aa, ab, ..."""
    if index < 0:
        raise ValueError("Panel index must be non-negative.")
    value = index + 1
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("a") + remainder) + result
    return result


def calculate_grid_layout(
    image_sizes: list[tuple[int, int]],
    *,
    width_cm: float,
    height_cm: Optional[float] = None,
    rows: int,
    cols: int,
    dpi: int = 300,
    margin_cm: float = 0.3,
    gap_cm: float = 0.2,
) -> Dict[str, Any]:
    """Calculate a row-major grid in pixels; no Photoshop connection is needed."""
    if not image_sizes:
        raise ValueError("At least one image is required.")
    if rows < 1 or cols < 1:
        raise ValueError("rows and cols must both be positive integers.")
    if len(image_sizes) != rows * cols:
        raise ValueError(
            "The number of images must equal rows × cols "
            f"({len(image_sizes)} supplied, {rows * cols} required)."
        )
    if dpi not in (300, 600):
        raise ValueError("dpi must be 300 or 600.")
    if width_cm <= 0:
        raise ValueError("width_cm must be greater than zero.")
    if height_cm is not None and height_cm <= 0:
        raise ValueError("height_cm must be greater than zero.")
    if margin_cm < 0 or gap_cm < 0:
        raise ValueError("margin_cm and gap_cm cannot be negative.")
    if any(w <= 0 or h <= 0 for w, h in image_sizes):
        raise ValueError("All source images must have positive dimensions.")

    width_px = cm_to_pixels(width_cm, dpi)
    margin_px = cm_to_pixels(margin_cm, dpi)
    gap_px = cm_to_pixels(gap_cm, dpi)

    usable_width = width_px - 2 * margin_px - gap_px * (cols - 1)
    if usable_width <= 0:
        raise ValueError("Margins and gaps leave no usable canvas width.")

    cell_width = float(usable_width) / cols

    if height_cm is None:
        # Derive cell height from the median source aspect ratio. This gives
        # the grid a sensible initial print aspect ratio without cropping data.
        aspect_ratios = sorted(float(w) / float(h) for w, h in image_sizes)
        middle = len(aspect_ratios) // 2
        if len(aspect_ratios) % 2:
            median_aspect = aspect_ratios[middle]
        else:
            median_aspect = (aspect_ratios[middle - 1] + aspect_ratios[middle]) / 2.0
        estimated_cell_height = cell_width / median_aspect
        height_px = int(round(
            estimated_cell_height * rows
            + 2 * margin_px
            + gap_px * (rows - 1)
        ))
    else:
        height_px = cm_to_pixels(height_cm, dpi)

    usable_height = height_px - 2 * margin_px - gap_px * (rows - 1)
    if usable_height <= 0:
        raise ValueError("Margins and gaps leave no usable canvas height.")

    cell_height = float(usable_height) / rows
    panels = []

    for index, (source_width, source_height) in enumerate(image_sizes):
        row = index // cols
        col = index % cols
        x = margin_px + col * (cell_width + gap_px)
        y = margin_px + row * (cell_height + gap_px)
        panels.append({
            "index": index,
            "panel": _panel_name(index),
            "group_name": "Group_Panel_" + _panel_name(index),
            "layer_name": "Image_Panel_" + _panel_name(index),
            "source_width_px": source_width,
            "source_height_px": source_height,
            "x": int(round(x)),
            "y": int(round(y)),
            "width": int(round(cell_width)),
            "height": int(round(cell_height)),
        })

    return {
        "width_px": width_px,
        "height_px": height_px,
        "actual_width_cm": width_px * 2.54 / dpi,
        "actual_height_cm": height_px * 2.54 / dpi,
        "dpi": dpi,
        "rows": rows,
        "cols": cols,
        "margin_px": margin_px,
        "gap_px": gap_px,
        "cell_width_px": cell_width,
        "cell_height_px": cell_height,
        "panels": panels,
    }


def _build_compose_grid_jsx(config: Dict[str, Any]) -> str:
    """Build a self-contained ExtendScript payload from validated JSON data."""
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function px(value) {
        return value.as("px");
    }
    function boundsOf(layer) {
        var b = layer.bounds;
        return {
            left: px(b[0]),
            top: px(b[1]),
            right: px(b[2]),
            bottom: px(b[3])
        };
    }
    function placeEmbeddedSmartObject(filePath) {
        var fileRef = new File(filePath);
        if (!fileRef.exists) {
            throw new Error("Source image does not exist: " + filePath);
        }
        var descriptor = new ActionDescriptor();
        descriptor.putPath(charIDToTypeID("null"), fileRef);
        descriptor.putEnumerated(
            charIDToTypeID("FTcs"),
            charIDToTypeID("QCSt"),
            charIDToTypeID("Qcsa")
        );
        var offset = new ActionDescriptor();
        offset.putUnitDouble(charIDToTypeID("Hrzn"), charIDToTypeID("#Pxl"), 0);
        offset.putUnitDouble(charIDToTypeID("Vrtc"), charIDToTypeID("#Pxl"), 0);
        descriptor.putObject(charIDToTypeID("Ofst"), charIDToTypeID("Ofst"), offset);
        executeAction(charIDToTypeID("Plc "), descriptor, DialogModes.NO);
        return app.activeDocument.activeLayer;
    }
    function trimEmbeddedObject(parentDoc, smartLayer, trimSteps) {
        var smartDoc = null;
        var changed = false;
        try {
            if (!trimSteps || trimSteps.length === 0) {
                return {
                    changed: false,
                    original_width: 0,
                    original_height: 0,
                    trimmed_width: 0,
                    trimmed_height: 0
                };
            }

            parentDoc.activeLayer = smartLayer;
            executeAction(
                stringIDToTypeID("placedLayerEditContents"),
                new ActionDescriptor(),
                DialogModes.NO
            );
            smartDoc = app.activeDocument;

            var oldWidth = px(smartDoc.width);
            var oldHeight = px(smartDoc.height);

            for (var stepIndex = 0; stepIndex < trimSteps.length; stepIndex++) {
                var trimStep = trimSteps[stepIndex];

                if (trimStep === "transparent") {
                    smartDoc.trim(TrimType.TRANSPARENT, true, true, true, true);
                } else if (trimStep === "top_left") {
                    smartDoc.trim(TrimType.TOPLEFT, true, true, true, true);
                } else {
                    throw new Error("Unsupported trim step: " + trimStep);
                }
            }

            var newWidth = px(smartDoc.width);
            var newHeight = px(smartDoc.height);
            changed = (newWidth < oldWidth || newHeight < oldHeight);

            smartDoc.close(SaveOptions.SAVECHANGES);
            smartDoc = null;
            app.activeDocument = parentDoc;

            return {
                changed: changed,
                original_width: oldWidth,
                original_height: oldHeight,
                trimmed_width: newWidth,
                trimmed_height: newHeight
            };
        } catch (trimError) {
            if (smartDoc !== null) {
                try {
                    smartDoc.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeError) {}
            }
            app.activeDocument = parentDoc;
            throw trimError;
        }
    }

    var cfg = __SCI_FIGURE_CONFIG__;
    var doc = null;
    var previousDialogs = app.displayDialogs;
    var trimmedCount = 0;

    try {
        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        if (cfg.output_path) {
            var existingOutput = new File(cfg.output_path);
            if (existingOutput.exists && !cfg.overwrite) {
                throw new Error("Output already exists; set overwrite=true to replace it: " + cfg.output_path);
            }
            if (existingOutput.exists && cfg.overwrite) {
                if (!existingOutput.remove()) {
                    throw new Error("Could not remove existing output file: " + cfg.output_path);
                }
            }
        }

        doc = app.documents.add(
            cfg.width_px,
            cfg.height_px,
            cfg.dpi,
            cfg.document_name,
            NewDocumentMode.RGB,
            DocumentFill.WHITE
        );

        for (var i = 0; i < cfg.panels.length; i++) {
            var panel = cfg.panels[i];
            app.activeDocument = doc;

            var panelGroup = doc.layerSets.add();
            panelGroup.name = panel.group_name;

            var imageLayer = placeEmbeddedSmartObject(panel.path);
            imageLayer.name = panel.layer_name;
            imageLayer.move(panelGroup, ElementPlacement.INSIDE);

            if (cfg.auto_trim && panel.trim_steps.length > 0) {
                var trimResult = trimEmbeddedObject(doc, imageLayer, panel.trim_steps);
                if (trimResult.changed) {
                    trimmedCount++;
                }
            }

            app.activeDocument = doc;
            doc.activeLayer = imageLayer;

            var imageBounds = boundsOf(imageLayer);
            var imageWidth = imageBounds.right - imageBounds.left;
            var imageHeight = imageBounds.bottom - imageBounds.top;

            if (imageWidth <= 0 || imageHeight <= 0) {
                throw new Error("Placed image has empty bounds: " + panel.path);
            }

            var scale = Math.min(
                panel.width / imageWidth,
                panel.height / imageHeight
            );
            if (!(scale > 0)) {
                throw new Error("Invalid panel scale for: " + panel.path);
            }

            imageLayer.resize(
                scale * 100,
                scale * 100,
                AnchorPosition.MIDDLECENTER
            );

            imageBounds = boundsOf(imageLayer);
            imageWidth = imageBounds.right - imageBounds.left;
            imageHeight = imageBounds.bottom - imageBounds.top;

            var targetLeft;
            var targetTop;
            if (cfg.align === "top-left") {
                targetLeft = panel.x;
                targetTop = panel.y;
            } else {
                targetLeft = panel.x + (panel.width - imageWidth) / 2;
                targetTop = panel.y + (panel.height - imageHeight) / 2;
            }

            imageLayer.translate(
                targetLeft - imageBounds.left,
                targetTop - imageBounds.top
            );
        }

        app.activeDocument = doc;

        var savedPath = "";
        if (cfg.output_path) {
            var saveFile = new File(cfg.output_path);
            var psdOptions = new PhotoshopSaveOptions();
            psdOptions.embedColorProfile = true;
            psdOptions.maximizeCompatibility = true;
            doc.saveAs(saveFile, psdOptions, true, Extension.LOWERCASE);
            savedPath = saveFile.fsName;
        }

        return [
            "SCI_PS_COMPOSE",
            "OK",
            clean(doc.name),
            String(px(doc.width)),
            String(px(doc.height)),
            String(doc.resolution),
            String(cfg.panels.length),
            String(trimmedCount),
            clean(savedPath)
        ].join("\t");

    } catch (e) {
        var errorLine = (e.line === undefined) ? "" : e.line;
        var partialDocument = doc === null ? "NO_DOCUMENT" : clean(doc.name);

        // The document was created by this compose operation. If composition
        // fails, close that partial document without saving so it cannot be
        // mistaken for a valid figure or leave Photoshop in a dirty state.
        if (doc !== null) {
            try {
                app.activeDocument = doc;
                doc.close(SaveOptions.DONOTSAVECHANGES);
                partialDocument = "CLOSED_PARTIAL_DOCUMENT";
            } catch (closeError) {
                partialDocument = "PARTIAL_DOCUMENT_REMAINS:" + clean(closeError.message);
            }
        }

        return [
            "SCI_PS_COMPOSE",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean(errorLine),
            partialDocument
        ].join("\t");
    } finally {
        app.displayDialogs = previousDialogs;
    }
})();
"""
    return template.replace("__SCI_FIGURE_CONFIG__", payload)


def compose_grid(
    image_paths: list[str | Path],
    output_path: Optional[str | Path] = None,
    *,
    preset: str = "double",
    width_cm: Optional[float] = None,
    height_cm: Optional[float] = None,
    rows: int = 2,
    cols: int = 2,
    dpi: int = 300,
    margin_cm: float = 0.3,
    gap_cm: float = 0.2,
    align: str = "center",
    auto_trim: bool = False,
    overwrite: bool = False,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Compose raster panels as embedded Smart Objects in a native Photoshop PSD."""
    if align not in ("center", "top-left"):
        raise ValueError("align must be 'center' or 'top-left'.")
    if preset not in ("single", "half_double", "double", "custom"):
        raise ValueError("preset must be single, half_double, double, or custom.")

    if preset == "custom":
        if width_cm is None:
            raise ValueError("width_cm is required when preset='custom'.")
        selected_width = float(width_cm)
    else:
        selected_width = JOURNAL_WIDTH_PRESETS[preset]
        if width_cm is not None and abs(float(width_cm) - selected_width) > 1e-9:
            raise ValueError(
                "width_cm conflicts with preset '" + preset +
                "'. Use preset='custom' to specify a different width."
            )

    paths = [Path(p).expanduser().resolve(strict=True) for p in image_paths]
    if len(paths) != rows * cols:
        raise ValueError(
            "The number of images must equal rows × cols "
            f"({len(paths)} supplied, {rows * cols} required)."
        )
    if any(not path.is_file() for path in paths):
        raise ValueError("Every input path must be a file.")
    image_sizes = [_image_dimensions(path) for path in paths]

    layout = calculate_grid_layout(
        image_sizes,
        width_cm=selected_width,
        height_cm=height_cm,
        rows=rows,
        cols=cols,
        dpi=dpi,
        margin_cm=margin_cm,
        gap_cm=gap_cm,
    )

    output_file: Optional[Path] = None
    if output_path is not None:
        output_file = Path(output_path).expanduser().resolve()
        if output_file.suffix.lower() != ".psd":
            raise ValueError("compose_grid output_path must end in .psd.")
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists() and not overwrite:
            raise FileExistsError(
                "Output already exists; pass overwrite=True to replace it: "
                + str(output_file)
            )

    document_name = (
        output_file.stem if output_file is not None else "Scientific_Figure"
    )
    panels = []
    for path, panel_geometry in zip(paths, layout["panels"]):
        panel = dict(panel_geometry)
        panel["path"] = str(path)
        panel["trim_steps"] = _detect_trim_steps(path) if auto_trim else []
        panels.append(panel)

    config = {
        "document_name": document_name,
        "width_px": layout["width_px"],
        "height_px": layout["height_px"],
        "dpi": dpi,
        "align": align,
        "auto_trim": bool(auto_trim),
        "output_path": str(output_file) if output_file is not None else "",
        "overwrite": bool(overwrite),
        "panels": panels,
    }

    active_bridge = bridge if bridge is not None else PhotoshopBridge()
    jsx = _build_compose_grid_jsx(config)
    raw = active_bridge.execute_jsx(jsx)
    fields = _parse_protocol(
        raw,
        "SCI_PS_COMPOSE",
        9,
        "compose_grid",
        jsx,
    )

    result = {
        "document_name": fields[2],
        "width_px": int(fields[3]),
        "height_px": int(fields[4]),
        "dpi": float(fields[5]),
        "panel_count": int(fields[6]),
        "trimmed_panel_count": int(fields[7]),
        "saved_psd": fields[8] or None,
        "requested_width_cm": selected_width,
        "actual_width_cm": layout["actual_width_cm"],
        "actual_height_cm": layout["actual_height_cm"],
        "rows": rows,
        "cols": cols,
        "align": align,
        "auto_trim": bool(auto_trim),
        "panels": panels,
        "passed": True,
    }

    if result["width_px"] != layout["width_px"] or result["height_px"] != layout["height_px"]:
        raise PhotoshopBridgeError(
            "compose_grid",
            "Photoshop document dimensions do not match the calculated layout.",
            exception_type="CanvasDimensionMismatch",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    if output_file is not None:
        if not output_file.is_file() or output_file.stat().st_size <= 0:
            raise PhotoshopBridgeError(
                "compose_grid",
                "Photoshop reported a save but the PSD file is missing or empty: "
                + str(output_file),
                exception_type="PsdSaveVerificationError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

    return result



# ---------------------------------------------------------------------------
# Stage 3: editable panel labels
# ---------------------------------------------------------------------------

PANEL_LABEL_STYLES = ("lowercase", "uppercase", "parenthesized")


def _panel_label_text(panel_id: str, style: str) -> str:
    """Format a panel identifier without changing its underlying group name."""
    if not panel_id or not panel_id.isalpha():
        raise ValueError("Panel identifiers must contain letters only.")
    normalized = panel_id.lower()
    if style == "lowercase":
        return normalized
    if style == "uppercase":
        return normalized.upper()
    if style == "parenthesized":
        return "(" + normalized + ")"
    raise ValueError("style must be lowercase, uppercase, or parenthesized.")


def _parse_hex_color(value: str) -> tuple[int, int, int]:
    """Parse a #RRGGBB string into integer RGB channels."""
    if not isinstance(value, str):
        raise ValueError("color_hex must be a #RRGGBB string.")
    token = value.strip()
    if len(token) != 7 or not token.startswith("#"):
        raise ValueError("color_hex must use the #RRGGBB format.")
    try:
        return tuple(int(token[index:index + 2], 16) for index in (1, 3, 5))
    except ValueError as exc:
        raise ValueError("color_hex contains a non-hexadecimal value.") from exc


def _build_add_panel_labels_jsx(config: Dict[str, Any]) -> str:
    """Build JSX that adds editable text layers and saves a separate PSD."""
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function px(value) {
        return value.as("px");
    }
    function labelFor(panelId, style) {
        var lower = panelId.toLowerCase();
        if (style === "lowercase") {
            return lower;
        }
        if (style === "uppercase") {
            return lower.toUpperCase();
        }
        if (style === "parenthesized") {
            return "(" + lower + ")";
        }
        throw new Error("Unsupported panel-label style: " + style);
    }

    var cfg = __SCI_PANEL_LABEL_CONFIG__;
    var previousDialogs = app.displayDialogs;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var doc = null;
    var response = "";

    try {
        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var sourceFile = new File(cfg.input_path);
        if (!sourceFile.exists) {
            throw new Error("Input PSD does not exist: " + cfg.input_path);
        }

        var outputFile = new File(cfg.output_path);
        if (outputFile.exists) {
            if (!cfg.overwrite) {
                throw new Error("Output already exists; set overwrite=true to replace it: " + cfg.output_path);
            }
            if (!outputFile.remove()) {
                throw new Error("Could not remove existing output file: " + cfg.output_path);
            }
        }

        doc = app.open(sourceFile);
        var insetPx = cfg.inset_cm * doc.resolution / 2.54;
        var fontSizePx = cfg.size_pt * doc.resolution / 72.0;
        var textColor = new SolidColor();
        textColor.rgb.red = cfg.color_rgb[0];
        textColor.rgb.green = cfg.color_rgb[1];
        textColor.rgb.blue = cfg.color_rgb[2];

        var groupCount = 0;
        var labelsCreated = 0;
        var labelNames = [];

        for (var i = 0; i < doc.layerSets.length; i++) {
            var panelGroup = doc.layerSets[i];
            var prefix = "Group_Panel_";
            if (panelGroup.name.indexOf(prefix) !== 0) {
                continue;
            }

            var panelId = panelGroup.name.substring(prefix.length);
            if (!/^[a-z]+$/i.test(panelId)) {
                throw new Error("Invalid panel group name: " + panelGroup.name);
            }

            groupCount++;
            var labelLayerName = "Panel_Label_" + panelId.toLowerCase();

            for (var existingIndex = 0; existingIndex < panelGroup.artLayers.length; existingIndex++) {
                if (panelGroup.artLayers[existingIndex].name === labelLayerName) {
                    throw new Error(
                        "A generated panel label already exists in " + panelGroup.name +
                        "; use the original unlabeled PSD as input."
                    );
                }
            }

            var groupBounds = panelGroup.bounds;
            var left = px(groupBounds[0]);
            var top = px(groupBounds[1]);
            var right = px(groupBounds[2]);
            var bottom = px(groupBounds[3]);
            if (!(right > left && bottom > top)) {
                throw new Error("Panel group has empty bounds: " + panelGroup.name);
            }

            var textLayer = panelGroup.artLayers.add();
            textLayer.kind = LayerKind.TEXT;
            textLayer.name = labelLayerName;

            var textItem = textLayer.textItem;
            textItem.contents = labelFor(panelId, cfg.style);
            textItem.font = cfg.font_postscript;
            textItem.size = new UnitValue(cfg.size_pt, "pt");
            textItem.color = textColor;
            textItem.justification = Justification.LEFT;
            textItem.position = [
                new UnitValue(left + insetPx, "px"),
                new UnitValue(top + insetPx + fontSizePx, "px")
            ];

            if (textLayer.kind !== LayerKind.TEXT) {
                throw new Error("Photoshop did not create an editable text layer for " + panelGroup.name);
            }

            labelsCreated++;
            labelNames.push(panelGroup.name + "=" + textItem.contents);
        }

        if (groupCount < 1) {
            throw new Error("No Group_Panel_* groups were found in the input PSD.");
        }
        if (labelsCreated !== groupCount) {
            throw new Error("Not every panel group received exactly one text label.");
        }

        var psdOptions = new PhotoshopSaveOptions();
        psdOptions.embedColorProfile = true;
        psdOptions.maximizeCompatibility = true;
        doc.saveAs(outputFile, psdOptions, true, Extension.LOWERCASE);

        if (!outputFile.exists || outputFile.length <= 0) {
            throw new Error("Photoshop did not create a non-empty labeled PSD.");
        }

        response = [
            "SCI_PS_LABELS",
            "OK",
            clean(doc.name),
            String(groupCount),
            String(labelsCreated),
            cfg.style,
            labelNames.join(","),
            outputFile.fsName
        ].join("\t");

    } catch (e) {
        response = [
            "SCI_PS_LABELS",
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
"""
    return template.replace("__SCI_PANEL_LABEL_CONFIG__", payload)


def add_panel_labels(
    input_path: str | Path,
    output_path: str | Path,
    *,
    style: str = "lowercase",
    size_pt: float = 10.0,
    font_postscript: str = "Arial-BoldMT",
    inset_cm: float = 0.1,
    color_hex: str = "#000000",
    overwrite: bool = False,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Add editable Arial Bold panel labels and save to a separate PSD.

    The source PSD is opened read-only from the caller's perspective and is
    never saved in place. The output is a new PSD containing native Photoshop
    text layers named Panel_Label_<id> inside the corresponding panel groups.
    """
    if style not in PANEL_LABEL_STYLES:
        raise ValueError(
            "style must be one of: " + ", ".join(PANEL_LABEL_STYLES)
        )
    if not 10.0 <= float(size_pt) <= 12.0:
        raise ValueError("size_pt must be between 10 and 12 points.")
    if not isinstance(font_postscript, str) or not font_postscript.strip():
        raise ValueError("font_postscript must be a non-empty Photoshop PostScript font name.")
    if any(char in font_postscript for char in "\r\n\t"):
        raise ValueError("font_postscript cannot contain tabs or newlines.")
    if float(inset_cm) < 0:
        raise ValueError("inset_cm cannot be negative.")

    source_file = Path(input_path).expanduser().resolve(strict=True)
    if not source_file.is_file() or source_file.suffix.lower() != ".psd":
        raise ValueError("input_path must be an existing PSD file.")

    output_file = Path(output_path).expanduser().resolve()
    if output_file.suffix.lower() != ".psd":
        raise ValueError("output_path must end in .psd.")
    if str(source_file).casefold() == str(output_file).casefold():
        raise ValueError("output_path must differ from input_path; the source PSD is never modified.")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    if output_file.exists() and not overwrite:
        raise FileExistsError(
            "Output already exists; pass overwrite=True to replace it: "
            + str(output_file)
        )

    color_rgb = _parse_hex_color(color_hex)
    config = {
        "input_path": str(source_file),
        "output_path": str(output_file),
        "style": style,
        "size_pt": float(size_pt),
        "font_postscript": font_postscript.strip(),
        "inset_cm": float(inset_cm),
        "color_rgb": list(color_rgb),
        "overwrite": bool(overwrite),
    }

    active_bridge = bridge if bridge is not None else PhotoshopBridge()
    jsx = _build_add_panel_labels_jsx(config)
    raw = active_bridge.execute_jsx(jsx)
    fields = _parse_protocol(
        raw,
        "SCI_PS_LABELS",
        8,
        "add_panel_labels",
        jsx,
    )

    group_count = int(fields[3])
    labels_created = int(fields[4])
    if group_count < 1 or labels_created != group_count:
        raise PhotoshopBridgeError(
            "add_panel_labels",
            "Photoshop did not create exactly one label for every panel group: "
            + repr(fields[3:6]),
            exception_type="PanelLabelCountMismatch",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    if not output_file.is_file() or output_file.stat().st_size <= 0:
        raise PhotoshopBridgeError(
            "add_panel_labels",
            "Photoshop reported a save but the labeled PSD is missing or empty: "
            + str(output_file),
            exception_type="PsdSaveVerificationError",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    return {
        "document_name": fields[2],
        "panel_group_count": group_count,
        "labels_created": labels_created,
        "style": fields[5],
        "labels": fields[6].split(",") if fields[6] else [],
        "input_psd": str(source_file),
        "saved_psd": str(output_file),
        "font_postscript": config["font_postscript"],
        "font_size_pt": config["size_pt"],
        "color_hex": color_hex.upper(),
        "passed": True,
    }




# ---------------------------------------------------------------------------
# Stage 4: editable scale bars
# ---------------------------------------------------------------------------

SCALE_BAR_MODES = ("pixels", "physical")


def _format_scale_number(value: float) -> str:
    """Format scale values compactly without unnecessary trailing zeros."""
    return format(float(value), ".8g")


def _physical_scale_width_px(
    physical_length: float,
    pixel_size: float,
    displayed_image_width_px: float,
    smart_object_width_px: float,
) -> float:
    """Convert a physical length to the final document-pixel width."""
    values = (
        physical_length,
        pixel_size,
        displayed_image_width_px,
        smart_object_width_px,
    )
    if any(float(value) <= 0 for value in values):
        raise ValueError("Physical calibration and image dimensions must be positive.")
    source_pixels = float(physical_length) / float(pixel_size)
    displayed_per_source_pixel = (
        float(displayed_image_width_px) / float(smart_object_width_px)
    )
    return source_pixels * displayed_per_source_pixel


def _build_add_scale_bar_jsx(config: Dict[str, Any]) -> str:
    """Build JSX that adds separate editable bar/outline/text layers."""
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function px(value) {
        return value.as("px");
    }
    function makeColor(rgb) {
        var color = new SolidColor();
        color.rgb.red = rgb[0];
        color.rgb.green = rgb[1];
        color.rgb.blue = rgb[2];
        return color;
    }
    function fillRectangle(parentDoc, targetGroup, name, left, top, right, bottom, color) {
        if (!(right > left && bottom > top)) {
            throw new Error("Invalid scale-bar rectangle bounds for " + name);
        }
        var layer = targetGroup.artLayers.add();
        layer.name = name;
        layer.kind = LayerKind.NORMAL;
        parentDoc.activeLayer = layer;
        parentDoc.selection.select([
            [left, top],
            [right, top],
            [right, bottom],
            [left, bottom]
        ]);
        parentDoc.selection.fill(color, ColorBlendMode.NORMAL, 100, false);
        parentDoc.selection.deselect();
        return layer;
    }

    var cfg = __SCI_SCALE_BAR_CONFIG__;
    var previousDialogs = app.displayDialogs;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var doc = null;
    var smartDoc = null;
    var sourceWidthPx = 0;
    var sourceHeightPx = 0;
    var measuredScale = 0;
    var barWidthPx = 0;
    var response = "";

    try {
        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var sourceFile = new File(cfg.input_path);
        if (!sourceFile.exists) {
            throw new Error("Input PSD does not exist: " + cfg.input_path);
        }

        var outputFile = new File(cfg.output_path);
        if (outputFile.exists) {
            if (!cfg.overwrite) {
                throw new Error("Output already exists; set overwrite=true to replace it: " + cfg.output_path);
            }
            if (!outputFile.remove()) {
                throw new Error("Could not remove existing output file: " + cfg.output_path);
            }
        }

        doc = app.open(sourceFile);
        var targetGroup = null;
        var expectedGroupName = "Group_Panel_" + cfg.panel_id.toLowerCase();

        for (var groupIndex = 0; groupIndex < doc.layerSets.length; groupIndex++) {
            if (doc.layerSets[groupIndex].name.toLowerCase() === expectedGroupName.toLowerCase()) {
                targetGroup = doc.layerSets[groupIndex];
                break;
            }
        }
        if (targetGroup === null) {
            throw new Error("Panel group not found: " + expectedGroupName);
        }

        var smartLayer = null;
        for (var layerIndex = 0; layerIndex < targetGroup.artLayers.length; layerIndex++) {
            var candidate = targetGroup.artLayers[layerIndex];
            if (candidate.name.indexOf("Image_Panel_") === 0 &&
                candidate.kind == LayerKind.SMARTOBJECT) {
                smartLayer = candidate;
                break;
            }
        }
        if (smartLayer === null) {
            throw new Error("Embedded Smart Object not found in " + expectedGroupName);
        }

        var imageBounds = smartLayer.bounds;
        var imageLeft = px(imageBounds[0]);
        var imageTop = px(imageBounds[1]);
        var imageRight = px(imageBounds[2]);
        var imageBottom = px(imageBounds[3]);
        var displayedImageWidth = imageRight - imageLeft;
        var displayedImageHeight = imageBottom - imageTop;
        if (!(displayedImageWidth > 0 && displayedImageHeight > 0)) {
            throw new Error("Panel Smart Object has empty on-canvas bounds.");
        }

        if (cfg.mode === "physical") {
            doc.activeLayer = smartLayer;
            executeAction(
                stringIDToTypeID("placedLayerEditContents"),
                new ActionDescriptor(),
                DialogModes.NO
            );
            smartDoc = app.activeDocument;
            sourceWidthPx = px(smartDoc.width);
            sourceHeightPx = px(smartDoc.height);
            if (!(sourceWidthPx > 0 && sourceHeightPx > 0)) {
                throw new Error("Smart Object content has invalid pixel dimensions.");
            }
            smartDoc.close(SaveOptions.DONOTSAVECHANGES);
            smartDoc = null;
            app.activeDocument = doc;

            // Use the actual post-trim Smart Object dimensions and on-canvas
            // transform; do not assume the source raster fills the PSD cell.
            measuredScale = displayedImageWidth / sourceWidthPx;
            barWidthPx = (cfg.physical_length / cfg.pixel_size) * measuredScale;
        } else if (cfg.mode === "pixels") {
            barWidthPx = cfg.bar_width_px;
            sourceWidthPx = 0;
            sourceHeightPx = 0;
            measuredScale = 1;
        } else {
            throw new Error("Unsupported scale-bar mode: " + cfg.mode);
        }

        if (!(barWidthPx >= 1)) {
            throw new Error("Calculated scale-bar width is less than one canvas pixel.");
        }

        var docResolution = doc.resolution;
        var insetPx = cfg.inset_cm * docResolution / 2.54;
        var barHeightPx = cfg.bar_height_pt * docResolution / 72.0;
        var outlinePx = cfg.outline_rgb !== null ? cfg.outline_width_pt * docResolution / 72.0 : 0;
        var labelGapPx = cfg.label_gap_pt * docResolution / 72.0;

        var innerRight = imageRight - insetPx - outlinePx;
        var innerBottom = imageBottom - insetPx - outlinePx;
        var barLeft = innerRight - barWidthPx;
        var barTop = innerBottom - barHeightPx;
        var panelInnerLeft = imageLeft + insetPx + outlinePx;
        var panelInnerTop = imageTop + insetPx + outlinePx;

        if (barLeft < panelInnerLeft || barTop < panelInnerTop) {
            throw new Error(
                "Scale bar does not fit inside " + expectedGroupName +
                "; reduce the width or inset."
            );
        }

        var barColor = makeColor(cfg.color_rgb);
        var outerBounds = {
            left: barLeft - outlinePx,
            top: barTop - outlinePx,
            right: innerRight + outlinePx,
            bottom: innerBottom + outlinePx
        };

        if (cfg.outline_rgb !== null) {
            fillRectangle(
                doc,
                targetGroup,
                "Scale_Bar_Outline_" + cfg.panel_id.toLowerCase(),
                outerBounds.left,
                outerBounds.top,
                outerBounds.right,
                outerBounds.bottom,
                makeColor(cfg.outline_rgb)
            );
        }

        fillRectangle(
            doc,
            targetGroup,
            "Scale_Bar_" + cfg.panel_id.toLowerCase(),
            barLeft,
            barTop,
            innerRight,
            innerBottom,
            barColor
        );

        var textLayer = targetGroup.artLayers.add();
        textLayer.kind = LayerKind.TEXT;
        textLayer.name = "Scale_Label_" + cfg.panel_id.toLowerCase();
        var textItem = textLayer.textItem;
        textItem.contents = cfg.label;
        textItem.font = cfg.font_postscript;
        textItem.size = new UnitValue(cfg.label_size_pt, "pt");
        textItem.color = barColor;
        textItem.justification = Justification.RIGHT;
        textItem.position = [
            new UnitValue(innerRight, "px"),
            new UnitValue(barTop - labelGapPx, "px")
        ];

        if (textLayer.kind !== LayerKind.TEXT) {
            throw new Error("Photoshop did not create an editable scale-label text layer.");
        }

        // Photoshop's textItem.position is a typographic anchor/baseline, not
        // the visible glyph bounds. Read the actual text-layer bounds first,
        // then translate the layer so its visible lower-right corner aligns
        // with the requested gap above the scale bar.
        var textBounds = textLayer.bounds;
        var labelLeft = px(textBounds[0]);
        var labelTop = px(textBounds[1]);
        var labelRight = px(textBounds[2]);
        var labelBottom = px(textBounds[3]);
        var labelTargetRight = innerRight;
        var labelTargetBottom = barTop - labelGapPx;

        textLayer.translate(
            labelTargetRight - labelRight,
            labelTargetBottom - labelBottom
        );

        textBounds = textLayer.bounds;
        labelLeft = px(textBounds[0]);
        labelTop = px(textBounds[1]);
        labelRight = px(textBounds[2]);
        labelBottom = px(textBounds[3]);

        if (labelLeft < panelInnerLeft - 0.5 ||
            labelRight > imageRight - insetPx + 0.5 ||
            labelTop < panelInnerTop - 0.5 ||
            labelBottom > barTop - labelGapPx + 0.5) {
            throw new Error(
                "Scale label cannot be fitted above the bar inside the selected panel; " +
                "use a shorter label, smaller font, or wider panel."
            );
        }

        var psdOptions = new PhotoshopSaveOptions();
        psdOptions.embedColorProfile = true;
        psdOptions.maximizeCompatibility = true;
        doc.saveAs(outputFile, psdOptions, true, Extension.LOWERCASE);

        if (!outputFile.exists || outputFile.length <= 0) {
            throw new Error("Photoshop did not create a non-empty scale-bar PSD.");
        }

        response = [
            "SCI_PS_SCALE_BAR",
            "OK",
            clean(doc.name),
            cfg.panel_id.toLowerCase(),
            cfg.mode,
            String(barWidthPx),
            clean(cfg.label),
            String(sourceWidthPx),
            String(sourceHeightPx),
            String(measuredScale),
            String(cfg.outline_rgb !== null),
            outputFile.fsName
        ].join("\t");

    } catch (e) {
        response = [
            "SCI_PS_SCALE_BAR",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean((e.line === undefined) ? "" : e.line)
        ].join("\t");
    } finally {
        if (smartDoc !== null) {
            try {
                smartDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (smartCloseError) {}
        }
        if (doc !== null) {
            try {
                doc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (docCloseError) {}
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
"""
    return template.replace("__SCI_SCALE_BAR_CONFIG__", payload)


def add_scale_bar(
    input_path: str | Path,
    output_path: str | Path,
    *,
    panel_id: str = "a",
    mode: str = "pixels",
    bar_width_px: Optional[float] = None,
    physical_length: Optional[float] = None,
    pixel_size: Optional[float] = None,
    unit: Optional[str] = None,
    label: Optional[str] = None,
    color_hex: str = "#000000",
    outline_hex: Optional[str] = None,
    bar_height_pt: float = 2.0,
    outline_width_pt: float = 0.5,
    label_size_pt: float = 8.0,
    label_gap_pt: float = 1.5,
    font_postscript: str = "ArialMT",
    inset_cm: float = 0.2,
    overwrite: bool = False,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Add an editable scale bar to one panel and save a separate PSD.

    mode='pixels' uses an explicit width measured in final document pixels.
    mode='physical' converts a requested physical length using pixel_size
    (physical units per pixel in the source Smart Object) and the actual
    post-trim Smart Object dimensions/on-canvas transform.
    """
    normalized_panel = str(panel_id).strip().lower()
    if not normalized_panel or not normalized_panel.isalpha():
        raise ValueError("panel_id must contain letters only, for example 'a' or 'aa'.")
    if mode not in SCALE_BAR_MODES:
        raise ValueError("mode must be 'pixels' or 'physical'.")

    if mode == "pixels":
        if bar_width_px is None or float(bar_width_px) <= 0:
            raise ValueError("bar_width_px must be greater than zero in pixel mode.")
        if physical_length is not None or pixel_size is not None or unit is not None:
            raise ValueError("Physical calibration arguments cannot be used in pixel mode.")
        selected_width = float(bar_width_px)
        default_label = _format_scale_number(selected_width) + " px"
    else:
        if physical_length is None or float(physical_length) <= 0:
            raise ValueError("physical_length must be greater than zero in physical mode.")
        if pixel_size is None or float(pixel_size) <= 0:
            raise ValueError("pixel_size must be greater than zero in physical mode.")
        if unit is None or not str(unit).strip():
            raise ValueError("unit is required in physical mode, for example 'µm'.")
        if bar_width_px is not None:
            raise ValueError("bar_width_px cannot be used in physical mode.")
        selected_width = None
        default_label = _format_scale_number(float(physical_length)) + " " + str(unit).strip()

    selected_label = default_label if label is None else str(label)
    if not selected_label.strip():
        raise ValueError("label cannot be empty.")
    if any(ord(char) < 32 for char in selected_label):
        raise ValueError("label cannot contain tabs, newlines, or control characters.")
    if len(selected_label) > 80:
        raise ValueError("label cannot exceed 80 characters.")

    if not 1.0 <= float(bar_height_pt) <= 12.0:
        raise ValueError("bar_height_pt must be between 1 and 12 points.")
    if not 5.0 <= float(label_size_pt) <= 12.0:
        raise ValueError("label_size_pt must be between 5 and 12 points.")
    if float(label_gap_pt) < 0:
        raise ValueError("label_gap_pt cannot be negative.")
    if float(inset_cm) < 0:
        raise ValueError("inset_cm cannot be negative.")
    if float(outline_width_pt) <= 0:
        raise ValueError("outline_width_pt must be greater than zero.")
    if not isinstance(font_postscript, str) or not font_postscript.strip():
        raise ValueError("font_postscript must be a non-empty Photoshop PostScript font name.")
    if any(char in font_postscript for char in "\r\n\t"):
        raise ValueError("font_postscript cannot contain tabs or newlines.")

    bar_rgb = _parse_hex_color(color_hex)
    if bar_rgb not in ((0, 0, 0), (255, 255, 255)):
        raise ValueError("color_hex must be black (#000000) or white (#FFFFFF).")
    if outline_hex is None:
        outline_rgb = None
    else:
        outline_rgb = _parse_hex_color(outline_hex)
        if outline_rgb == bar_rgb:
            raise ValueError("outline_hex must contrast with the bar color.")

    source_file = Path(input_path).expanduser().resolve(strict=True)
    if not source_file.is_file() or source_file.suffix.lower() != ".psd":
        raise ValueError("input_path must be an existing PSD file.")
    output_file = Path(output_path).expanduser().resolve()
    if output_file.suffix.lower() != ".psd":
        raise ValueError("output_path must end in .psd.")
    if str(source_file).casefold() == str(output_file).casefold():
        raise ValueError("output_path must differ from input_path; the source PSD is never modified.")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    if output_file.exists() and not overwrite:
        raise FileExistsError(
            "Output already exists; pass overwrite=True to replace it: "
            + str(output_file)
        )

    config = {
        "input_path": str(source_file),
        "output_path": str(output_file),
        "panel_id": normalized_panel,
        "mode": mode,
        "bar_width_px": selected_width,
        "physical_length": float(physical_length) if physical_length is not None else None,
        "pixel_size": float(pixel_size) if pixel_size is not None else None,
        "unit": str(unit).strip() if unit is not None else None,
        "label": selected_label,
        "color_rgb": list(bar_rgb),
        "outline_hex": outline_hex.upper() if outline_hex is not None else None,
        "outline_rgb": list(outline_rgb) if outline_rgb is not None else None,
        "bar_height_pt": float(bar_height_pt),
        "outline_width_pt": float(outline_width_pt),
        "label_size_pt": float(label_size_pt),
        "label_gap_pt": float(label_gap_pt),
        "font_postscript": font_postscript.strip(),
        "inset_cm": float(inset_cm),
        "overwrite": bool(overwrite),
    }

    active_bridge = bridge if bridge is not None else PhotoshopBridge()
    jsx = _build_add_scale_bar_jsx(config)
    raw = active_bridge.execute_jsx(jsx)
    fields = _parse_protocol(
        raw,
        "SCI_PS_SCALE_BAR",
        12,
        "add_scale_bar",
        jsx,
    )

    saved_width = float(fields[5])
    if saved_width < 1:
        raise PhotoshopBridgeError(
            "add_scale_bar",
            "Photoshop returned an invalid scale-bar width: " + fields[5],
            exception_type="InvalidScaleBarWidth",
            jsx_preview=_clean_jsx_preview(jsx),
        )
    if not output_file.is_file() or output_file.stat().st_size <= 0:
        raise PhotoshopBridgeError(
            "add_scale_bar",
            "Photoshop reported a save but the scale-bar PSD is missing or empty: "
            + str(output_file),
            exception_type="PsdSaveVerificationError",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    return {
        "document_name": fields[2],
        "panel_id": fields[3],
        "mode": fields[4],
        "bar_width_px": saved_width,
        "label": fields[6],
        "smart_object_width_px": int(float(fields[7])),
        "smart_object_height_px": int(float(fields[8])),
        "displayed_pixels_per_source_pixel": float(fields[9]),
        "outline": fields[10].lower() == "true",
        "input_psd": str(source_file),
        "saved_psd": str(output_file),
        "color_hex": color_hex.upper(),
        "outline_hex": outline_hex.upper() if outline_hex is not None else None,
        "passed": True,
    }



# ---------------------------------------------------------------------------
# Stage 5: recursive text-layer font normalization
# ---------------------------------------------------------------------------

def _build_unify_fonts_jsx(config: Dict[str, Any]) -> str:
    """Build one native Photoshop JSX pass for recursive text normalization."""
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function makeColor(rgb) {
        var color = new SolidColor();
        color.rgb.red = rgb[0];
        color.rgb.green = rgb[1];
        color.rgb.blue = rgb[2];
        return color;
    }
    function colorMatches(color, rgb) {
        return (
            Math.round(color.rgb.red) === rgb[0] &&
            Math.round(color.rgb.green) === rgb[1] &&
            Math.round(color.rgb.blue) === rgb[2]
        );
    }

    var cfg = __SCI_UNIFY_FONTS_CONFIG__;
    var previousDialogs = app.displayDialogs;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var doc = null;
    var textLayerCount = 0;
    var changedLayerCount = 0;
    var layerPaths = [];
    var response = "";

    function processContainer(container, parentPath) {
        var i;
        for (i = 0; i < container.artLayers.length; i++) {
            var layer = container.artLayers[i];
            if (layer.kind !== LayerKind.TEXT) {
                continue;
            }

            textLayerCount++;
            var textItem = layer.textItem;

            if (cfg.font_postscript !== null) {
                textItem.font = cfg.font_postscript;
                if (String(textItem.font) !== cfg.font_postscript) {
                    throw new Error(
                        "Font verification failed for " + parentPath + "/" + layer.name +
                        ": expected " + cfg.font_postscript + ", got " + String(textItem.font)
                    );
                }
            }

            if (cfg.size_pt !== null) {
                textItem.size = new UnitValue(cfg.size_pt, "pt");
                var actualSize = textItem.size.as("pt");
                if (Math.abs(actualSize - cfg.size_pt) > 0.05) {
                    throw new Error(
                        "Font-size verification failed for " + parentPath + "/" + layer.name +
                        ": expected " + cfg.size_pt + " pt, got " + actualSize + " pt"
                    );
                }
            }

            if (cfg.color_rgb !== null) {
                textItem.color = makeColor(cfg.color_rgb);
                if (!colorMatches(textItem.color, cfg.color_rgb)) {
                    throw new Error(
                        "Text-color verification failed for " + parentPath + "/" + layer.name
                    );
                }
            }

            changedLayerCount++;
            layerPaths.push(parentPath + "/" + layer.name);
        }

        for (var groupIndex = 0; groupIndex < container.layerSets.length; groupIndex++) {
            var nestedGroup = container.layerSets[groupIndex];
            processContainer(nestedGroup, parentPath + "/" + nestedGroup.name);
        }
    }

    try {
        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var sourceFile = new File(cfg.input_path);
        if (!sourceFile.exists) {
            throw new Error("Input PSD does not exist: " + cfg.input_path);
        }

        var outputFile = new File(cfg.output_path);
        if (outputFile.exists) {
            if (!cfg.overwrite) {
                throw new Error(
                    "Output already exists; set overwrite=true to replace it: " +
                    cfg.output_path
                );
            }
            if (!outputFile.remove()) {
                throw new Error("Could not remove existing output file: " + cfg.output_path);
            }
        }

        doc = app.open(sourceFile);
        processContainer(doc, doc.name);

        if (textLayerCount < 1) {
            throw new Error("No editable Photoshop text layers were found in the document.");
        }
        if (changedLayerCount !== textLayerCount) {
            throw new Error(
                "Not every text layer was normalized: " + changedLayerCount +
                " of " + textLayerCount
            );
        }

        var psdOptions = new PhotoshopSaveOptions();
        psdOptions.embedColorProfile = true;
        psdOptions.maximizeCompatibility = true;
        doc.saveAs(outputFile, psdOptions, true, Extension.LOWERCASE);

        if (!outputFile.exists || outputFile.length <= 0) {
            throw new Error("Photoshop did not create a non-empty font-normalized PSD.");
        }

        response = [
            "SCI_PS_UNIFY_FONTS",
            "OK",
            clean(doc.name),
            String(textLayerCount),
            String(changedLayerCount),
            cfg.font_postscript === null ? "" : clean(cfg.font_postscript),
            cfg.size_pt === null ? "" : String(cfg.size_pt),
            cfg.color_hex === null ? "" : cfg.color_hex,
            outputFile.fsName,
            layerPaths.join(";")
        ].join("\t");

    } catch (e) {
        response = [
            "SCI_PS_UNIFY_FONTS",
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
"""
    return template.replace("__SCI_UNIFY_FONTS_CONFIG__", payload)


def unify_fonts(
    input_path: str | Path,
    output_path: str | Path,
    *,
    font_postscript: Optional[str] = None,
    size_pt: Optional[float] = None,
    color_hex: Optional[str] = None,
    overwrite: bool = False,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Normalize selected attributes of every editable text layer recursively.

    Text layers are found in the document and every nested LayerSet. Any
    attribute left as None is preserved. The source PSD is never saved in
    place; the normalized document is saved to a separate PSD.
    """
    if font_postscript is None and size_pt is None and color_hex is None:
        raise ValueError(
            "Specify at least one of font_postscript, size_pt, or color_hex."
        )
    if font_postscript is not None:
        if not isinstance(font_postscript, str) or not font_postscript.strip():
            raise ValueError("font_postscript must be a non-empty Photoshop PostScript font name.")
        if any(char in font_postscript for char in "\r\n\t"):
            raise ValueError("font_postscript cannot contain tabs or newlines.")
        font_postscript = font_postscript.strip()

    normalized_size: Optional[float] = None
    if size_pt is not None:
        normalized_size = float(size_pt)
        if not 1.0 <= normalized_size <= 200.0:
            raise ValueError("size_pt must be between 1 and 200 points.")

    color_rgb = _parse_hex_color(color_hex) if color_hex is not None else None

    source_file = Path(input_path).expanduser().resolve(strict=True)
    if not source_file.is_file() or source_file.suffix.lower() != ".psd":
        raise ValueError("input_path must be an existing PSD file.")

    output_file = Path(output_path).expanduser().resolve()
    if output_file.suffix.lower() != ".psd":
        raise ValueError("output_path must end in .psd.")
    if str(source_file).casefold() == str(output_file).casefold():
        raise ValueError("output_path must differ from input_path; the source PSD is never modified.")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    if output_file.exists() and not overwrite:
        raise FileExistsError(
            "Output already exists; pass overwrite=True to replace it: "
            + str(output_file)
        )

    config = {
        "input_path": str(source_file),
        "output_path": str(output_file),
        "font_postscript": font_postscript,
        "size_pt": normalized_size,
        "color_rgb": list(color_rgb) if color_rgb is not None else None,
        "color_hex": color_hex.upper() if color_hex is not None else None,
        "overwrite": bool(overwrite),
    }

    active_bridge = bridge if bridge is not None else PhotoshopBridge()
    jsx = _build_unify_fonts_jsx(config)
    raw = active_bridge.execute_jsx(jsx)
    fields = _parse_protocol(
        raw,
        "SCI_PS_UNIFY_FONTS",
        10,
        "unify_fonts",
        jsx,
    )

    text_layer_count = int(fields[3])
    changed_layer_count = int(fields[4])
    if text_layer_count < 1 or changed_layer_count != text_layer_count:
        raise PhotoshopBridgeError(
            "unify_fonts",
            "Photoshop did not normalize every editable text layer: "
            + repr(fields[3:5]),
            exception_type="TextLayerCountMismatch",
            jsx_preview=_clean_jsx_preview(jsx),
        )
    if not output_file.is_file() or output_file.stat().st_size <= 0:
        raise PhotoshopBridgeError(
            "unify_fonts",
            "Photoshop reported a save but the normalized PSD is missing or empty: "
            + str(output_file),
            exception_type="PsdSaveVerificationError",
            jsx_preview=_clean_jsx_preview(jsx),
        )

    return {
        "document_name": fields[2],
        "text_layer_count": text_layer_count,
        "changed_layer_count": changed_layer_count,
        "font_postscript": fields[5] or None,
        "font_size_pt": float(fields[6]) if fields[6] else None,
        "color_hex": fields[7] or None,
        "input_psd": str(source_file),
        "saved_psd": str(output_file),
        "layer_paths": fields[9].split(";") if fields[9] else [],
        "passed": True,
    }



# ---------------------------------------------------------------------------
# Stage 6: publication export bundle (editable PSD + LZW TIFF + PNG)
# ---------------------------------------------------------------------------

def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _inspect_export_raster(
    path: Path,
    *,
    expected_format: str,
    expected_size: tuple[int, int],
    expected_dpi: float,
) -> Dict[str, Any]:
    """Read export metadata only; Pillow never modifies or re-encodes exports."""
    try:
        from PIL import Image
    except ImportError as exc:
        raise PhotoshopBridgeError(
            "export_figure",
            "Pillow is required to inspect TIFF/PNG export metadata.",
            exception_type="MissingDependency",
        ) from exc

    try:
        with Image.open(str(path)) as image:
            actual_format = str(image.format or "").upper()
            if actual_format != expected_format.upper():
                raise ValueError(
                    "Expected " + expected_format + "; got " + actual_format
                )
            if image.mode != "RGB":
                raise ValueError(
                    expected_format + " must be RGB; got " + str(image.mode)
                )
            if tuple(image.size) != tuple(expected_size):
                raise ValueError(
                    expected_format + " dimensions differ from the source PSD: "
                    + repr(image.size) + " vs " + repr(expected_size)
                )

            dpi_info = image.info.get("dpi")
            if dpi_info is None and expected_format.upper() == "TIFF":
                x_resolution = image.tag_v2.get(282)
                y_resolution = image.tag_v2.get(283)
                resolution_unit = image.tag_v2.get(296, 2)
                if x_resolution is not None and y_resolution is not None:
                    x_resolution = float(x_resolution[0]) / float(x_resolution[1]) if isinstance(x_resolution, tuple) else float(x_resolution)
                    y_resolution = float(y_resolution[0]) / float(y_resolution[1]) if isinstance(y_resolution, tuple) else float(y_resolution)
                    if int(resolution_unit) == 3:
                        x_resolution *= 2.54
                        y_resolution *= 2.54
                    dpi_info = (x_resolution, y_resolution)

            if dpi_info is None:
                raise ValueError(expected_format + " is missing DPI metadata.")
            if isinstance(dpi_info, (int, float)):
                dpi_x = dpi_y = float(dpi_info)
            else:
                dpi_x = float(dpi_info[0])
                dpi_y = float(dpi_info[1])

            if abs(dpi_x - expected_dpi) > 0.5 or abs(dpi_y - expected_dpi) > 0.5:
                raise ValueError(
                    expected_format + " DPI mismatch: "
                    + repr((dpi_x, dpi_y))
                    + " vs " + str(expected_dpi)
                )

            compression = None
            compression_tag = None
            if expected_format.upper() == "TIFF":
                compression_tag = int(image.tag_v2.get(259, 1))
                compression = str(image.info.get("compression", ""))
                if compression_tag != 5:
                    raise ValueError(
                        "TIFF compression tag must be 5 (LZW); got "
                        + str(compression_tag) + " (" + compression + ")"
                    )
                if "lzw" not in compression.lower():
                    # The numeric TIFF tag is authoritative; Pillow versions
                    # differ in the human-readable compression string.
                    compression = "tiff_lzw (verified by tag 259=5)"

            return {
                "path": str(path),
                "format": actual_format,
                "width_px": int(image.width),
                "height_px": int(image.height),
                "mode": image.mode,
                "dpi_x": dpi_x,
                "dpi_y": dpi_y,
                "compression": compression,
                "compression_tag": compression_tag,
                "bytes": path.stat().st_size,
            }
    except Exception as exc:
        raise PhotoshopBridgeError(
            "export_figure",
            "Export metadata verification failed for " + str(path) + ": " + str(exc),
            exception_type=type(exc).__name__,
        ) from exc


def _build_export_figure_jsx(config: Dict[str, Any]) -> str:
    """Build native Photoshop JSX for editable PSD-copy verification and raster exports."""
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function px(value) {
        return value.as("px");
    }
    function modeName(doc) {
        return doc.mode == DocumentMode.RGB ? "RGB" : String(doc.mode);
    }
    function verifyRaster(path, expectedName, expectedWidth, expectedHeight, expectedDpi) {
        var file = new File(path);
        if (!file.exists || file.length <= 0) {
            throw new Error(expectedName + " export is missing or empty: " + path);
        }

        var exportDoc = null;
        try {
            exportDoc = app.open(file);
            var width = px(exportDoc.width);
            var height = px(exportDoc.height);
            var dpi = Number(exportDoc.resolution);
            var mode = modeName(exportDoc);

            if (width !== expectedWidth || height !== expectedHeight) {
                throw new Error(
                    expectedName + " dimensions mismatch: " + width + "x" + height +
                    " vs " + expectedWidth + "x" + expectedHeight
                );
            }
            if (mode !== "RGB") {
                throw new Error(expectedName + " is not RGB: " + mode);
            }
            if (Math.abs(dpi - expectedDpi) > 0.5) {
                throw new Error(
                    expectedName + " resolution mismatch: " + dpi + " vs " + expectedDpi
                );
            }

            exportDoc.close(SaveOptions.DONOTSAVECHANGES);
            exportDoc = null;
            return {
                width: width,
                height: height,
                dpi: dpi,
                mode: mode
            };
        } finally {
            if (exportDoc !== null) {
                try {
                    exportDoc.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeError) {}
            }
        }
    }
    function inspectLayers(container, counts) {
        for (var i = 0; i < container.artLayers.length; i++) {
            var layer = container.artLayers[i];
            if (layer.kind == LayerKind.SMARTOBJECT) {
                counts.smart_objects++;
            }
            if (layer.kind == LayerKind.TEXT) {
                counts.text_layers++;
            }
        }
        for (var groupIndex = 0; groupIndex < container.layerSets.length; groupIndex++) {
            inspectLayers(container.layerSets[groupIndex], counts);
        }
    }

    var cfg = __SCI_EXPORT_FIGURE_CONFIG__;
    var previousDialogs = app.displayDialogs;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var sourceDoc = null;
    var tiffDoc = null;
    var pngDoc = null;
    var response = "";

    try {
        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var sourceFile = new File(cfg.input_path);
        if (!sourceFile.exists) {
            throw new Error("Input PSD does not exist: " + cfg.input_path);
        }
        sourceDoc = app.open(sourceFile);
        if (modeName(sourceDoc) !== "RGB") {
            throw new Error("Publication export requires an RGB PSD; conversion is not performed automatically.");
        }

        var width = px(sourceDoc.width);
        var height = px(sourceDoc.height);
        var dpi = Number(sourceDoc.resolution);
        if (!(Math.abs(dpi - 300) <= 0.5 || Math.abs(dpi - 600) <= 0.5)) {
            throw new Error("Source PSD must be 300 or 600 DPI; got " + dpi + ".");
        }
        if (cfg.expected_dpi !== null && Math.abs(dpi - cfg.expected_dpi) > 0.5) {
            throw new Error(
                "Source PSD resolution does not match requested export DPI: " +
                dpi + " vs " + cfg.expected_dpi
            );
        }

        var counts = {
            panel_groups: 0,
            smart_objects: 0,
            text_layers: 0
        };
        for (var groupIndex = 0; groupIndex < sourceDoc.layerSets.length; groupIndex++) {
            if (sourceDoc.layerSets[groupIndex].name.indexOf("Group_Panel_") === 0) {
                counts.panel_groups++;
            }
        }
        inspectLayers(sourceDoc, counts);
        if (counts.panel_groups < 1 || counts.smart_objects < 1) {
            throw new Error(
                "Source PSD does not contain the expected editable panel/Smart Object structure."
            );
        }

        // TIFF: flatten only a duplicate. The source PSD and its Smart Objects
        // remain open and untouched.
        tiffDoc = sourceDoc.duplicate(cfg.document_name + "_TIFF_EXPORT", false);
        app.activeDocument = tiffDoc;
        tiffDoc.flatten();
        var tiffOptions = new TiffSaveOptions();
        tiffOptions.imageCompression = TIFFEncoding.TIFFLZW;
        tiffOptions.embedColorProfile = true;
        tiffOptions.layers = false;
        tiffOptions.byteOrder = ByteOrder.IBM;
        var tiffFile = new File(cfg.tiff_temp_path);
        tiffDoc.saveAs(tiffFile, tiffOptions, true, Extension.LOWERCASE);
        if (!tiffFile.exists || tiffFile.length <= 0) {
            throw new Error("Photoshop did not create the LZW TIFF.");
        }
        tiffDoc.close(SaveOptions.DONOTSAVECHANGES);
        tiffDoc = null;
        app.activeDocument = sourceDoc;

        // PNG: use another flattened duplicate so the TIFF save operation
        // cannot alter the settings or state of the PNG export.
        pngDoc = sourceDoc.duplicate(cfg.document_name + "_PNG_EXPORT", false);
        app.activeDocument = pngDoc;
        pngDoc.flatten();
        var pngOptions = new PNGSaveOptions();
        pngOptions.interlaced = false;
        var pngFile = new File(cfg.png_temp_path);
        pngDoc.saveAs(pngFile, pngOptions, true, Extension.LOWERCASE);
        if (!pngFile.exists || pngFile.length <= 0) {
            throw new Error("Photoshop did not create the PNG.");
        }
        pngDoc.close(SaveOptions.DONOTSAVECHANGES);
        pngDoc = null;
        app.activeDocument = sourceDoc;

        // Reopen the byte-identical editable PSD copy and both raster exports
        // inside Photoshop before Python checks TIFF LZW and metadata tags.
        var psdCopyFile = new File(cfg.psd_temp_path);
        if (!psdCopyFile.exists || psdCopyFile.length <= 0) {
            throw new Error("Editable PSD copy is missing or empty.");
        }
        var copiedDoc = null;
        try {
            copiedDoc = app.open(psdCopyFile);
            if (px(copiedDoc.width) !== width ||
                px(copiedDoc.height) !== height ||
                modeName(copiedDoc) !== "RGB" ||
                Math.abs(Number(copiedDoc.resolution) - dpi) > 0.5) {
                throw new Error("Editable PSD copy does not match the source document.");
            }
            var copyCounts = {
                panel_groups: 0,
                smart_objects: 0,
                text_layers: 0
            };
            for (var copyGroup = 0; copyGroup < copiedDoc.layerSets.length; copyGroup++) {
                if (copiedDoc.layerSets[copyGroup].name.indexOf("Group_Panel_") === 0) {
                    copyCounts.panel_groups++;
                }
            }
            inspectLayers(copiedDoc, copyCounts);
            if (copyCounts.panel_groups !== counts.panel_groups ||
                copyCounts.smart_objects !== counts.smart_objects ||
                copyCounts.text_layers !== counts.text_layers) {
                throw new Error("Editable PSD copy layer structure differs from source.");
            }
            copiedDoc.close(SaveOptions.DONOTSAVECHANGES);
            copiedDoc = null;
        } finally {
            if (copiedDoc !== null) {
                try {
                    copiedDoc.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeCopyError) {}
            }
        }
        app.activeDocument = sourceDoc;

        var tiffCheck = verifyRaster(
            cfg.tiff_temp_path, "TIFF", width, height, dpi
        );
        app.activeDocument = sourceDoc;
        var pngCheck = verifyRaster(
            cfg.png_temp_path, "PNG", width, height, dpi
        );
        app.activeDocument = sourceDoc;

        sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        sourceDoc = null;

        response = [
            "SCI_PS_EXPORT_FIGURE",
            "OK",
            clean(cfg.document_name),
            String(width),
            String(height),
            String(dpi),
            "RGB",
            String(counts.panel_groups),
            String(counts.smart_objects),
            String(counts.text_layers),
            String(tiffCheck.width),
            String(tiffCheck.height),
            String(tiffCheck.dpi),
            tiffCheck.mode,
            String(pngCheck.width),
            String(pngCheck.height),
            String(pngCheck.dpi),
            pngCheck.mode,
            cfg.tiff_temp_path,
            cfg.png_temp_path
        ].join("\t");

    } catch (e) {
        response = [
            "SCI_PS_EXPORT_FIGURE",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean((e.line === undefined) ? "" : e.line)
        ].join("\t");
    } finally {
        if (tiffDoc !== null) {
            try {
                tiffDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeTiffError) {}
        }
        if (pngDoc !== null) {
            try {
                pngDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closePngError) {}
        }
        if (sourceDoc !== null) {
            try {
                sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeSourceError) {}
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
"""
    return template.replace("__SCI_EXPORT_FIGURE_CONFIG__", payload)


def export_figure(
    input_path: str | Path,
    output_dir: Optional[str | Path] = None,
    *,
    basename: Optional[str] = None,
    expected_dpi: Optional[int] = None,
    overwrite: bool = False,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Export an editable PSD copy, flattened RGB LZW TIFF, and RGB PNG.

    Photoshop performs both raster exports from independent flattened
    duplicates. Pillow is used only to inspect file format, dimensions, RGB
    mode, DPI, and the TIFF compression tag; it does not re-encode outputs.
    """
    source_file = Path(input_path).expanduser().resolve(strict=True)
    if not source_file.is_file() or source_file.suffix.lower() != ".psd":
        raise ValueError("input_path must be an existing PSD file.")

    if expected_dpi is not None and expected_dpi not in (300, 600):
        raise ValueError("expected_dpi must be 300, 600, or None.")

    export_dir = (
        Path(output_dir).expanduser().resolve()
        if output_dir is not None
        else source_file.with_name(source_file.stem + "_exports")
    )
    if basename is None:
        export_basename = source_file.stem
    else:
        export_basename = str(basename).strip()
    invalid_chars = '<>:"/\\\\|?*'
    if (
        not export_basename
        or export_basename in (".", "..")
        or any(char in invalid_chars or ord(char) < 32 for char in export_basename)
    ):
        raise ValueError("basename must be a simple Windows-safe file name.")

    export_dir.mkdir(parents=True, exist_ok=True)
    psd_output = export_dir / (export_basename + ".psd")
    tiff_output = export_dir / (export_basename + ".tif")
    png_output = export_dir / (export_basename + ".png")

    if str(source_file).casefold() in {
        str(psd_output.resolve()).casefold(),
        str(tiff_output.resolve()).casefold(),
        str(png_output.resolve()).casefold(),
    }:
        raise ValueError("Export outputs must not overwrite the input PSD.")

    existing_outputs = [path for path in (psd_output, tiff_output, png_output) if path.exists()]
    if existing_outputs and not overwrite:
        raise FileExistsError(
            "Export output already exists; pass overwrite=True to replace: "
            + ", ".join(str(path) for path in existing_outputs)
        )

    token = uuid.uuid4().hex
    # Keep transient paths short: duplicating a long scientific figure basename
    # here can exceed Windows MAX_PATH even when the final published names fit.
    psd_temp = export_dir / (".sci_ps_" + token + ".tmp.psd")
    tiff_temp = export_dir / (".sci_tif_" + token + ".tmp.tif")
    png_temp = export_dir / (".sci_png_" + token + ".tmp.png")

    input_hash_before = _sha256_file(source_file)
    temp_paths = (psd_temp, tiff_temp, png_temp)
    phase = "create_editable_psd_staging_copy"
    jsx_preview: Optional[str] = None

    try:
        shutil.copy2(source_file, psd_temp)
        if _sha256_file(psd_temp) != input_hash_before:
            raise PhotoshopBridgeError(
                "export_figure",
                "The editable PSD staging copy does not match the input PSD.",
                exception_type="PsdCopyHashMismatch",
            )

        config = {
            "input_path": str(source_file),
            "psd_temp_path": str(psd_temp),
            "tiff_temp_path": str(tiff_temp),
            "png_temp_path": str(png_temp),
            "document_name": export_basename,
            "expected_dpi": expected_dpi,
        }

        active_bridge = bridge if bridge is not None else PhotoshopBridge()
        jsx = _build_export_figure_jsx(config)
        jsx_preview = _clean_jsx_preview(jsx)
        phase = "photoshop_native_tiff_png_export"
        raw = active_bridge.execute_jsx(jsx)
        fields = _parse_protocol(
            raw,
            "SCI_PS_EXPORT_FIGURE",
            20,
            "export_figure",
            jsx,
        )

        width_px = int(float(fields[3]))
        height_px = int(float(fields[4]))
        actual_dpi = float(fields[5])
        if fields[6] != "RGB":
            raise PhotoshopBridgeError(
                "export_figure",
                "Photoshop source mode is not RGB: " + fields[6],
                exception_type="UnsupportedColorMode",
                jsx_preview=_clean_jsx_preview(jsx),
            )
        if expected_dpi is not None and abs(actual_dpi - expected_dpi) > 0.5:
            raise PhotoshopBridgeError(
                "export_figure",
                "Source PSD DPI differs from expected_dpi.",
                exception_type="DpiMismatch",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        phase = "inspect_tiff_metadata"
        tiff_metadata = _inspect_export_raster(
            tiff_temp,
            expected_format="TIFF",
            expected_size=(width_px, height_px),
            expected_dpi=actual_dpi,
        )
        phase = "inspect_png_metadata"
        png_metadata = _inspect_export_raster(
            png_temp,
            expected_format="PNG",
            expected_size=(width_px, height_px),
            expected_dpi=actual_dpi,
        )

        if _sha256_file(source_file) != input_hash_before:
            raise PhotoshopBridgeError(
                "export_figure",
                "The source PSD changed during export; this must never happen.",
                exception_type="SourcePsdModified",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        # Publish the bundle only after all three deliverables and both raster
        # metadata checks have succeeded. os.replace keeps each final path
        # from ever pointing at a partial write.
        phase = "publish_editable_psd"
        os.replace(psd_temp, psd_output)
        phase = "publish_tiff"
        os.replace(tiff_temp, tiff_output)
        phase = "publish_png"
        os.replace(png_temp, png_output)

        psd_hash = _sha256_file(psd_output)
        if psd_hash != input_hash_before:
            raise PhotoshopBridgeError(
                "export_figure",
                "The published editable PSD is not byte-identical to the input.",
                exception_type="PsdCopyHashMismatch",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        return {
            "document_name": fields[2],
            "width_px": width_px,
            "height_px": height_px,
            "dpi": actual_dpi,
            "mode": fields[6],
            "panel_group_count": int(fields[7]),
            "smart_object_count": int(fields[8]),
            "text_layer_count": int(fields[9]),
            "editable_psd": {
                "path": str(psd_output),
                "bytes": psd_output.stat().st_size,
                "sha256_matches_input": True,
            },
            "tiff": {
                **tiff_metadata,
                "path": str(tiff_output),
            },
            "png": {
                **png_metadata,
                "path": str(png_output),
            },
            "input_psd": str(source_file),
            "input_sha256_unchanged": True,
            "passed": True,
        }
    except OSError as exc:
        raise PhotoshopBridgeError(
            "export_figure",
            "Filesystem failure during " + phase + ": " + str(exc),
            exception_type=type(exc).__name__,
            jsx_preview=jsx_preview,
        ) from exc
    finally:
        for temp_path in temp_paths:
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except OSError:
                pass



# ---------------------------------------------------------------------------
# Stage 7: non-destructive publication preflight
# ---------------------------------------------------------------------------

_PREFLIGHT_PANEL_FIELDS = (
    "panel_id",
    "group_name",
    "smart_layer_name",
    "source_width_px",
    "source_height_px",
    "image_left",
    "image_top",
    "image_right",
    "image_bottom",
    "effective_ppi_x",
    "effective_ppi_y",
    "panel_label_count",
    "panel_label_text",
    "panel_label_font",
    "panel_label_size_pt",
    "panel_label_color",
    "panel_label_left",
    "panel_label_top",
    "panel_label_right",
    "panel_label_bottom",
    "scale_bar_count",
    "scale_bar_width_px",
    "scale_bar_height_px",
    "scale_bar_left",
    "scale_bar_top",
    "scale_bar_right",
    "scale_bar_bottom",
    "scale_outline_count",
    "scale_outline_left",
    "scale_outline_top",
    "scale_outline_right",
    "scale_outline_bottom",
    "scale_label_count",
    "scale_label_text",
    "scale_label_font",
    "scale_label_size_pt",
    "scale_label_color",
    "scale_label_left",
    "scale_label_top",
    "scale_label_right",
    "scale_label_bottom",
)


def _build_preflight_figure_jsx(config: Dict[str, Any]) -> str:
    """Build read-only JSX that inspects panels, Smart Objects, and text layers."""
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function px(value) {
        return value.as("px");
    }
    function number(value) {
        return (value === undefined || value === null) ? "" : String(value);
    }
    function encodeRecord(values) {
        var encoded = [];
        for (var i = 0; i < values.length; i++) {
            encoded.push(encodeURIComponent(number(values[i])));
        }
        return encoded.join("|");
    }
    function boundsOf(layer) {
        var b = layer.bounds;
        return [
            px(b[0]),
            px(b[1]),
            px(b[2]),
            px(b[3])
        ];
    }
    function rgbHex(color) {
        function hex(value) {
            var result = Math.round(value).toString(16).toUpperCase();
            return result.length < 2 ? "0" + result : result;
        }
        try {
            return "#" +
                hex(color.rgb.red) +
                hex(color.rgb.green) +
                hex(color.rgb.blue);
        } catch (colorError) {
            return "";
        }
    }
    function textProperties(layer) {
        if (layer === null) {
            return ["", "", "", "", "", "", "", "", ""];
        }
        var item = layer.textItem;
        var b = boundsOf(layer);
        return [
            String(item.contents),
            String(item.font),
            item.size.as("pt"),
            rgbHex(item.color),
            b[0], b[1], b[2], b[3]
        ];
    }
    function collectContainer(container, parentPath, result) {
        for (var i = 0; i < container.artLayers.length; i++) {
            var layer = container.artLayers[i];
            var layerPath = parentPath + "/" + layer.name;
            if (layer.kind == LayerKind.SMARTOBJECT) {
                result.smartObjects.push(layer);
            }
            if (layer.kind == LayerKind.TEXT) {
                var textItem = layer.textItem;
                var b = boundsOf(layer);
                result.textRecords.push([
                    "TEXT",
                    layerPath,
                    layer.name,
                    String(textItem.contents),
                    String(textItem.font),
                    textItem.size.as("pt"),
                    rgbHex(textItem.color),
                    b[0], b[1], b[2], b[3]
                ]);
            }
        }
        for (var groupIndex = 0; groupIndex < container.layerSets.length; groupIndex++) {
            var group = container.layerSets[groupIndex];
            collectContainer(group, parentPath + "/" + group.name, result);
        }
    }
    function findNamedLayers(container, exactName, matches) {
        for (var i = 0; i < container.artLayers.length; i++) {
            var layer = container.artLayers[i];
            if (layer.name === exactName) {
                matches.push(layer);
            }
        }
        for (var groupIndex = 0; groupIndex < container.layerSets.length; groupIndex++) {
            findNamedLayers(container.layerSets[groupIndex], exactName, matches);
        }
    }
    function countSmartObjects(container) {
        var count = 0;
        for (var i = 0; i < container.artLayers.length; i++) {
            if (container.artLayers[i].kind == LayerKind.SMARTOBJECT) {
                count++;
            }
        }
        for (var groupIndex = 0; groupIndex < container.layerSets.length; groupIndex++) {
            count += countSmartObjects(container.layerSets[groupIndex]);
        }
        return count;
    }

    var cfg = __SCI_PREFLIGHT_CONFIG__;
    var previousDialogs = app.displayDialogs;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var previousRulerUnits = null;
    var previousTypeUnits = null;
    var rulerUnitsCaptured = false;
    var typeUnitsCaptured = false;
    var doc = null;
    var smartDoc = null;
    var rows = [];
    var response = "";

    try {
        previousRulerUnits = app.preferences.rulerUnits;
        rulerUnitsCaptured = true;
        previousTypeUnits = app.preferences.typeUnits;
        typeUnitsCaptured = true;

        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var sourceFile = new File(cfg.input_path);
        if (!sourceFile.exists) {
            throw new Error("Input PSD does not exist: " + cfg.input_path);
        }

        doc = app.open(sourceFile);
        var width = px(doc.width);
        var height = px(doc.height);
        var dpi = Number(doc.resolution);
        var mode = doc.mode == DocumentMode.RGB ? "RGB" : String(doc.mode);
        var profile = "";
        try {
            profile = String(doc.colorProfileName);
        } catch (profileError) {}

        var panelGroups = [];
        for (var groupIndex = 0; groupIndex < doc.layerSets.length; groupIndex++) {
            var candidateGroup = doc.layerSets[groupIndex];
            if (candidateGroup.name.indexOf("Group_Panel_") === 0) {
                panelGroups.push(candidateGroup);
            }
        }

        var all = {
            smartObjects: [],
            textRecords: []
        };
        collectContainer(doc, doc.name, all);

        rows.push(encodeRecord([
            "DOC",
            doc.name,
            width,
            height,
            dpi,
            mode,
            profile,
            panelGroups.length,
            all.smartObjects.length,
            all.textRecords.length
        ]));

        for (var panelIndex = 0; panelIndex < panelGroups.length; panelIndex++) {
            var panelGroup = panelGroups[panelIndex];
            var panelId = panelGroup.name.substring("Group_Panel_".length);
            var imageLayers = [];
            var panelTexts = [];
            var panelResult = {
                smartObjects: [],
                textRecords: []
            };
            collectContainer(panelGroup, panelGroup.name, panelResult);

            for (var smartIndex = 0; smartIndex < panelResult.smartObjects.length; smartIndex++) {
                var smartCandidate = panelResult.smartObjects[smartIndex];
                if (smartCandidate.name.indexOf("Image_Panel_") === 0) {
                    imageLayers.push(smartCandidate);
                }
            }

            for (var textIndex = 0; textIndex < panelResult.textRecords.length; textIndex++) {
                panelTexts.push(panelResult.textRecords[textIndex]);
            }

            var imageLayer = imageLayers.length > 0 ? imageLayers[0] : null;
            var sourceWidth = 0;
            var sourceHeight = 0;
            var imageBounds = [0, 0, 0, 0];
            var ppiX = 0;
            var ppiY = 0;
            var smartName = "";

            if (imageLayer !== null) {
                smartName = imageLayer.name;
                imageBounds = boundsOf(imageLayer);

                doc.activeLayer = imageLayer;
                executeAction(
                    stringIDToTypeID("placedLayerEditContents"),
                    new ActionDescriptor(),
                    DialogModes.NO
                );
                smartDoc = app.activeDocument;
                sourceWidth = px(smartDoc.width);
                sourceHeight = px(smartDoc.height);
                smartDoc.close(SaveOptions.DONOTSAVECHANGES);
                smartDoc = null;
                app.activeDocument = doc;

                var displayedWidth = imageBounds[2] - imageBounds[0];
                var displayedHeight = imageBounds[3] - imageBounds[1];
                if (displayedWidth > 0 && displayedHeight > 0) {
                    ppiX = sourceWidth / displayedWidth * dpi;
                    ppiY = sourceHeight / displayedHeight * dpi;
                }
            }

            var panelLabels = [];
            var scaleBars = [];
            var scaleOutlines = [];
            var scaleLabels = [];
            findNamedLayers(panelGroup, "Panel_Label_" + panelId, panelLabels);
            findNamedLayers(panelGroup, "Scale_Bar_" + panelId, scaleBars);
            findNamedLayers(panelGroup, "Scale_Bar_Outline_" + panelId, scaleOutlines);
            findNamedLayers(panelGroup, "Scale_Label_" + panelId, scaleLabels);

            var panelLabelProps = textProperties(
                panelLabels.length > 0 ? panelLabels[0] : null
            );
            var scaleLabelProps = textProperties(
                scaleLabels.length > 0 ? scaleLabels[0] : null
            );
            var barBounds = scaleBars.length > 0 ?
                boundsOf(scaleBars[0]) : [0, 0, 0, 0];
            var outlineBounds = scaleOutlines.length > 0 ?
                boundsOf(scaleOutlines[0]) : [0, 0, 0, 0];

            rows.push(encodeRecord([
                "PANEL",
                panelId,
                panelGroup.name,
                smartName,
                sourceWidth,
                sourceHeight,
                imageBounds[0], imageBounds[1], imageBounds[2], imageBounds[3],
                ppiX, ppiY,
                panelLabels.length,
                panelLabelProps[0], panelLabelProps[1], panelLabelProps[2],
                panelLabelProps[3],
                panelLabelProps[4], panelLabelProps[5],
                panelLabelProps[6], panelLabelProps[7],
                scaleBars.length,
                barBounds[2] - barBounds[0],
                barBounds[3] - barBounds[1],
                barBounds[0], barBounds[1], barBounds[2], barBounds[3],
                scaleOutlines.length,
                outlineBounds[0], outlineBounds[1],
                outlineBounds[2], outlineBounds[3],
                scaleLabels.length,
                scaleLabelProps[0], scaleLabelProps[1], scaleLabelProps[2],
                scaleLabelProps[3],
                scaleLabelProps[4], scaleLabelProps[5],
                scaleLabelProps[6], scaleLabelProps[7]
            ]));
        }

        for (var textRecordIndex = 0; textRecordIndex < all.textRecords.length; textRecordIndex++) {
            rows.push(encodeRecord(all.textRecords[textRecordIndex]));
        }

        doc.close(SaveOptions.DONOTSAVECHANGES);
        doc = null;

        response = [
            "SCI_PS_PREFLIGHT",
            "OK",
            rows.join("\n")
        ].join("\t");

    } catch (e) {
        response = [
            "SCI_PS_PREFLIGHT",
            "ERROR",
            clean(e.name),
            clean(e.message),
            clean((e.line === undefined) ? "" : e.line)
        ].join("\t");
    } finally {
        if (smartDoc !== null) {
            try {
                smartDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (smartCloseError) {}
        }
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
        if (typeUnitsCaptured) {
            try {
                app.preferences.typeUnits = previousTypeUnits;
            } catch (restoreTypeUnitsError) {}
        }
        if (rulerUnitsCaptured) {
            try {
                app.preferences.rulerUnits = previousRulerUnits;
            } catch (restoreRulerUnitsError) {}
        }
        app.displayDialogs = previousDialogs;
    }

    return response;
})();
"""
    return template.replace("__SCI_PREFLIGHT_CONFIG__", payload)


def _parse_preflight_records(encoded_records: str) -> Dict[str, Any]:
    """Decode record-delimited inspection values returned by ExtendScript."""
    parsed: Dict[str, Any] = {
        "document": None,
        "panels": [],
        "text_layers": [],
    }
    for raw_line in encoded_records.split("\n"):
        if not raw_line:
            continue
        values = [unquote(value) for value in raw_line.split("|")]
        if not values:
            continue
        if values[0] == "DOC":
            if len(values) != 10:
                raise ValueError("Invalid DOC preflight record: " + repr(values))
            parsed["document"] = {
                "name": values[1],
                "width_px": int(float(values[2])),
                "height_px": int(float(values[3])),
                "dpi": float(values[4]),
                "mode": values[5],
                "color_profile": values[6],
                "panel_group_count": int(float(values[7])),
                "smart_object_count": int(float(values[8])),
                "text_layer_count": int(float(values[9])),
            }
        elif values[0] == "PANEL":
            if len(values) != 42:
                raise ValueError(
                    "Invalid PANEL preflight record; expected 42 fields, got "
                    + str(len(values))
                )
            panel = dict(zip(_PREFLIGHT_PANEL_FIELDS, values[1:]))
            integer_fields = (
                "source_width_px", "source_height_px",
                "panel_label_count", "scale_bar_count", "scale_outline_count",
                "scale_label_count",
            )
            float_fields = (
                "image_left", "image_top", "image_right", "image_bottom",
                "effective_ppi_x", "effective_ppi_y",
                "panel_label_size_pt", "panel_label_left", "panel_label_top",
                "panel_label_right", "panel_label_bottom",
                "scale_bar_width_px", "scale_bar_height_px",
                "scale_bar_left", "scale_bar_top", "scale_bar_right", "scale_bar_bottom",
                "scale_outline_left", "scale_outline_top",
                "scale_outline_right", "scale_outline_bottom",
                "scale_label_size_pt", "scale_label_left", "scale_label_top",
                "scale_label_right", "scale_label_bottom",
            )
            for key in integer_fields:
                panel[key] = int(float(panel[key] or 0))
            for key in float_fields:
                panel[key] = float(panel[key] or 0)
            parsed["panels"].append(panel)
        elif values[0] == "TEXT":
            if len(values) != 11:
                raise ValueError("Invalid TEXT preflight record: " + repr(values))
            parsed["text_layers"].append({
                "path": values[1],
                "name": values[2],
                "contents": values[3],
                "font_postscript": values[4],
                "size_pt": float(values[5] or 0),
                "color_hex": values[6],
                "left": float(values[7] or 0),
                "top": float(values[8] or 0),
                "right": float(values[9] or 0),
                "bottom": float(values[10] or 0),
            })
        else:
            raise ValueError("Unknown preflight record type: " + values[0])

    if parsed["document"] is None:
        raise ValueError("Photoshop preflight response did not contain a DOC record.")
    return parsed


def _normalize_scale_calibrations(
    calibrations: Optional[Dict[str, Any] | str | Path],
    *,
    default_tolerance_px: float = 2.0,
) -> Dict[str, Dict[str, Any]]:
    if calibrations is None:
        return {}
    if isinstance(calibrations, (str, Path)):
        calibration_path = Path(calibrations).expanduser().resolve(strict=True)
        loaded = json.loads(calibration_path.read_text(encoding="utf-8-sig"))
    else:
        loaded = calibrations
    if not isinstance(loaded, dict):
        raise ValueError("Calibration manifest must be a JSON object keyed by panel ID.")

    normalized: Dict[str, Dict[str, Any]] = {}
    for panel_id, item in loaded.items():
        if not isinstance(item, dict):
            raise ValueError("Calibration for panel " + str(panel_id) + " must be an object.")
        try:
            physical_length = float(item["physical_length"])
            pixel_size = float(item["pixel_size"])
            unit = str(item["unit"]).strip()
            tolerance_px = float(item.get("tolerance_px", default_tolerance_px))
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(
                "Calibration for panel " + str(panel_id) +
                " requires positive physical_length, pixel_size, and unit."
            ) from exc
        if physical_length <= 0 or pixel_size <= 0 or not unit:
            raise ValueError("Calibration values for panel " + str(panel_id) + " must be positive.")
        if tolerance_px < 0:
            raise ValueError("Calibration tolerance_px cannot be negative.")
        normalized[str(panel_id).strip().lower()] = {
            "physical_length": physical_length,
            "pixel_size": pixel_size,
            "unit": unit,
            "tolerance_px": tolerance_px,
        }
    return normalized


def preflight_figure(
    input_path: str | Path,
    *,
    expected_dpi: Optional[int] = None,
    expected_width_cm: Optional[float] = None,
    width_tolerance_cm: float = 0.05,
    min_panel_ppi: float = 300.0,
    require_panel_labels: bool = True,
    panel_label_font: str = "Arial-BoldMT",
    panel_label_min_pt: float = 10.0,
    panel_label_max_pt: float = 12.0,
    figure_text_font: str = "ArialMT",
    figure_text_size_pt: float = 7.0,
    text_size_tolerance_pt: float = 0.5,
    check_text_styles: bool = True,
    calibrations: Optional[Dict[str, Any] | str | Path] = None,
    calibration_default_tolerance_px: float = 2.0,
    export_dir: Optional[str | Path] = None,
    export_basename: Optional[str] = None,
    strict: bool = False,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Inspect publication readiness without saving or modifying the source PSD.

    Hard structural/document errors fail the preflight. Resolution and
    typography deviations are warnings by default; strict=True promotes them
    to failures. Optional calibration data validates physical scale bars, and
    optional export_dir checks a previously generated PSD/TIFF/PNG bundle.
    """
    source_file = Path(input_path).expanduser().resolve(strict=True)
    if not source_file.is_file() or source_file.suffix.lower() != ".psd":
        raise ValueError("input_path must be an existing PSD file.")
    if expected_dpi is not None and expected_dpi not in (300, 600):
        raise ValueError("expected_dpi must be 300, 600, or None.")
    if expected_width_cm is not None and float(expected_width_cm) <= 0:
        raise ValueError("expected_width_cm must be greater than zero.")
    if float(width_tolerance_cm) < 0:
        raise ValueError("width_tolerance_cm cannot be negative.")
    if float(min_panel_ppi) <= 0:
        raise ValueError("min_panel_ppi must be greater than zero.")
    if not 1.0 <= float(panel_label_min_pt) <= float(panel_label_max_pt) <= 200.0:
        raise ValueError("Panel-label point-size range is invalid.")
    if float(text_size_tolerance_pt) < 0:
        raise ValueError("text_size_tolerance_pt cannot be negative.")
    if float(figure_text_size_pt) <= 0:
        raise ValueError("figure_text_size_pt must be greater than zero.")
    if float(calibration_default_tolerance_px) < 0:
        raise ValueError("calibration_default_tolerance_px cannot be negative.")

    calibration_map = _normalize_scale_calibrations(
        calibrations,
        default_tolerance_px=float(calibration_default_tolerance_px),
    )
    if not isinstance(panel_label_font, str) or not panel_label_font.strip():
        raise ValueError("panel_label_font must be a non-empty PostScript font name.")
    if not isinstance(figure_text_font, str) or not figure_text_font.strip():
        raise ValueError("figure_text_font must be a non-empty PostScript font name.")

    export_folder: Optional[Path] = None
    export_stem: Optional[str] = None
    if export_dir is not None:
        export_folder = Path(export_dir).expanduser().resolve(strict=True)
        if not export_folder.is_dir():
            raise ValueError("export_dir must be an existing directory.")
        export_stem = (
            source_file.stem
            if export_basename is None
            else str(export_basename).strip()
        )
        invalid_chars = '<>:"/\\\\|?*'
        if (
            not export_stem
            or export_stem in (".", "..")
            or any(char in invalid_chars or ord(char) < 32 for char in export_stem)
        ):
            raise ValueError("export_basename must be a simple Windows-safe file name.")

    source_hash_before = _sha256_file(source_file)
    active_bridge = bridge if bridge is not None else PhotoshopBridge()
    jsx = _build_preflight_figure_jsx({"input_path": str(source_file)})
    raw = active_bridge.execute_jsx(jsx)
    protocol = _parse_protocol(raw, "SCI_PS_PREFLIGHT", 3, "preflight_figure", jsx)
    inspected = _parse_preflight_records(protocol[2])
    document = inspected["document"]
    panels = inspected["panels"]
    text_layers = inspected["text_layers"]

    checks: list[Dict[str, Any]] = []

    def add_check(
        check_id: str,
        status: str,
        message: str,
        *,
        actual: Any = None,
        expected: Any = None,
    ) -> None:
        final_status = "FAIL" if strict and status == "WARN" else status
        checks.append({
            "id": check_id,
            "status": final_status,
            "message": message,
            "actual": actual,
            "expected": expected,
        })

    width_cm = document["width_px"] * 2.54 / document["dpi"] if document["dpi"] > 0 else 0.0
    add_check(
        "canvas.geometry",
        "PASS" if document["width_px"] > 0 and document["height_px"] > 0 else "FAIL",
        "Canvas dimensions are valid." if document["width_px"] > 0 and document["height_px"] > 0 else "Canvas has invalid dimensions.",
        actual={
            "width_px": document["width_px"],
            "height_px": document["height_px"],
            "width_cm": round(width_cm, 4),
            "height_cm": round(document["height_px"] * 2.54 / document["dpi"], 4) if document["dpi"] > 0 else None,
        },
    )
    dpi_valid = document["dpi"] in (300.0, 600.0)
    dpi_matches = expected_dpi is None or abs(document["dpi"] - expected_dpi) <= 0.5
    add_check(
        "canvas.dpi",
        "PASS" if dpi_valid and dpi_matches else "FAIL",
        "Canvas resolution meets the 300/600 DPI requirement." if dpi_valid and dpi_matches else "Canvas resolution is not an allowed or requested publication DPI.",
        actual=document["dpi"],
        expected=expected_dpi if expected_dpi is not None else [300, 600],
    )
    add_check(
        "canvas.color_mode",
        "PASS" if document["mode"] == "RGB" else "FAIL",
        "Canvas is RGB." if document["mode"] == "RGB" else "Publication figure is not RGB.",
        actual=document["mode"],
        expected="RGB",
    )

    add_check(
        "canvas.color_profile",
        "INFO" if document["color_profile"] else "WARN",
        "Embedded color profile: " + document["color_profile"] if document["color_profile"] else "Photoshop did not report an embedded RGB color profile.",
        actual=document["color_profile"] or None,
        expected="journal/publisher color profile as required",
    )

    canvas_width = document["width_px"]
    canvas_height = document["height_px"]
    for panel in panels:
        image_fits_canvas = (
            panel["image_right"] <= canvas_width + 0.5 and
            panel["image_bottom"] <= canvas_height + 0.5 and
            panel["image_left"] >= -0.5 and
            panel["image_top"] >= -0.5
        )
        add_check(
            "panel_" + panel["panel_id"] + ".canvas_bounds",
            "PASS" if image_fits_canvas else "FAIL",
            "Panel image is contained by the canvas." if image_fits_canvas else "Panel image extends beyond the figure canvas.",
            actual=[
                panel["image_left"], panel["image_top"],
                panel["image_right"], panel["image_bottom"],
            ],
            expected=[0, 0, canvas_width, canvas_height],
        )

    for panel_index, panel in enumerate(panels):
        for other in panels[panel_index + 1:]:
            overlap_width = min(panel["image_right"], other["image_right"]) - max(panel["image_left"], other["image_left"])
            overlap_height = min(panel["image_bottom"], other["image_bottom"]) - max(panel["image_top"], other["image_top"])
            if overlap_width > 0.5 and overlap_height > 0.5:
                add_check(
                    "panel_overlap." + panel["panel_id"] + "_" + other["panel_id"],
                    "WARN",
                    "Panel image bounds overlap; verify that the overlap is intentional.",
                    actual={"overlap_width_px": overlap_width, "overlap_height_px": overlap_height},
                    expected="no unintended panel-image overlap",
                )

    if expected_width_cm is not None:
        width_ok = abs(width_cm - float(expected_width_cm)) <= float(width_tolerance_cm)
        add_check(
            "canvas.width",
            "PASS" if width_ok else "FAIL",
            "Canvas width matches the requested journal geometry." if width_ok else "Canvas width differs from the requested journal geometry.",
            actual=round(width_cm, 4),
            expected={"width_cm": float(expected_width_cm), "tolerance_cm": float(width_tolerance_cm)},
        )
    else:
        known_widths = {"single": 8.5, "half_double": 12.0, "double": 17.6}
        nearest = min(known_widths.items(), key=lambda item: abs(item[1] - width_cm))
        if abs(nearest[1] - width_cm) <= 0.05:
            add_check(
                "canvas.width_preset",
                "PASS",
                "Canvas width matches a configured journal-width preset.",
                actual=round(width_cm, 4),
                expected={"preset": nearest[0], "width_cm": nearest[1]},
            )
        else:
            add_check(
                "canvas.width_preset",
                "INFO",
                "Canvas uses a custom width; no journal-specific target width was supplied.",
                actual=round(width_cm, 4),
                expected="single=8.5 cm, half_double=12 cm, double=17.6 cm or explicit custom width",
            )

    panel_count = document["panel_group_count"]
    add_check(
        "structure.panel_groups",
        "PASS" if panel_count > 0 and len(panels) == panel_count else "FAIL",
        "Panel groups were detected." if panel_count > 0 and len(panels) == panel_count else "Panel-group structure is missing or incomplete.",
        actual=panel_count,
        expected="one inspectable Group_Panel_* per figure panel",
    )
    add_check(
        "structure.smart_objects",
        "PASS" if document["smart_object_count"] >= panel_count and panel_count > 0 else "FAIL",
        "Embedded Smart Objects were detected." if document["smart_object_count"] >= panel_count and panel_count > 0 else "One or more panel Smart Objects are missing.",
        actual=document["smart_object_count"],
        expected="at least one image Smart Object per panel group",
    )

    for panel in panels:
        panel_id = panel["panel_id"]
        prefix = "panel_" + panel_id + "."
        source_w = panel["source_width_px"]
        source_h = panel["source_height_px"]
        image_bounds = (
            panel["image_left"], panel["image_top"],
            panel["image_right"], panel["image_bottom"],
        )
        image_valid = (
            source_w > 0 and source_h > 0 and
            image_bounds[2] > image_bounds[0] and
            image_bounds[3] > image_bounds[1]
        )
        add_check(
            prefix + "smart_object",
            "PASS" if image_valid and panel["smart_layer_name"] else "FAIL",
            "Panel image Smart Object and bounds are valid." if image_valid and panel["smart_layer_name"] else "Panel is missing a readable image Smart Object.",
            actual={
                "layer": panel["smart_layer_name"],
                "source_size_px": [source_w, source_h],
                "display_bounds_px": list(image_bounds),
            },
        )

        ppi_x = panel["effective_ppi_x"]
        ppi_y = panel["effective_ppi_y"]
        if image_valid and ppi_x > 0 and ppi_y > 0:
            effective_ppi = min(ppi_x, ppi_y)
            ppi_status = "PASS" if effective_ppi >= float(min_panel_ppi) else "WARN"
            add_check(
                prefix + "effective_ppi",
                ppi_status,
                "Effective panel resolution meets the requested minimum." if ppi_status == "PASS" else "Effective panel resolution is below the requested minimum; inspect the source raster or reduce its on-canvas size.",
                actual={"x_ppi": round(ppi_x, 2), "y_ppi": round(ppi_y, 2), "limiting_ppi": round(effective_ppi, 2)},
                expected={"minimum_ppi": float(min_panel_ppi)},
            )
        else:
            add_check(
                prefix + "effective_ppi",
                "FAIL",
                "Effective panel PPI could not be calculated.",
                actual={"x_ppi": ppi_x, "y_ppi": ppi_y},
                expected={"minimum_ppi": float(min_panel_ppi)},
            )

        label_count = panel["panel_label_count"]
        if label_count == 0:
            status = "WARN" if require_panel_labels and panel_count > 1 else "INFO"
            add_check(
                prefix + "panel_label",
                status,
                "Panel label is missing." if status == "WARN" else "No panel label is present; this is optional for the current figure.",
                actual=0,
                expected="one editable Panel_Label_<id> text layer" if require_panel_labels and panel_count > 1 else "optional",
            )
        elif label_count > 1:
            add_check(
                prefix + "panel_label",
                "FAIL",
                "More than one panel label layer exists in this panel group.",
                actual=label_count,
                expected=1,
            )
        else:
            label_text = panel["panel_label_text"]
            label_font = panel["panel_label_font"]
            label_size = panel["panel_label_size_pt"]
            label_color = panel["panel_label_color"]
            style_ok = (
                label_font == panel_label_font and
                float(panel_label_min_pt) <= label_size <= float(panel_label_max_pt) and
                label_color.upper() == "#000000"
            )
            add_check(
                prefix + "panel_label_style",
                "PASS" if style_ok else "WARN",
                "Panel label font, size, and color match the figure standard." if style_ok else "Panel label typography differs from the configured standard.",
                actual={"text": label_text, "font": label_font, "size_pt": label_size, "color": label_color},
                expected={"font": panel_label_font, "size_range_pt": [panel_label_min_pt, panel_label_max_pt], "color": "#000000"},
            )
            expected_label_texts = (panel_id, panel_id.upper(), "(" + panel_id + ")")
            label_text_ok = label_text in expected_label_texts
            add_check(
                prefix + "panel_label_content",
                "PASS" if label_text_ok else "WARN",
                "Panel label content follows the panel identifier." if label_text_ok else "Panel label content does not match its panel identifier.",
                actual=label_text,
                expected=list(expected_label_texts),
            )
            label_inside = (
                panel["panel_label_left"] >= image_bounds[0] - 0.5 and
                panel["panel_label_top"] >= image_bounds[1] - 0.5 and
                panel["panel_label_right"] <= image_bounds[2] + 0.5 and
                panel["panel_label_bottom"] <= image_bounds[3] + 0.5
            )
            add_check(
                prefix + "panel_label_bounds",
                "PASS" if label_inside else "WARN",
                "Panel label lies within the image bounds." if label_inside else "Panel label extends outside the panel image bounds.",
                actual=[
                    panel["panel_label_left"], panel["panel_label_top"],
                    panel["panel_label_right"], panel["panel_label_bottom"],
                ],
                expected=list(image_bounds),
            )

        bar_count = panel["scale_bar_count"]
        outline_count = panel["scale_outline_count"]
        scale_label_count = panel["scale_label_count"]
        if bar_count == 0 and scale_label_count == 0 and outline_count == 0:
            add_check(prefix + "scale_bar", "INFO", "No scale bar is present in this panel.")
        elif bar_count != 1 or scale_label_count != 1 or outline_count > 1:
            add_check(
                prefix + "scale_bar",
                "FAIL",
                "Scale-bar components are incomplete or duplicated.",
                actual={"bar": bar_count, "outline": outline_count, "label": scale_label_count},
                expected={"bar": 1, "outline": "0 or 1", "label": 1},
            )
        else:
            bar_bounds = (
                panel["scale_bar_left"], panel["scale_bar_top"],
                panel["scale_bar_right"], panel["scale_bar_bottom"],
            )
            bar_inside = (
                panel["scale_bar_width_px"] > 1 and
                panel["scale_bar_height_px"] > 1 and
                bar_bounds[0] >= image_bounds[0] - 0.5 and
                bar_bounds[1] >= image_bounds[1] - 0.5 and
                bar_bounds[2] <= image_bounds[2] + 0.5 and
                bar_bounds[3] <= image_bounds[3] + 0.5
            )
            add_check(
                prefix + "scale_bar_geometry",
                "PASS" if bar_inside else "FAIL",
                "Scale-bar rectangle is inside the panel image." if bar_inside else "Scale-bar rectangle has invalid dimensions or extends outside its panel.",
                actual={"width_px": panel["scale_bar_width_px"], "height_px": panel["scale_bar_height_px"], "bounds": list(bar_bounds)},
                expected="positive rectangle fully inside panel image bounds",
            )
            if outline_count == 1:
                outline_inside = (
                    panel["scale_outline_left"] <= bar_bounds[0] + 0.5 and
                    panel["scale_outline_top"] <= bar_bounds[1] + 0.5 and
                    panel["scale_outline_right"] >= bar_bounds[2] - 0.5 and
                    panel["scale_outline_bottom"] >= bar_bounds[3] - 0.5
                )
                add_check(
                    prefix + "scale_bar_outline",
                    "PASS" if outline_inside else "WARN",
                    "Contrasting outline encloses the scale bar." if outline_inside else "Scale-bar outline does not fully enclose the bar.",
                    actual=[
                        panel["scale_outline_left"], panel["scale_outline_top"],
                        panel["scale_outline_right"], panel["scale_outline_bottom"],
                    ],
                    expected="outline contains the bar rectangle",
                )

            scale_bounds = (
                panel["scale_label_left"], panel["scale_label_top"],
                panel["scale_label_right"], panel["scale_label_bottom"],
            )
            label_inside = (
                scale_bounds[0] >= image_bounds[0] - 0.5 and
                scale_bounds[1] >= image_bounds[1] - 0.5 and
                scale_bounds[2] <= image_bounds[2] + 0.5 and
                scale_bounds[3] <= image_bounds[3] + 0.5
            )
            add_check(
                prefix + "scale_label_bounds",
                "PASS" if label_inside else "WARN",
                "Scale label lies within panel bounds." if label_inside else "Scale label extends outside its panel.",
                actual=list(scale_bounds),
                expected=list(image_bounds),
            )
            label_above_bar = panel["scale_label_bottom"] <= panel["scale_bar_top"] + 0.5
            add_check(
                prefix + "scale_label_position",
                "PASS" if label_above_bar else "WARN",
                "Scale label is positioned above the bar." if label_above_bar else "Scale label overlaps or falls below the scale bar.",
                actual={"label_bottom": panel["scale_label_bottom"], "bar_top": panel["scale_bar_top"]},
                expected="label_bottom <= bar_top",
            )

            calibration = calibration_map.get(panel_id.lower())
            if calibration is None:
                add_check(
                    prefix + "scale_calibration",
                    "INFO",
                    "Scale-bar geometry was checked, but physical calibration was not independently verified; supply a calibration manifest.",
                    actual={"label": panel["scale_label_text"], "bar_width_px": panel["scale_bar_width_px"]},
                    expected="physical_length, pixel_size, unit",
                )
            elif source_w <= 0:
                add_check(
                    prefix + "scale_calibration",
                    "FAIL",
                    "Scale calibration cannot be checked because the Smart Object source width is unavailable.",
                )
            else:
                displayed_width = image_bounds[2] - image_bounds[0]
                expected_bar_width = (
                    calibration["physical_length"] / calibration["pixel_size"]
                ) * displayed_width / source_w
                tolerance = calibration.get(
                    "tolerance_px", float(calibration_default_tolerance_px)
                )
                actual_bar_width = panel["scale_bar_width_px"]
                width_ok = abs(actual_bar_width - expected_bar_width) <= tolerance
                expected_label = (
                    _format_scale_number(calibration["physical_length"]) +
                    " " + calibration["unit"]
                )
                label_ok = panel["scale_label_text"] == expected_label
                add_check(
                    prefix + "scale_calibration_width",
                    "PASS" if width_ok else "FAIL",
                    "Measured scale-bar width matches the supplied physical calibration." if width_ok else "Scale-bar width differs from the supplied physical calibration.",
                    actual={"bar_width_px": actual_bar_width, "absolute_error_px": round(abs(actual_bar_width - expected_bar_width), 4)},
                    expected={"bar_width_px": expected_bar_width, "tolerance_px": tolerance},
                )
                add_check(
                    prefix + "scale_calibration_label",
                    "PASS" if label_ok else "FAIL",
                    "Scale-bar label matches the supplied physical calibration." if label_ok else "Scale-bar label differs from the supplied physical calibration.",
                    actual=panel["scale_label_text"],
                    expected=expected_label,
                )

    if check_text_styles:
        for text_layer in text_layers:
            name = text_layer["name"]
            text_inside_canvas = (
                text_layer["left"] >= -0.5 and
                text_layer["top"] >= -0.5 and
                text_layer["right"] <= document["width_px"] + 0.5 and
                text_layer["bottom"] <= document["height_px"] + 0.5
            )
            add_check(
                "text_bounds." + text_layer["path"],
                "PASS" if text_inside_canvas else "WARN",
                "Text layer is within the canvas." if text_inside_canvas else "Text layer extends beyond the canvas.",
                actual=[
                    text_layer["left"], text_layer["top"],
                    text_layer["right"], text_layer["bottom"],
                ],
                expected=[0, 0, document["width_px"], document["height_px"]],
            )
            if name.startswith("Panel_Label_"):
                continue
            if text_layer["color_hex"].upper() != "#000000":
                add_check(
                    "text_color." + text_layer["path"],
                    "WARN",
                    "Text color differs from black.",
                    actual=text_layer["color_hex"],
                    expected="#000000",
                )
            font_ok = text_layer["font_postscript"] == figure_text_font
            size_ok = abs(text_layer["size_pt"] - float(figure_text_size_pt)) <= float(text_size_tolerance_pt)
            add_check(
                "text_style." + text_layer["path"],
                "PASS" if font_ok and size_ok else "WARN",
                "Figure text matches the configured font and point size." if font_ok and size_ok else "Figure text differs from the configured font or point size.",
                actual={"font": text_layer["font_postscript"], "size_pt": text_layer["size_pt"]},
                expected={"font": figure_text_font, "size_pt": float(figure_text_size_pt), "tolerance_pt": float(text_size_tolerance_pt)},
            )

    if any(
        panel["scale_bar_count"] > 0
        for panel in panels
    ) and not calibration_map:
        add_check(
            "scale_calibration.manifest",
            "INFO",
            "No physical calibration manifest was supplied; scale-bar units cannot be inferred from the PSD alone.",
        )

    if export_folder is None:
        add_check(
            "exports.metadata",
            "INFO",
            "Export-bundle metadata was not checked; supply export_dir to inspect a PSD/TIFF/PNG bundle.",
        )
    else:
        psd_export = export_folder / (str(export_stem) + ".psd")
        tiff_export = export_folder / (str(export_stem) + ".tif")
        png_export = export_folder / (str(export_stem) + ".png")
        try:
            psd_matches = psd_export.is_file() and _sha256_file(psd_export) == source_hash_before
            add_check(
                "exports.editable_psd",
                "PASS" if psd_matches else "FAIL",
                "Editable PSD export is byte-identical to the preflight source." if psd_matches else "Editable PSD export is missing or differs from the preflight source.",
                actual={"path": str(psd_export), "exists": psd_export.is_file()},
                expected="SHA-256 identical to input PSD",
            )
        except OSError as exc:
            add_check("exports.editable_psd", "FAIL", "Could not inspect exported PSD: " + str(exc))
        for fmt, raster_path, expected_format in (
            ("tiff", tiff_export, "TIFF"),
            ("png", png_export, "PNG"),
        ):
            try:
                metadata = _inspect_export_raster(
                    raster_path,
                    expected_format=expected_format,
                    expected_size=(document["width_px"], document["height_px"]),
                    expected_dpi=document["dpi"],
                )
                if fmt == "tiff":
                    add_check(
                        "exports.tiff",
                        "PASS",
                        "TIFF has matching dimensions, RGB mode, DPI, and LZW compression.",
                        actual=metadata,
                        expected={"mode": "RGB", "dpi": document["dpi"], "compression_tag": 5},
                    )
                else:
                    add_check(
                        "exports.png",
                        "PASS",
                        "PNG has matching dimensions, RGB mode, and DPI metadata.",
                        actual=metadata,
                        expected={"mode": "RGB", "dpi": document["dpi"]},
                    )
            except Exception as exc:
                add_check(
                    "exports." + fmt,
                    "FAIL",
                    "Export metadata check failed: " + str(exc),
                    actual=str(raster_path),
                )

    source_hash_unchanged = _sha256_file(source_file) == source_hash_before
    add_check(
        "source.integrity",
        "PASS" if source_hash_unchanged else "FAIL",
        "Source PSD was not modified during preflight." if source_hash_unchanged else "Source PSD changed during the read-only preflight.",
        actual={"sha256_unchanged": source_hash_unchanged},
        expected=True,
    )

    failure_count = sum(1 for item in checks if item["status"] == "FAIL")
    warning_count = sum(1 for item in checks if item["status"] == "WARN")
    return {
        "document": document,
        "canvas_width_cm": round(width_cm, 4),
        "panels": panels,
        "text_layers": text_layers,
        "checks": checks,
        "passed": failure_count == 0,
        "strict": bool(strict),
        "warning_count": warning_count,
        "failure_count": failure_count,
        "input_psd": str(source_file),
        "input_sha256_unchanged": source_hash_unchanged,
    }




# ---------------------------------------------------------------------------
# Stage 8A: inspect an existing PSD without modifying it
# ---------------------------------------------------------------------------


def _build_inspect_psd_jsx(config: Dict[str, Any]) -> str:
    """Build read-only JSX that enumerates the mixed, ordered Photoshop layer tree."""
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function encodeRecord(values) {
        var encoded = [];
        for (var i = 0; i < values.length; i++) {
            encoded.push(encodeURIComponent(
                values[i] === undefined || values[i] === null ? "" : String(values[i])
            ));
        }
        return encoded.join("|");
    }
    function safeProperty(object, propertyName, fallback) {
        try {
            var value = object[propertyName];
            return value === undefined || value === null ? fallback : value;
        } catch (e) {
            return fallback;
        }
    }
    function px(value) {
        try {
            return value.as("px");
        } catch (e) {
            return "";
        }
    }
    function layerBounds(layer) {
        try {
            var b = layer.bounds;
            return [px(b[0]), px(b[1]), px(b[2]), px(b[3])];
        } catch (e) {
            return ["", "", "", ""];
        }
    }
    function rgbHex(color) {
        function hex(value) {
            var result = Math.round(value).toString(16).toUpperCase();
            return result.length < 2 ? "0" + result : result;
        }
        try {
            return "#" + hex(color.rgb.red) + hex(color.rgb.green) + hex(color.rgb.blue);
        } catch (e) {
            return "";
        }
    }
    function inspectText(layer) {
        if (String(safeProperty(layer, "kind", "")) !== String(LayerKind.TEXT)) {
            return ["", "", "", ""];
        }
        try {
            var item = layer.textItem;
            return [
                String(item.contents),
                String(item.font),
                item.size.as("pt"),
                rgbHex(item.color)
            ];
        } catch (e) {
            return ["", "", "", ""];
        }
    }

    var cfg = __SCI_INSPECT_CONFIG__;
    var previousDialogs = app.displayDialogs;
    var previousRulerUnits = app.preferences.rulerUnits;
    var previousTypeUnits = app.preferences.typeUnits;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var doc = null;
    var rows = [];
    var layerCount = 0;

    try {
        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var inputFile = new File(cfg.input_path);
        if (!inputFile.exists) {
            throw new Error("Input PSD does not exist: " + cfg.input_path);
        }
        doc = app.open(inputFile);

        var width = doc.width.as("px");
        var height = doc.height.as("px");
        var dpi = Number(doc.resolution);
        var mode = doc.mode == DocumentMode.RGB ? "RGB" : String(doc.mode);
        var profile = "";
        try {
            profile = String(doc.colorProfileName);
        } catch (profileError) {}
        rows.push(encodeRecord([
            "DOC", cfg.document_label, width, height, dpi, mode, profile
        ]));

        function walk(container, parentKey, parentPath, depth) {
            // Document.layers / LayerSet.layers preserves the interleaved
            // front-to-back order of ArtLayers and LayerSets.
            var collection = container.layers;
            for (var i = 0; i < collection.length; i++) {
                var layer = collection[i];
                var isGroup = layer.typename === "LayerSet";
                var nodeKey = parentKey + "/" + (isGroup ? "G" : "L") + i;
                var layerPath = parentPath + "/" + layer.name;
                var bounds = layerBounds(layer);
                var text = isGroup ? ["", "", "", ""] : inspectText(layer);
                var kind = isGroup ? "LayerSet" : String(safeProperty(layer, "kind", "unknown"));

                rows.push(encodeRecord([
                    "LAYER",
                    nodeKey,
                    parentKey,
                    i,
                    depth,
                    layer.typename,
                    layer.name,
                    kind,
                    safeProperty(layer, "visible", true),
                    safeProperty(layer, "opacity", 100),
                    safeProperty(layer, "blendMode", ""),
                    safeProperty(layer, "isBackgroundLayer", false),
                    safeProperty(layer, "allLocked", false),
                    safeProperty(layer, "positionLocked", false),
                    safeProperty(layer, "pixelsLocked", false),
                    safeProperty(layer, "grouped", false),
                    bounds[0], bounds[1], bounds[2], bounds[3],
                    text[0], text[1], text[2], text[3],
                    layerPath
                ]));
                layerCount++;

                if (isGroup) {
                    walk(layer, nodeKey, layerPath, depth + 1);
                }
            }
        }

        walk(doc, "root", cfg.document_label, 0);
        rows.unshift(encodeRecord(["COUNT", layerCount]));

        doc.close(SaveOptions.DONOTSAVECHANGES);
        doc = null;
        return ["SCI_PS_INSPECT_PSD", "OK", rows.join("\n")].join("\t");
    } catch (e) {
        return [
            "SCI_PS_INSPECT_PSD",
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
            } catch (restoreDocError) {}
        }
        try {
            app.preferences.rulerUnits = previousRulerUnits;
        } catch (restoreRulerError) {}
        try {
            app.preferences.typeUnits = previousTypeUnits;
        } catch (restoreTypeError) {}
        app.displayDialogs = previousDialogs;
    }
})();
"""
    return template.replace("__SCI_INSPECT_CONFIG__", payload)


def _parse_psd_inspection_records(encoded_records: str) -> Dict[str, Any]:
    """Decode a read-only PSD inspection payload into an ordered layer tree."""
    parsed: Dict[str, Any] = {
        "document": None,
        "layers": [],
        "layer_count": 0,
    }

    def number(value: str, default: Optional[float] = None) -> Optional[float]:
        if value == "":
            return default
        return float(value)

    def boolean(value: str) -> bool:
        return str(value).lower() == "true"

    for raw_line in encoded_records.split("\n"):
        if not raw_line:
            continue
        values = [unquote(value) for value in raw_line.split("|")]
        if values[0] == "COUNT":
            if len(values) != 2:
                raise ValueError("Invalid COUNT inspection record: " + repr(values))
            parsed["layer_count"] = int(values[1])
        elif values[0] == "DOC":
            if len(values) != 7:
                raise ValueError("Invalid DOC inspection record: " + repr(values))
            parsed["document"] = {
                "name": values[1],
                "width_px": int(float(values[2])),
                "height_px": int(float(values[3])),
                "dpi": float(values[4]),
                "mode": values[5],
                "color_profile": values[6],
            }
        elif values[0] == "LAYER":
            if len(values) != 25:
                raise ValueError(
                    "Invalid LAYER inspection record; expected 25 fields, got "
                    + str(len(values))
                )
            parsed["layers"].append({
                "node_key": values[1],
                "parent_key": values[2],
                "stack_index": int(values[3]),
                "depth": int(values[4]),
                "typename": values[5],
                "name": values[6],
                "kind": values[7],
                "visible": boolean(values[8]),
                "opacity": float(values[9]),
                "blend_mode": values[10],
                "is_background": boolean(values[11]),
                "all_locked": boolean(values[12]),
                "position_locked": boolean(values[13]),
                "pixels_locked": boolean(values[14]),
                "clipped_to_below": boolean(values[15]),
                "bounds_px": [
                    number(values[16]), number(values[17]),
                    number(values[18]), number(values[19]),
                ],
                "text": {
                    "contents": values[20],
                    "font_postscript": values[21],
                    "size_pt": number(values[22]),
                    "color_hex": values[23],
                } if values[7] == "LayerKind.TEXT" else None,
                "path": values[24],
            })
        else:
            raise ValueError("Unknown PSD inspection record type: " + values[0])

    if parsed["document"] is None:
        raise ValueError("Photoshop PSD inspection did not contain a DOC record.")
    if parsed["layer_count"] != len(parsed["layers"]):
        raise ValueError(
            "PSD inspection layer count mismatch: "
            + str(parsed["layer_count"]) + " vs " + str(len(parsed["layers"]))
        )

    keys = [layer["node_key"] for layer in parsed["layers"]]
    if len(keys) != len(set(keys)):
        raise ValueError("PSD inspection returned duplicate node keys.")
    known_keys = {"root", *keys}
    layer_by_key = {layer["node_key"]: layer for layer in parsed["layers"]}
    siblings: Dict[tuple[str, int], str] = {}
    for layer in parsed["layers"]:
        parent_key = layer["parent_key"]
        node_key = layer["node_key"]
        if parent_key not in known_keys:
            raise ValueError(
                "PSD inspection returned an unknown parent key for " + node_key
            )
        if node_key.rsplit("/", 1)[0] != parent_key:
            raise ValueError("PSD inspection returned an inconsistent node path: " + node_key)
        if parent_key != "root":
            parent_layer = layer_by_key[parent_key]
            if parent_layer["typename"] != "LayerSet":
                raise ValueError("PSD inspection parent is not a LayerSet: " + node_key)
            if layer["depth"] != parent_layer["depth"] + 1:
                raise ValueError("PSD inspection returned an inconsistent depth: " + node_key)
        elif layer["depth"] != 0:
            raise ValueError("Top-level PSD layer has a non-zero depth: " + node_key)
        sibling_key = (parent_key, layer["stack_index"])
        if sibling_key in siblings:
            raise ValueError(
                "PSD inspection returned duplicate sibling indices: "
                + siblings[sibling_key] + " and " + node_key
            )
        siblings[sibling_key] = node_key
    return parsed


def inspect_psd(
    input_path: str | Path,
    *,
    report_path: Optional[str | Path] = None,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Inspect an existing PSD read-only and return its full ordered layer tree."""
    source_file = Path(input_path).expanduser().resolve(strict=True)
    if not source_file.is_file() or source_file.suffix.lower() != ".psd":
        raise ValueError("input_path must be an existing PSD file.")

    report_file: Optional[Path] = None
    if report_path is not None:
        report_file = Path(report_path).expanduser().resolve()
        if report_file.suffix.lower() != ".json":
            raise ValueError("report_path must end in .json.")
        if report_file.exists():
            raise FileExistsError(
                "Inspection report already exists; choose another path: "
                + str(report_file)
            )
        report_file.parent.mkdir(parents=True, exist_ok=True)

    source_hash_before = _sha256_file(source_file)
    active_bridge = bridge if bridge is not None else PhotoshopBridge()
    token = uuid.uuid4().hex[:16]
    temp_input = Path(os.environ.get("TEMP", os.environ.get("TMP", "."))) / (
        "sci_insp_" + token + ".psd"
    )
    jsx = ""
    try:
        # Inspect a byte-identical temporary copy so an already-open, unsaved
        # Photoshop document cannot silently change the on-disk inspection basis.
        shutil.copy2(source_file, temp_input)
        if _sha256_file(temp_input) != source_hash_before:
            raise PhotoshopBridgeError(
                "inspect_psd",
                "Temporary inspection copy is not byte-identical to the source PSD.",
                exception_type="WorkingCopyIntegrityError",
            )

        jsx = _build_inspect_psd_jsx({
            "input_path": str(temp_input),
            "document_label": source_file.name,
        })
        raw = active_bridge.execute_jsx(jsx)
        protocol = _parse_protocol(raw, "SCI_PS_INSPECT_PSD", 3, "inspect_psd", jsx)
        inspected = _parse_psd_inspection_records(protocol[2])

        source_hash_after = _sha256_file(source_file)
        if source_hash_after != source_hash_before:
            raise PhotoshopBridgeError(
                "inspect_psd",
                "The source PSD SHA-256 changed during read-only inspection.",
                exception_type="SourceIntegrityError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        fingerprint_payload = [
            {
                "node_key": layer["node_key"],
                "parent_key": layer["parent_key"],
                "typename": layer["typename"],
                "name": layer["name"],
                "kind": layer["kind"],
                "stack_index": layer["stack_index"],
                "bounds_px": layer["bounds_px"],
                "visible": layer["visible"],
            }
            for layer in inspected["layers"]
        ]
        fingerprint = hashlib.sha256(
            json.dumps(
                fingerprint_payload,
                ensure_ascii=False,
                sort_keys=True,
            ).encode("utf-8")
        ).hexdigest()

        if _sha256_file(source_file) != source_hash_before:
            raise PhotoshopBridgeError(
                "inspect_psd",
                "The source PSD changed before inspection completed.",
                exception_type="SourceIntegrityError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        report = {
            "schema_version": 1,
            "passed": True,
            "input_psd": str(source_file),
            "source_sha256": source_hash_before,
            "layer_tree_sha256": fingerprint,
            "document": inspected["document"],
            "layer_count": inspected["layer_count"],
            "group_count": sum(
                1 for layer in inspected["layers"]
                if layer["typename"] == "LayerSet"
            ),
            "text_layer_count": sum(
                1 for layer in inspected["layers"]
                if layer["kind"] == "LayerKind.TEXT"
            ),
            "smart_object_count": sum(
                1 for layer in inspected["layers"]
                if layer["kind"] == "LayerKind.SMARTOBJECT"
            ),
            "layers": inspected["layers"],
            "source_sha256_unchanged": True,
        }

        if report_file is not None:
            report["report_path"] = str(report_file)
            temp_report = report_file.with_name(report_file.stem + ".tmp" + report_file.suffix)
            try:
                temp_report.write_text(
                    json.dumps(report, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                os.replace(temp_report, report_file)
            finally:
                if temp_report.exists():
                    temp_report.unlink()

        return report
    finally:
        try:
            if temp_input.exists():
                temp_input.unlink()
        except OSError:
            pass




# ---------------------------------------------------------------------------
# Stage 8B: approved, non-destructive existing-PSD organization plans
# ---------------------------------------------------------------------------


def _load_organize_plan(plan: Dict[str, Any] | str | Path) -> Dict[str, Any]:
    if isinstance(plan, (str, Path)):
        plan_path = Path(plan).expanduser().resolve(strict=True)
        if not plan_path.is_file() or plan_path.suffix.lower() != ".json":
            raise ValueError("plan_path must be an existing JSON file.")
        loaded = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    elif isinstance(plan, dict):
        loaded = json.loads(json.dumps(plan, ensure_ascii=False))
    else:
        raise ValueError("plan must be a dictionary or a JSON file path.")
    if not isinstance(loaded, dict):
        raise ValueError("Organizer plan must be a JSON object.")
    return loaded


def _valid_layer_name(value: Any, field: str) -> str:
    if not isinstance(value, str):
        raise ValueError(field + " must be a string.")
    name = value.strip()
    if not name or len(name) > 255 or any(ord(char) < 32 for char in name):
        raise ValueError(field + " must be 1–255 printable characters.")
    return name


def _validate_organize_plan(
    plan: Dict[str, Any],
    inspection: Dict[str, Any],
) -> Dict[str, Any]:
    """Validate source identity, exact layer references, and safe structure-only operations."""
    if plan.get("schema_version") != 1:
        raise ValueError("Unsupported organizer plan schema_version; expected 1.")
    if plan.get("source_sha256") != inspection.get("source_sha256"):
        raise ValueError("Organizer plan source_sha256 does not match the current PSD.")
    if plan.get("layer_tree_sha256") != inspection.get("layer_tree_sha256"):
        raise ValueError("Organizer plan layer_tree_sha256 does not match the current PSD.")

    source_layers = {
        layer["node_key"]: layer for layer in inspection.get("layers", [])
    }

    def has_locked_ancestor(key: str) -> bool:
        current = source_layers[key]["parent_key"]
        while current != "root":
            parent_layer = source_layers[current]
            if parent_layer["all_locked"] or parent_layer["position_locked"]:
                return True
            current = parent_layer["parent_key"]
        return False

    def participates_in_clipping_chain(key: str) -> bool:
        layer = source_layers[key]
        if layer["clipped_to_below"]:
            return True
        siblings = sorted(
            (
                item for item in source_layers.values()
                if item["parent_key"] == layer["parent_key"]
            ),
            key=lambda item: item["stack_index"],
        )
        position = next(
            index for index, item in enumerate(siblings)
            if item["node_key"] == key
        )
        # In Photoshop's front-to-back layer order, clipped layers sit above
        # the base layer they clip to. Moving/reordering that base can detach
        # the clipping relationship just as moving a clipped layer can.
        index = position - 1
        while index >= 0 and siblings[index]["clipped_to_below"]:
            return True
        return False
    operations = plan.get("operations")
    if not isinstance(operations, list) or not operations:
        raise ValueError("Organizer plan must contain a non-empty operations array.")

    existing_keys = set(source_layers)
    group_parents: Dict[str, str] = {}
    normalized: list[Dict[str, Any]] = []
    source_ref_keys: set[str] = set()
    renamed: set[str] = set()
    moved: set[str] = set()
    reordered: set[str] = set()
    visibility_set: set[str] = set()
    non_create_started = False

    def require_layer(key: Any, field: str) -> Dict[str, Any]:
        if not isinstance(key, str) or key not in source_layers:
            raise ValueError(field + " must reference an inspected source node_key.")
        source_ref_keys.add(key)
        return source_layers[key]

    def require_group_parent(key: Any, field: str) -> str:
        if key == "root":
            return "root"
        if key in group_parents:
            return key
        layer = require_layer(key, field)
        if layer["typename"] != "LayerSet":
            raise ValueError(field + " must reference root, a LayerSet, or an earlier created group.")
        if layer["all_locked"] or layer["position_locked"] or has_locked_ancestor(key):
            raise ValueError(field + " cannot target a fully or position-locked LayerSet or ancestor.")
        return key

    def destination_is_inside_target(destination: str, target_key: str) -> bool:
        current = destination
        while current != "root":
            if current == target_key:
                return True
            if current in group_parents:
                current = group_parents[current]
            else:
                parent_layer = source_layers.get(current)
                if parent_layer is None:
                    return False
                current = parent_layer["parent_key"]
        return False

    for index, raw_operation in enumerate(operations):
        if not isinstance(raw_operation, dict):
            raise ValueError("Organizer operation " + str(index) + " must be an object.")
        operation = dict(raw_operation)
        action = operation.get("action")

        if action == "create_group":
            if non_create_started:
                raise ValueError("create_group operations must appear before layer operations.")
            key = operation.get("key")
            if (
                not isinstance(key, str)
                or not key
                or not key[0].isalpha()
                or not key.replace("_", "").replace("-", "").isalnum()
                or len(key) > 64
                or key in group_parents
                or key in existing_keys
            ):
                raise ValueError("create_group key must be unique and use letters, numbers, '_' or '-'.")
            parent = require_group_parent(operation.get("parent", "root"), "create_group.parent")
            name = _valid_layer_name(operation.get("name"), "create_group.name")
            group_parents[key] = parent
            operation["parent"] = parent
            operation["name"] = name

        else:
            non_create_started = True
            if action == "rename_layer":
                target = operation.get("target")
                layer = require_layer(target, "rename_layer.target")
                if target in renamed:
                    raise ValueError("A layer may be renamed only once in one plan: " + target)
                if layer["is_background"]:
                    raise ValueError("Background layers cannot be renamed by organizer v1: " + target)
                expected_name = operation.get("expected_name")
                if expected_name != layer["name"]:
                    raise ValueError("rename_layer.expected_name no longer matches " + target)
                operation["new_name"] = _valid_layer_name(
                    operation.get("new_name"), "rename_layer.new_name"
                )
                renamed.add(target)

            elif action == "move_layer":
                target = operation.get("target")
                layer = require_layer(target, "move_layer.target")
                if target in moved or target in reordered:
                    raise ValueError("A layer may be moved/reordered only once in one plan: " + target)
                if (
                    layer["is_background"] or layer["all_locked"]
                    or layer["position_locked"] or has_locked_ancestor(target)
                ):
                    raise ValueError("Locked/background layer or a child of a locked group cannot be moved: " + target)
                if participates_in_clipping_chain(target):
                    raise ValueError("Layers participating in clipping-mask chains cannot be moved in organizer v1: " + target)
                if operation.get("expected_name") != layer["name"]:
                    raise ValueError("move_layer.expected_name no longer matches " + target)
                destination = operation.get("destination")
                if destination == "root":
                    destination = "root"
                elif destination in group_parents:
                    destination = destination
                else:
                    destination_layer = require_layer(destination, "move_layer.destination")
                    if destination_layer["typename"] != "LayerSet":
                        raise ValueError("move_layer.destination must be root, a LayerSet, or a created group.")
                    if destination_layer["all_locked"] or destination_layer["position_locked"]:
                        raise ValueError("move_layer.destination cannot be a fully or position-locked LayerSet.")
                if layer["typename"] == "LayerSet" and destination_is_inside_target(
                    destination, target
                ):
                    raise ValueError("Cannot move a LayerSet into itself or one of its descendants.")
                if operation.get("position", "end") not in ("start", "end"):
                    raise ValueError("move_layer.position must be 'start' or 'end'.")
                operation["destination"] = destination
                operation["position"] = operation.get("position", "end")
                moved.add(target)

            elif action == "reorder_layer":
                target = operation.get("target")
                layer = require_layer(target, "reorder_layer.target")
                relative = operation.get("relative_to")
                relative_layer = require_layer(relative, "reorder_layer.relative_to")
                if target == relative:
                    raise ValueError("reorder_layer.target and relative_to must differ.")
                if target in moved or target in reordered:
                    raise ValueError("A layer may be moved/reordered only once in one plan: " + target)
                if (
                    layer["is_background"] or layer["all_locked"]
                    or layer["position_locked"] or has_locked_ancestor(target)
                ):
                    raise ValueError("Locked/background layer or a child of a locked group cannot be reordered: " + target)
                if participates_in_clipping_chain(target):
                    raise ValueError("Layers participating in clipping-mask chains cannot be reordered in organizer v1: " + target)
                if operation.get("expected_name") != layer["name"]:
                    raise ValueError("reorder_layer.expected_name no longer matches " + target)
                if operation.get("relative_expected_name") != relative_layer["name"]:
                    raise ValueError("reorder_layer.relative_expected_name no longer matches " + relative)
                if participates_in_clipping_chain(relative):
                    raise ValueError("reorder_layer cannot use a layer participating in a clipping-mask chain as its reference.")
                if layer["parent_key"] != relative_layer["parent_key"]:
                    raise ValueError("reorder_layer references must be siblings.")
                if operation.get("position") not in ("before", "after"):
                    raise ValueError("reorder_layer.position must be 'before' or 'after'.")
                if relative_layer["is_background"] and operation["position"] == "after":
                    raise ValueError("A layer cannot be reordered below the locked Background layer.")
                reordered.add(target)

            elif action == "set_visibility":
                target = operation.get("target")
                layer = require_layer(target, "set_visibility.target")
                if target in visibility_set:
                    raise ValueError("Visibility may be changed only once per layer: " + target)
                if operation.get("expected_name") != layer["name"]:
                    raise ValueError("set_visibility.expected_name no longer matches " + target)
                if layer["is_background"]:
                    raise ValueError("Background-layer visibility cannot be changed by organizer v1.")
                if not isinstance(operation.get("visible"), bool):
                    raise ValueError("set_visibility.visible must be a boolean.")
                visibility_set.add(target)

            else:
                raise ValueError(
                    "Unsupported organizer action at index " + str(index) + ": " + str(action)
                )

        normalized.append(operation)

    # Group-to-group moves are checked as one final parent graph, not only
    # one operation at a time, so two individually plausible moves cannot
    # create a cycle when combined.
    planned_group_parents = {
        key: layer["parent_key"]
        for key, layer in source_layers.items()
        if layer["typename"] == "LayerSet"
    }
    planned_group_parents.update(group_parents)
    for operation in normalized:
        if operation["action"] == "move_layer":
            target_key = operation["target"]
            if source_layers[target_key]["typename"] == "LayerSet":
                planned_group_parents[target_key] = operation["destination"]

    for group_key in planned_group_parents:
        visited: set[str] = set()
        current = group_key
        while current != "root":
            if current in visited:
                raise ValueError("Organizer plan would create a LayerSet parent cycle at " + current)
            visited.add(current)
            current = planned_group_parents.get(current)
            if current is None:
                raise ValueError("Organizer plan references an invalid LayerSet parent.")

    # Reordering cannot use a reference layer that will be moved/reordered,
    # and a target cannot be moved then reordered in separate plan sections.
    for operation in normalized:
        if operation["action"] == "reorder_layer":
            if operation["target"] in moved:
                raise ValueError("reorder_layer target cannot also be moved.")
            if operation["relative_to"] in moved or operation["relative_to"] in reordered:
                raise ValueError("reorder_layer relative_to cannot itself be moved/reordered.")

    verify_refs: Dict[str, Dict[str, str]] = {}

    def add_source_ref(key: str) -> None:
        verify_refs["source:" + key] = {"identity": "source:" + key, "source_key": key}

    def add_group_ref(key: str) -> None:
        verify_refs["group:" + key] = {"identity": "group:" + key, "group_key": key}

    for key in source_ref_keys:
        add_source_ref(key)
    for key in group_parents:
        add_group_ref(key)

    for operation in normalized:
        action = operation["action"]
        if action == "create_group":
            parent = operation["parent"]
            if parent in group_parents:
                add_group_ref(parent)
            elif parent != "root":
                add_source_ref(parent)
        elif action == "move_layer":
            destination = operation["destination"]
            if destination in group_parents:
                add_group_ref(destination)
            elif destination != "root":
                add_source_ref(destination)
        elif action == "reorder_layer":
            add_source_ref(operation["relative_to"])

    return {
        "operations": normalized,
        "verify_refs": list(verify_refs.values()),
        "created_groups": {
            operation["key"]: operation
            for operation in normalized
            if operation["action"] == "create_group"
        },
    }


def _build_apply_organize_plan_jsx(config: Dict[str, Any]) -> str:
    payload = json.dumps(config, ensure_ascii=True, separators=(",", ":"))
    template = r"""
(function () {
    function clean(value) {
        return String(value).replace(/[\r\n\t]/g, " ");
    }
    function encodeRecord(values) {
        var encoded = [];
        for (var i = 0; i < values.length; i++) {
            encoded.push(encodeURIComponent(
                values[i] === undefined || values[i] === null ? "" : String(values[i])
            ));
        }
        return encoded.join("|");
    }
    function walk(container, parentKey, map, count) {
        var collection = container.layers;
        for (var i = 0; i < collection.length; i++) {
            var layer = collection[i];
            var isGroup = layer.typename === "LayerSet";
            var key = parentKey + "/" + (isGroup ? "G" : "L") + i;
            map[key] = layer;
            count.value++;
            if (isGroup) {
                walk(layer, key, map, count);
            }
        }
    }
    function resolveContainer(key, doc, sourceLayers, newGroups) {
        if (key === "root") {
            return doc;
        }
        if (newGroups[key] !== undefined) {
            return newGroups[key];
        }
        if (sourceLayers[key] !== undefined &&
            sourceLayers[key].typename === "LayerSet") {
            return sourceLayers[key];
        }
        throw new Error("Destination group could not be resolved: " + key);
    }
    function moveToEnd(layer, container) {
        // Photoshop's locked Background layer must remain at the bottom.
        // Place a new final sibling immediately above it instead of trying
        // to move a layer below the Background.
        if (container.typename === "Document") {
            var layers = container.layers;
            if (layers.length > 0) {
                var bottom = layers[layers.length - 1];
                try {
                    if (bottom.isBackgroundLayer) {
                        layer.move(bottom, ElementPlacement.PLACEBEFORE);
                        return;
                    }
                } catch (backgroundError) {}
            }
        }
        layer.move(container, ElementPlacement.PLACEATEND);
    }
    function findNodeInfo(target, container, parentKey) {
        var collection = container.layers;
        for (var i = 0; i < collection.length; i++) {
            var layer = collection[i];
            var isGroup = layer.typename === "LayerSet";
            var key = parentKey + "/" + (isGroup ? "G" : "L") + i;
            if (layer === target) {
                return {
                    node_key: key,
                    parent_key: parentKey,
                    stack_index: i,
                    typename: layer.typename,
                    name: layer.name,
                    visible: layer.visible
                };
            }
            if (isGroup) {
                var nested = findNodeInfo(target, layer, key);
                if (nested !== null) {
                    return nested;
                }
            }
        }
        return null;
    }

    var cfg = __SCI_ORGANIZE_CONFIG__;
    var previousDialogs = app.displayDialogs;
    var previousRulerUnits = app.preferences.rulerUnits;
    var previousTypeUnits = app.preferences.typeUnits;
    var previousDoc = app.documents.length > 0 ? app.activeDocument : null;
    var doc = null;
    var sourceLayers = {};
    var newGroups = {};
    var count = {value: 0};
    var response = "";

    try {
        app.displayDialogs = DialogModes.NO;
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.POINTS;

        var inputFile = new File(cfg.temp_input_path);
        if (!inputFile.exists) {
            throw new Error("Temporary source PSD copy is missing.");
        }
        doc = app.open(inputFile);
        walk(doc, "root", sourceLayers, count);

        // Create all requested groups before resolving moves.
        for (var i = 0; i < cfg.operations.length; i++) {
            var create = cfg.operations[i];
            if (create.action !== "create_group") {
                continue;
            }
            var parent = resolveContainer(create.parent, doc, sourceLayers, newGroups);
            var group = parent.layerSets.add();
            group.name = create.name;
            moveToEnd(group, parent);
            newGroups[create.key] = group;
        }

        // Rename before moving so labels and paths remain recognizable in reports.
        for (var r = 0; r < cfg.operations.length; r++) {
            var rename = cfg.operations[r];
            if (rename.action === "rename_layer") {
                sourceLayers[rename.target].name = rename.new_name;
            }
        }

        for (var m = 0; m < cfg.operations.length; m++) {
            var move = cfg.operations[m];
            if (move.action !== "move_layer") {
                continue;
            }
            var target = sourceLayers[move.target];
            var destination = resolveContainer(
                move.destination, doc, sourceLayers, newGroups
            );
            if (move.position === "start") {
                target.move(destination, ElementPlacement.PLACEATBEGINNING);
            } else {
                moveToEnd(target, destination);
            }
        }

        for (var o = 0; o < cfg.operations.length; o++) {
            var reorder = cfg.operations[o];
            if (reorder.action !== "reorder_layer") {
                continue;
            }
            sourceLayers[reorder.target].move(
                sourceLayers[reorder.relative_to],
                reorder.position === "before" ?
                    ElementPlacement.PLACEBEFORE :
                    ElementPlacement.PLACEAFTER
            );
        }

        for (var v = 0; v < cfg.operations.length; v++) {
            var visibility = cfg.operations[v];
            if (visibility.action === "set_visibility") {
                sourceLayers[visibility.target].visible = visibility.visible;
            }
        }

        var outputFile = new File(cfg.staged_output_path);
        if (outputFile.exists) {
            throw new Error("Staged organizer output already exists.");
        }
        var saveOptions = new PhotoshopSaveOptions();
        saveOptions.embedColorProfile = true;
        saveOptions.maximizeCompatibility = true;
        doc.saveAs(outputFile, saveOptions, true, Extension.LOWERCASE);

        if (!outputFile.exists || outputFile.length <= 0) {
            throw new Error("Photoshop did not create the staged organizer PSD.");
        }

        var verificationRows = [];
        for (var q = 0; q < cfg.verify_refs.length; q++) {
            var ref = cfg.verify_refs[q];
            var targetRef = ref.source_key !== undefined ?
                sourceLayers[ref.source_key] : newGroups[ref.group_key];
            if (targetRef === undefined) {
                throw new Error("Could not resolve verification reference: " + ref.identity);
            }
            var info = findNodeInfo(targetRef, doc, "root");
            if (info === null) {
                throw new Error("Could not locate a modified layer after organization: " + ref.identity);
            }
            verificationRows.push(encodeRecord([
                "VERIFY",
                ref.identity,
                info.node_key,
                info.parent_key,
                info.stack_index,
                info.typename,
                info.name,
                info.visible
            ]));
        }

        response = [
            "SCI_PS_ORGANIZE_PSD",
            "OK",
            outputFile.fsName,
            String(cfg.operations.length),
            String(count.value),
            verificationRows.join("\n")
        ].join("\t");

        doc.close(SaveOptions.DONOTSAVECHANGES);
        doc = null;
    } catch (e) {
        response = [
            "SCI_PS_ORGANIZE_PSD",
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
            } catch (restoreDocError) {}
        }
        try {
            app.preferences.rulerUnits = previousRulerUnits;
        } catch (restoreRulerError) {}
        try {
            app.preferences.typeUnits = previousTypeUnits;
        } catch (restoreTypeError) {}
        app.displayDialogs = previousDialogs;
    }
    return response;
})();
"""
    return template.replace("__SCI_ORGANIZE_CONFIG__", payload)


def _parse_organize_verify_records(encoded_records: str) -> Dict[str, Dict[str, Any]]:
    verified: Dict[str, Dict[str, Any]] = {}
    for raw_line in encoded_records.split("\n"):
        if not raw_line:
            continue
        values = [unquote(value) for value in raw_line.split("|")]
        if len(values) != 8 or values[0] != "VERIFY":
            raise ValueError("Invalid organizer verification record: " + repr(values))
        identity = values[1]
        if identity in verified:
            raise ValueError("Duplicate organizer verification identity: " + identity)
        verified[identity] = {
            "node_key": values[2],
            "parent_key": values[3],
            "stack_index": int(values[4]),
            "typename": values[5],
            "name": values[6],
            "visible": values[7].lower() == "true",
        }
    return verified


def apply_organize_plan(
    input_path: str | Path,
    plan: Dict[str, Any] | str | Path,
    output_path: str | Path,
    *,
    approved: bool = False,
    overwrite: bool = False,
    bridge: Optional[PhotoshopBridge] = None,
) -> Dict[str, Any]:
    """Apply an explicitly approved, structure-only plan to a temporary copy of a PSD."""
    if not approved:
        raise PermissionError(
            "Organizer plan has not been approved. Review it with the user and pass approved=True."
        )

    source_file = Path(input_path).expanduser().resolve(strict=True)
    if not source_file.is_file() or source_file.suffix.lower() != ".psd":
        raise ValueError("input_path must be an existing PSD file.")
    output_file = Path(output_path).expanduser().resolve()
    if output_file.suffix.lower() != ".psd":
        raise ValueError("output_path must end in .psd.")
    if str(source_file).casefold() == str(output_file).casefold():
        raise ValueError("output_path must differ from input_path; the source PSD is never modified.")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    if output_file.exists() and not overwrite:
        raise FileExistsError(
            "Output PSD already exists; pass overwrite=True to replace it: "
            + str(output_file)
        )

    loaded_plan = _load_organize_plan(plan)
    source_hash_before = _sha256_file(source_file)
    active_bridge = bridge if bridge is not None else PhotoshopBridge()

    # Inspect immediately before application to reject stale or mismatched plans.
    before = inspect_psd(source_file, bridge=active_bridge)
    normalized = _validate_organize_plan(loaded_plan, before)

    token = uuid.uuid4().hex[:16]
    temp_input = Path(os.environ.get("TEMP", os.environ.get("TMP", "."))) / (
        "sci_org_" + token + ".psd"
    )
    staged_output = output_file.parent / (".sci_org_" + token + ".psd")
    try:
        shutil.copy2(source_file, temp_input)
        if _sha256_file(temp_input) != source_hash_before:
            raise PhotoshopBridgeError(
                "apply_organize_plan",
                "Temporary working copy is not byte-identical to the source PSD.",
                exception_type="WorkingCopyIntegrityError",
            )

        config = {
            "temp_input_path": str(temp_input),
            "staged_output_path": str(staged_output),
            "operations": normalized["operations"],
            "verify_refs": normalized["verify_refs"],
        }
        jsx = _build_apply_organize_plan_jsx(config)
        raw = active_bridge.execute_jsx(jsx)
        protocol = _parse_protocol(raw, "SCI_PS_ORGANIZE_PSD", 6, "apply_organize_plan", jsx)

        if not staged_output.is_file() or staged_output.stat().st_size <= 0:
            raise PhotoshopBridgeError(
                "apply_organize_plan",
                "Photoshop reported success but the staged PSD is missing or empty.",
                exception_type="PsdSaveVerificationError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        verified_refs = _parse_organize_verify_records(protocol[5])
        expected_identities = {
            item["identity"] for item in normalized["verify_refs"]
        }
        if set(verified_refs) != expected_identities:
            raise PhotoshopBridgeError(
                "apply_organize_plan",
                "Photoshop verification references do not match the approved plan.",
                exception_type="OrganizerVerificationMismatch",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        after = inspect_psd(staged_output, bridge=active_bridge)
        expected_layer_count = before["layer_count"] + len(normalized["created_groups"])
        expected_group_count = before["group_count"] + len(normalized["created_groups"])
        if (
            after["layer_count"] != expected_layer_count
            or after["group_count"] != expected_group_count
            or after["text_layer_count"] != before["text_layer_count"]
            or after["smart_object_count"] != before["smart_object_count"]
        ):
            raise PhotoshopBridgeError(
                "apply_organize_plan",
                "The reopened PSD changed layer/content-category counts beyond the approved group creation: "
                + repr({
                    "layers_before": before["layer_count"],
                    "layers_after": after["layer_count"],
                    "layers_expected": expected_layer_count,
                    "groups_before": before["group_count"],
                    "groups_after": after["group_count"],
                    "groups_expected": expected_group_count,
                    "text_before": before["text_layer_count"],
                    "text_after": after["text_layer_count"],
                    "smart_objects_before": before["smart_object_count"],
                    "smart_objects_after": after["smart_object_count"],
                }),
                exception_type="OrganizerContentCountMismatch",
            )
        after_by_key = {layer["node_key"]: layer for layer in after["layers"]}

        for identity, observed in verified_refs.items():
            reopened = after_by_key.get(observed["node_key"])
            if reopened is None:
                raise PhotoshopBridgeError(
                    "apply_organize_plan",
                    "Reopened PSD is missing an expected organized layer: " + identity,
                    exception_type="OrganizerReopenVerificationError",
                )
            if (
                reopened["parent_key"] != observed["parent_key"]
                or reopened["stack_index"] != observed["stack_index"]
                or reopened["typename"] != observed["typename"]
                or reopened["name"] != observed["name"]
                or reopened["visible"] != observed["visible"]
            ):
                raise PhotoshopBridgeError(
                    "apply_organize_plan",
                    "Reopened PSD layer state differs from Photoshop's organization result: "
                    + identity,
                    exception_type="OrganizerReopenVerificationError",
                )

        for operation in normalized["operations"]:
            action = operation["action"]
            if action == "create_group":
                observed = verified_refs["group:" + operation["key"]]
                if observed["typename"] != "LayerSet" or observed["name"] != operation["name"]:
                    raise PhotoshopBridgeError(
                        "apply_organize_plan",
                        "Created group failed verification: " + operation["key"],
                        exception_type="OrganizerOperationVerificationError",
                    )
                parent_identity = (
                    "group:" + operation["parent"]
                    if operation["parent"] in normalized["created_groups"]
                    else "source:" + operation["parent"]
                    if operation["parent"] != "root"
                    else None
                )
                expected_parent = (
                    verified_refs[parent_identity]["node_key"]
                    if parent_identity is not None else "root"
                )
                if observed["parent_key"] != expected_parent:
                    raise PhotoshopBridgeError(
                        "apply_organize_plan",
                        "Created group was placed under the wrong parent: " + operation["key"],
                        exception_type="OrganizerOperationVerificationError",
                    )

            elif action == "rename_layer":
                observed = verified_refs["source:" + operation["target"]]
                if observed["name"] != operation["new_name"]:
                    raise PhotoshopBridgeError(
                        "apply_organize_plan",
                        "Renamed layer failed verification: " + operation["target"],
                        exception_type="OrganizerOperationVerificationError",
                    )

            elif action == "move_layer":
                observed = verified_refs["source:" + operation["target"]]
                destination = operation["destination"]
                if destination == "root":
                    expected_parent = "root"
                elif destination in normalized["created_groups"]:
                    expected_parent = verified_refs["group:" + destination]["node_key"]
                else:
                    expected_parent = verified_refs["source:" + destination]["node_key"]
                if observed["parent_key"] != expected_parent:
                    raise PhotoshopBridgeError(
                        "apply_organize_plan",
                        "Moved layer has an unexpected parent: " + operation["target"],
                        exception_type="OrganizerOperationVerificationError",
                    )

            elif action == "reorder_layer":
                observed = verified_refs["source:" + operation["target"]]
                relative = verified_refs["source:" + operation["relative_to"]]
                correct = (
                    observed["parent_key"] == relative["parent_key"]
                    and (
                        observed["stack_index"] < relative["stack_index"]
                        if operation["position"] == "before"
                        else observed["stack_index"] > relative["stack_index"]
                    )
                )
                if not correct:
                    raise PhotoshopBridgeError(
                        "apply_organize_plan",
                        "Reordered layer is not on the requested side of its reference: "
                        + operation["target"],
                        exception_type="OrganizerOperationVerificationError",
                    )

            elif action == "set_visibility":
                observed = verified_refs["source:" + operation["target"]]
                if observed["visible"] is not operation["visible"]:
                    raise PhotoshopBridgeError(
                        "apply_organize_plan",
                        "Layer visibility failed verification: " + operation["target"],
                        exception_type="OrganizerOperationVerificationError",
                    )

        if _sha256_file(source_file) != source_hash_before:
            raise PhotoshopBridgeError(
                "apply_organize_plan",
                "The source PSD SHA-256 changed during organization.",
                exception_type="SourceIntegrityError",
                jsx_preview=_clean_jsx_preview(jsx),
            )

        staged_hash = _sha256_file(staged_output)
        if output_file.exists() and not overwrite:
            raise FileExistsError(
                "Output PSD appeared during processing; refusing to overwrite: "
                + str(output_file)
            )
        os.replace(staged_output, output_file)
        if _sha256_file(output_file) != staged_hash:
            raise PhotoshopBridgeError(
                "apply_organize_plan",
                "Final PSD differs from its verified staging file.",
                exception_type="OutputIntegrityError",
            )

        return {
            "passed": True,
            "input_psd": str(source_file),
            "output_psd": str(output_file),
            "source_sha256_unchanged": True,
            "source_layer_tree_sha256": before["layer_tree_sha256"],
            "output_layer_tree_sha256": after["layer_tree_sha256"],
            "operation_count": len(normalized["operations"]),
            "layer_count_before": before["layer_count"],
            "layer_count_after": after["layer_count"],
            "group_count_after": after["group_count"],
            "text_layer_count_after": after["text_layer_count"],
            "smart_object_count_after": after["smart_object_count"],
            "operations": normalized["operations"],
            "verified_references": verified_refs,
        }
    finally:
        for temporary in (temp_input, staged_output):
            try:
                if temporary.exists():
                    temporary.unlink()
            except OSError:
                pass



def _print_human_report(action: str, result: Dict[str, Any]) -> None:
    print("SCI FIGURE PHOTOSHOP BRIDGE")
    print("Action: " + action)

    if action == "diagnose":
        print("Python: " + str(result.get("python_executable")))
        print("Python version: " + str(result.get("python_version")))
        print("photoshop-python-api: " + str(result.get("photoshop_python_api_version")))
        print("comtypes: " + str(result.get("comtypes_version")))

        photoshop = result.get("photoshop") or {}
        print("COM connection: " + str(photoshop.get("connected", False)))
        print("Photoshop version: " + str(photoshop.get("photoshop_version")))
        print("Open documents: " + str(photoshop.get("open_document_count")))

        units = result.get("unit_lock") or {}
        print("Ruler units Pixels: " + str(units.get("ruler_units_pixels")))
        print("Type units Points: " + str(units.get("type_units_points")))

        exception_test = result.get("jsx_exception_capture") or {}
        print("JSX exception capture: " + str(exception_test.get("passed")))

    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))

    print("RESULT: " + ("PASS" if result.get("passed", True) else "FAILED"))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Windows Photoshop COM and ExtendScript bridge."
    )
    parser.add_argument(
        "--action",
        choices=("diagnose", "status", "lock-units", "exec-jsx", "compose-grid", "add-panel-labels", "add-scale-bar", "unify-fonts", "export-figure", "preflight-figure", "inspect-psd", "apply-organize-plan"),
        default="diagnose",
        help="Bridge operation. Default: diagnose.",
    )
    parser.add_argument(
        "--execution-mode",
        choices=("quiet", "interactive"),
        default="quiet",
        help=(
            "quiet suppresses Photoshop dialogs but still requires desktop Photoshop; "
            "interactive permits Photoshop dialogs. Default: quiet."
        ),
    )
    parser.add_argument(
        "--jsx-file",
        type=Path,
        help="JSX file to execute when --action exec-jsx is selected.",
    )
    parser.add_argument(
        "--image",
        action="append",
        help="Input raster image; repeat once per panel in row-major order.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Output PSD path. Existing files are protected unless --overwrite is supplied.",
    )
    parser.add_argument(
        "--preset",
        choices=("single", "half_double", "double", "custom"),
        default="double",
        help="Canvas width preset. Default: double (17.6 cm).",
    )
    parser.add_argument("--width-cm", type=float, help="Custom canvas width in cm.")
    parser.add_argument("--height-cm", type=float, help="Optional explicit canvas height in cm.")
    parser.add_argument("--rows", type=int, default=2)
    parser.add_argument("--cols", type=int, default=2)
    parser.add_argument("--dpi", type=int, choices=(300, 600), default=300)
    parser.add_argument("--margin-cm", type=float, default=0.3)
    parser.add_argument("--gap-cm", type=float, default=0.2)
    parser.add_argument("--align", choices=("center", "top-left"), default="center")
    parser.add_argument("--auto-trim", action="store_true")
    parser.add_argument(
        "--input-psd",
        type=Path,
        help="Input PSD for --action add-panel-labels; source is never modified.",
    )
    parser.add_argument(
        "--inspection-report",
        type=Path,
        help="Optional JSON output path for --action inspect-psd.",
    )
    parser.add_argument(
        "--organize-plan",
        type=Path,
        help="Approved structure-only JSON plan for --action apply-organize-plan.",
    )
    parser.add_argument(
        "--confirm-plan",
        action="store_true",
        help="Explicitly approve execution of the organizer plan on a temporary PSD copy.",
    )
    parser.add_argument(
        "--label-style",
        choices=PANEL_LABEL_STYLES,
        default="lowercase",
        help="Panel label style for --action add-panel-labels.",
    )
    parser.add_argument("--font-postscript", default="Arial-BoldMT")
    parser.add_argument("--font-size-pt", type=float, default=10.0)
    parser.add_argument("--label-inset-cm", type=float, default=0.1)
    parser.add_argument("--label-color", default="#000000")
    parser.add_argument(
        "--panel-id",
        default="a",
        help="Panel group identifier for --action add-scale-bar.",
    )
    parser.add_argument(
        "--scale-mode",
        choices=SCALE_BAR_MODES,
        default="pixels",
        help="Scale-bar width mode: pixels or calibrated physical length.",
    )
    parser.add_argument("--bar-width-px", type=float)
    parser.add_argument("--scale-length", type=float)
    parser.add_argument("--pixel-size", type=float)
    parser.add_argument("--scale-unit")
    parser.add_argument("--scale-label")
    parser.add_argument("--bar-color", default="#000000")
    parser.add_argument("--bar-outline")
    parser.add_argument("--bar-height-pt", type=float, default=2.0)
    parser.add_argument("--outline-width-pt", type=float, default=0.5)
    parser.add_argument("--scale-label-size-pt", type=float, default=8.0)
    parser.add_argument("--scale-label-gap-pt", type=float, default=1.5)
    parser.add_argument("--scale-inset-cm", type=float, default=0.2)
    parser.add_argument("--unify-font-postscript")
    parser.add_argument("--unify-font-size-pt", type=float)
    parser.add_argument("--unify-font-color", help="Hex text color, for example #000000.")
    parser.add_argument("--export-dir", type=Path)
    parser.add_argument("--export-basename")
    parser.add_argument("--export-dpi", type=int, choices=(300, 600))
    parser.add_argument("--preflight-dpi", type=int, choices=(300, 600))
    parser.add_argument("--preflight-width-cm", type=float)
    parser.add_argument("--preflight-width-tolerance-cm", type=float, default=0.05)
    parser.add_argument("--preflight-min-panel-ppi", type=float, default=300.0)
    parser.add_argument("--preflight-panel-label-font", default="Arial-BoldMT")
    parser.add_argument("--preflight-panel-label-min-pt", type=float, default=10.0)
    parser.add_argument("--preflight-panel-label-max-pt", type=float, default=12.0)
    parser.add_argument("--preflight-figure-text-font", default="ArialMT")
    parser.add_argument("--preflight-figure-text-size-pt", type=float, default=7.0)
    parser.add_argument("--preflight-text-size-tolerance-pt", type=float, default=0.5)
    parser.add_argument("--preflight-skip-text-style-checks", action="store_true")
    parser.add_argument("--preflight-no-panel-label-required", action="store_true")
    parser.add_argument("--preflight-calibration-json", type=Path)
    parser.add_argument("--preflight-calibration-tolerance-px", type=float, default=2.0)
    parser.add_argument("--preflight-export-dir", type=Path)
    parser.add_argument("--preflight-export-basename")
    parser.add_argument("--preflight-strict", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print machine-readable JSON.",
    )
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if args.action == "exec-jsx":
        if args.jsx_file is None:
            print("ERROR: --jsx-file is required for --action exec-jsx.", file=sys.stderr)
            return 2
        if not args.jsx_file.is_file():
            print("ERROR: JSX file does not exist: " + str(args.jsx_file), file=sys.stderr)
            return 2
    if args.action == "add-panel-labels":
        if args.input_psd is None or args.output is None:
            print(
                "ERROR: --input-psd and --output are required for "
                "--action add-panel-labels.",
                file=sys.stderr,
            )
            return 2
    if args.action == "add-scale-bar":
        if args.input_psd is None or args.output is None:
            print(
                "ERROR: --input-psd and --output are required for "
                "--action add-scale-bar.",
                file=sys.stderr,
            )
            return 2
        if args.scale_mode == "pixels" and args.bar_width_px is None:
            print("ERROR: --bar-width-px is required in pixel mode.", file=sys.stderr)
            return 2
        if args.scale_mode == "physical" and (
            args.scale_length is None or args.pixel_size is None or not args.scale_unit
        ):
            print(
                "ERROR: --scale-length, --pixel-size, and --scale-unit "
                "are required in physical mode.",
                file=sys.stderr,
            )
            return 2

    if args.action == "unify-fonts":
        if args.input_psd is None or args.output is None:
            print(
                "ERROR: --input-psd and --output are required for "
                "--action unify-fonts.",
                file=sys.stderr,
            )
            return 2
        if not any((
            args.unify_font_postscript is not None,
            args.unify_font_size_pt is not None,
            args.unify_font_color is not None,
        )):
            print(
                "ERROR: specify at least one of --unify-font-postscript, "
                "--unify-font-size-pt, or --unify-font-color.",
                file=sys.stderr,
            )
            return 2

    if args.action == "inspect-psd":
        if args.input_psd is None:
            print(
                "ERROR: --input-psd is required for --action inspect-psd.",
                file=sys.stderr,
            )
            return 2

    if args.action == "apply-organize-plan":
        if args.input_psd is None or args.organize_plan is None or args.output is None:
            print(
                "ERROR: --input-psd, --organize-plan, and --output are required "
                "for --action apply-organize-plan.",
                file=sys.stderr,
            )
            return 2
        if not args.organize_plan.is_file():
            print("ERROR: organizer plan does not exist: " + str(args.organize_plan), file=sys.stderr)
            return 2
        if not args.confirm_plan:
            print(
                "ERROR: review the organizer plan and pass --confirm-plan to apply it.",
                file=sys.stderr,
            )
            return 2

    if args.action == "preflight-figure":
        if args.input_psd is None:
            print(
                "ERROR: --input-psd is required for --action preflight-figure.",
                file=sys.stderr,
            )
            return 2
        if args.preflight_calibration_json is not None and not args.preflight_calibration_json.is_file():
            print(
                "ERROR: calibration manifest does not exist: " +
                str(args.preflight_calibration_json),
                file=sys.stderr,
            )
            return 2

    if args.action == "export-figure":
        if args.input_psd is None:
            print(
                "ERROR: --input-psd is required for --action export-figure.",
                file=sys.stderr,
            )
            return 2

    try:
        bridge = PhotoshopBridge(execution_mode=args.execution_mode)

        if args.action == "diagnose":
            result = bridge.diagnose()
        elif args.action == "status":
            result = bridge.get_status()
            result["passed"] = True
        elif args.action == "lock-units":
            result = bridge.lock_units()
            result["passed"] = True
        elif args.action == "add-panel-labels":
            result = add_panel_labels(
                args.input_psd,
                args.output,
                style=args.label_style,
                size_pt=args.font_size_pt,
                font_postscript=args.font_postscript,
                inset_cm=args.label_inset_cm,
                color_hex=args.label_color,
                overwrite=args.overwrite,
                bridge=bridge,
            )
        elif args.action == "add-scale-bar":
            result = add_scale_bar(
                args.input_psd,
                args.output,
                panel_id=args.panel_id,
                mode=args.scale_mode,
                bar_width_px=args.bar_width_px,
                physical_length=args.scale_length,
                pixel_size=args.pixel_size,
                unit=args.scale_unit,
                label=args.scale_label,
                color_hex=args.bar_color,
                outline_hex=args.bar_outline,
                bar_height_pt=args.bar_height_pt,
                outline_width_pt=args.outline_width_pt,
                label_size_pt=args.scale_label_size_pt,
                label_gap_pt=args.scale_label_gap_pt,
                inset_cm=args.scale_inset_cm,
                overwrite=args.overwrite,
                bridge=bridge,
            )
        elif args.action == "unify-fonts":
            result = unify_fonts(
                args.input_psd,
                args.output,
                font_postscript=args.unify_font_postscript,
                size_pt=args.unify_font_size_pt,
                color_hex=args.unify_font_color,
                overwrite=args.overwrite,
                bridge=bridge,
            )
        elif args.action == "inspect-psd":
            result = inspect_psd(
                args.input_psd,
                report_path=args.inspection_report,
                bridge=bridge,
            )
        elif args.action == "apply-organize-plan":
            result = apply_organize_plan(
                args.input_psd,
                args.organize_plan,
                args.output,
                approved=args.confirm_plan,
                overwrite=args.overwrite,
                bridge=bridge,
            )
        elif args.action == "preflight-figure":
            result = preflight_figure(
                args.input_psd,
                expected_dpi=args.preflight_dpi,
                expected_width_cm=args.preflight_width_cm,
                width_tolerance_cm=args.preflight_width_tolerance_cm,
                min_panel_ppi=args.preflight_min_panel_ppi,
                require_panel_labels=not args.preflight_no_panel_label_required,
                panel_label_font=args.preflight_panel_label_font,
                panel_label_min_pt=args.preflight_panel_label_min_pt,
                panel_label_max_pt=args.preflight_panel_label_max_pt,
                figure_text_font=args.preflight_figure_text_font,
                figure_text_size_pt=args.preflight_figure_text_size_pt,
                text_size_tolerance_pt=args.preflight_text_size_tolerance_pt,
                check_text_styles=not args.preflight_skip_text_style_checks,
                calibrations=args.preflight_calibration_json,
                calibration_default_tolerance_px=args.preflight_calibration_tolerance_px,
                export_dir=args.preflight_export_dir,
                export_basename=args.preflight_export_basename,
                strict=args.preflight_strict,
                bridge=bridge,
            )
        elif args.action == "export-figure":
            result = export_figure(
                args.input_psd,
                args.export_dir,
                basename=args.export_basename,
                expected_dpi=args.export_dpi,
                overwrite=args.overwrite,
                bridge=bridge,
            )
        elif args.action == "compose-grid":
            result = compose_grid(
                args.image or [],
                args.output,
                preset=args.preset,
                width_cm=args.width_cm,
                height_cm=args.height_cm,
                rows=args.rows,
                cols=args.cols,
                dpi=args.dpi,
                margin_cm=args.margin_cm,
                gap_cm=args.gap_cm,
                align=args.align,
                auto_trim=args.auto_trim,
                overwrite=args.overwrite,
                bridge=bridge,
            )
        else:
            jsx = args.jsx_file.read_text(encoding="utf-8-sig")
            result = {
                "jsx_file": str(args.jsx_file.resolve()),
                "result": bridge.execute_jsx(jsx),
                "passed": True,
            }

        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            _print_human_report(args.action, result)

        return 0 if result.get("passed", True) else 1

    except PhotoshopBridgeError as exc:
        error = exc.to_dict()
        if args.json:
            print(json.dumps({"passed": False, "error": error}, ensure_ascii=False, indent=2))
        else:
            print("BRIDGE ERROR")
            print("Operation: " + str(error["operation"]))
            print("Exception: " + str(error["exception_type"]))
            print("Message: " + str(error["message"]))
            if error.get("jsx_preview"):
                print("JSX preview: " + str(error["jsx_preview"]))
        return 1
    except Exception as exc:
        if args.json:
            print(json.dumps({
                "passed": False,
                "error": {
                    "operation": args.action,
                    "exception_type": type(exc).__name__,
                    "message": str(exc),
                },
            }, ensure_ascii=False, indent=2))
        else:
            print("UNEXPECTED ERROR: " + type(exc).__name__ + ": " + str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

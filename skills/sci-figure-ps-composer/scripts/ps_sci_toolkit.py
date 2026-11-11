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

    def __init__(self, app: Any = None) -> None:
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

        try:
            result = self.app.doJavaScript(jsx)
        except Exception as exc:
            raise PhotoshopBridgeError(
                "execute_jsx",
                "Photoshop rejected or failed to execute JSX. Original error: " + str(exc),
                exception_type=type(exc).__name__,
                jsx_preview=_clean_jsx_preview(jsx),
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
        choices=("diagnose", "status", "lock-units", "exec-jsx", "compose-grid", "add-panel-labels", "add-scale-bar", "unify-fonts", "export-figure"),
        default="diagnose",
        help="Bridge operation. Default: diagnose.",
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

    if args.action == "export-figure":
        if args.input_psd is None:
            print(
                "ERROR: --input-psd is required for --action export-figure.",
                file=sys.stderr,
            )
            return 2

    try:
        bridge = PhotoshopBridge()

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

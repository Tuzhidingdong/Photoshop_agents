#!/usr/bin/env python
"""Photoshop COM/JSX bridge for sci-figure-ps-composer.

This module deliberately keeps Photoshop as the native document engine.
It provides connection, JSX execution, preference locking, and diagnostics.
"""

from __future__ import annotations

import argparse
import importlib.metadata
import json
import os
import platform
import sys
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
    function trimEmbeddedObject(parentDoc, smartLayer) {
        var smartDoc = null;
        var changed = false;
        try {
            parentDoc.activeLayer = smartLayer;
            executeAction(
                stringIDToTypeID("placedLayerEditContents"),
                new ActionDescriptor(),
                DialogModes.NO
            );
            smartDoc = app.activeDocument;

            var oldWidth = px(smartDoc.width);
            var oldHeight = px(smartDoc.height);
            var whiteTopLeft = false;

            // Only invoke color-based trim when the top-left pixel is near-white.
            // This avoids trimming a dark feature when the source has no white
            // corner. TOPLEFT uses the exact corner pixel as its trim color.
            try {
                var sampler = smartDoc.colorSamplers.add([
                    new UnitValue(0, "px"),
                    new UnitValue(0, "px")
                ]);
                var rgb = sampler.color.rgb;
                whiteTopLeft = (
                    rgb.red >= 245 &&
                    rgb.green >= 245 &&
                    rgb.blue >= 245
                );
                sampler.remove();
            } catch (sampleError) {
                whiteTopLeft = false;
            }

            // Trim using exactly one mode. Photoshop can reject the
            // TRANSPARENT trim command when an opaque RGB/white-background
            // image has no transparent border. The top-left color trim is
            // sufficient for opaque white-margin exports; alpha-based
            // sources use the transparent-pixel trim instead.
            if (whiteTopLeft) {
                smartDoc.trim(TrimType.TOPLEFT, true, true, true, true);
            } else {
                smartDoc.trim(TrimType.TRANSPARENT, true, true, true, true);
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

            if (cfg.auto_trim) {
                var trimResult = trimEmbeddedObject(doc, imageLayer);
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
        choices=("diagnose", "status", "lock-units", "exec-jsx", "compose-grid"),
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

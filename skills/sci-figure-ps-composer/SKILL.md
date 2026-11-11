---
name: sci-figure-ps-composer
description: Automate publication-ready scientific multi-panel figure composition in desktop Adobe Photoshop on Windows, using Python COM automation and native ExtendScript/JSX while preserving editable PSD layers.
---

# sci-figure-ps-composer

## Purpose

Create and maintain publication-ready scientific multi-panel figures in the Windows desktop edition of Adobe Photoshop. Python is the orchestration and geometry layer; Photoshop remains the native document, typography, layer, smart-object, and export engine.

The implementation is staged. The Photoshop Bridge, `compose_grid`, editable panel labels, and calibrated scale bars have passed real-machine verification on the user's Photoshop 2026 environment. Recursive text-layer font normalization has passed the nested-group Photoshop save/reopen test. The Stage 6 PSD/TIFF/PNG publication export is implemented and covered by unit tests; its real Photoshop export and metadata verification is the current integration check. Publication preflight remains on the roadmap.

## Runtime requirements

- Windows desktop session.
- Adobe Photoshop installed. The Photoshop process should be started before running the bridge.
- Python 3.10 or later; the validated user environment is Python 3.13.9.
- photoshop-python-api 0.24.2 or a compatible 0.24.x release.
- COM support through the package's comtypes dependency.
- Pillow 10 or later, used only to read source raster dimensions; it does not re-encode or modify original images.

Install the dependency into the exact Python interpreter that will run the skill:

~~~powershell
python -m pip install -r requirements.txt
~~~

Do not assume the python command resolves to the intended environment. Prefer the interpreter explicitly selected by the user or the active environment's python.exe.

## Current implemented interface

The toolkit is located at scripts/ps_sci_toolkit.py.

~~~powershell
python scripts/ps_sci_toolkit.py --action diagnose
python scripts/ps_sci_toolkit.py --action status
python scripts/ps_sci_toolkit.py --action lock-units
python scripts/ps_sci_toolkit.py --action exec-jsx --jsx-file path\to\script.jsx

# Four input panels, row-major order; 17.6 cm double-column, 300 DPI.
python scripts/ps_sci_toolkit.py --action compose-grid `
  --image panel-a.tif --image panel-b.tif `
  --image panel-c.tif --image panel-d.tif `
  --rows 2 --cols 2 --preset double --dpi 300 `
  --margin-cm 0.3 --gap-cm 0.2 --auto-trim `
  --output Figure_1.psd

# Add editable labels to a copy of the composed PSD.
python scripts/ps_sci_toolkit.py --action add-panel-labels `
  --input-psd Figure_1.psd --output Figure_1_labeled.psd `
  --label-style lowercase --font-postscript Arial-BoldMT `
  --font-size-pt 10 --label-inset-cm 0.1

# Add a 120-document-pixel scale bar to panel a.
python scripts/ps_sci_toolkit.py --action add-scale-bar `
  --input-psd Figure_1_labeled.psd --output Figure_1_scalebar.psd `
  --panel-id a --scale-mode pixels --bar-width-px 120 `
  --bar-color "#000000" --bar-outline "#FFFFFF"

# Add a physically calibrated 10 µm scale bar to panel a.
# pixel-size means µm per source-image pixel.
python scripts/ps_sci_toolkit.py --action add-scale-bar `
  --input-psd Figure_1_labeled.psd --output Figure_1_scalebar.psd `
  --panel-id a --scale-mode physical --scale-length 10 `
  --pixel-size 0.2 --scale-unit "µm" `
  --bar-color "#000000" --bar-outline "#FFFFFF"

# Recursively normalize every editable text layer in the PSD hierarchy.
python scripts/ps_sci_toolkit.py --action unify-fonts `
  --input-psd Figure_1_scalebar.psd --output Figure_1_fonts.psd `
  --unify-font-postscript ArialMT --unify-font-size-pt 7 `
  --unify-font-color "#000000"

# Export an editable PSD copy, flattened RGB LZW TIFF, and RGB PNG.
python scripts/ps_sci_toolkit.py --action export-figure `
  --input-psd Figure_1_fonts.psd --export-dir Figure_1_exports `
  --export-dpi 300
~~~

The default action is diagnose.

Run the real Photoshop integration demo on Windows with Photoshop open:

~~~powershell
python tests/test_compose_grid_demo.py
~~~

The demo generates four synthetic source panels, composes them into a 2×2 PSD, and reopens the saved PSD in Photoshop to check the 2079 px canvas width, 300 DPI, RGB mode, four panel groups, and four embedded Smart Objects. It also checks that the original synthetic source files were not modified. The demo writes its assets and PSD under Documents in a timestamped folder.

After Stage 2 passes, run the Stage 3 text-layer integration test against that PSD:

~~~powershell
python tests/test_add_panel_labels_demo.py --input-psd path\to\Figure_Grid_Stage2.psd
~~~

The Stage 3 test saves a separate labeled PSD, reopens it in Photoshop, and verifies one editable Arial Bold text layer per panel group, the selected label style, 10 pt size by default, black text, and non-empty output. It never saves changes back to the input PSD.

After Stage 3 passes, run the Stage 4 scale-bar test against the labeled PSD:

~~~powershell
python tests/test_add_scale_bar_demo.py --input-psd path\to\Figure_Grid_Stage2_Stage3_Labeled.psd
~~~

The Stage 4 integration test adds a calibrated 10 µm bar to panel a, with a separate bar layer, contrasting outline layer, and editable text layer. It reopens the output and verifies the bar width against the physical calibration, the bottom-right placement, and preservation of the existing panel label and all other groups. It also verifies that the source PSD hash is unchanged.

After Stage 4 passes, run the Stage 5 recursive font-normalization integration test:

~~~powershell
python tests/test_unify_fonts_demo.py --input-psd path\\to\\Figure_Grid_Stage4_ScaleBar.psd
python tests/test_export_figure_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd
~~~

The Stage 5 demo creates a separate fixture PSD with an additional text layer nested two groups deep, then normalizes the font, point size, and color of every editable text layer into a new PSD. It reopens that output and verifies every text layer, including the nested test layer, as well as preservation of the four panel labels and scale-bar label. Both the supplied input PSD and the generated fixture PSD are checked for unchanged hashes.

After Stage 5 passes, run the Stage 6 publication export integration test:

~~~powershell
python tests/test_export_figure_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd
~~~

The Stage 6 test creates a new export folder containing a byte-identical editable PSD copy, a flattened RGB TIFF using LZW compression, and a flattened RGB PNG. Photoshop reopens the PSD/TIFF/PNG and verifies their dimensions, RGB mode, resolution, panel groups, Smart Objects, and text layers. Pillow reads—but does not rewrite—the raster exports to verify RGB mode, dimensions, DPI metadata, and TIFF compression tag 259 = 5 (LZW). The source PSD hash must remain unchanged.

### Bridge behavior

1. Import photoshop.api and create photoshop.api.Application().
2. Execute JSX through the package's Application.doJavaScript() method.
3. Normalize Python/COM/JSX failures as PhotoshopBridgeError, preserving the operation stage, exception type, message, and a short JSX preview.
4. Lock app.preferences.rulerUnits to Units.PIXELS.
5. Lock app.preferences.typeUnits to TypeUnits.POINTS.
6. Verify both preferences by reading them back from Photoshop.

The unit lock intentionally changes Photoshop's global preferences and leaves them set to Pixels and Points. It does not create, close, or modify a document.

## Agent workflow

### Before any document operation

- Confirm the user is on Windows and identify the exact Python interpreter.
- Confirm Photoshop is running. Do not launch or terminate Photoshop automatically.
- Run --action diagnose and inspect its structured result.
- If COM connection fails, stop before attempting document changes. Report the stage and original error.
- If JSX execution works but unit verification fails, stop and report the observed values; do not silently proceed with geometry calculations.
- Never claim real Photoshop integration passed based only on mocked tests.

### JSX authoring requirements

- Use Photoshop-native DOM/ExtendScript operations for document, layer, text, smart-object, and export work.
- Keep JSX payloads self-contained and return a machine-readable result. The current bridge uses tab-delimited protocol markers to avoid dependence on newer JavaScript features in ExtendScript.
- Catch errors inside JSX when a controlled result is required; include the error name, message, and line when available.
- Allow Python to catch COM exceptions from doJavaScript() and wrap them in PhotoshopBridgeError.
- Do not use browser APIs, Node APIs, or modern JavaScript syntax unsupported by ExtendScript.
- Escape all user-provided strings before embedding them in JSX. Never concatenate untrusted file paths or labels without encoding.

### Document safety

- Do not close, overwrite, flatten, or save a user's active document unless the requested task explicitly requires it.
- Inspect app.documents.length before document-dependent operations.
- Use explicit output paths and avoid overwriting existing files by default.
- Keep original input images unchanged. Preprocessing must write to a separate working location.
- Preserve editable layers in PSD output. Flatten only export copies when producing TIFF or PNG.
- Report what was changed and where files were written.

## compose_grid: current implementation

The Python API is:

~~~python
from ps_sci_toolkit import compose_grid

result = compose_grid(
    ["panel-a.tif", "panel-b.tif", "panel-c.tif", "panel-d.tif"],
    "Figure_1.psd",
    preset="double",
    rows=2,
    cols=2,
    dpi=300,
    margin_cm=0.3,
    gap_cm=0.2,
    align="center",
    auto_trim=True,
)
~~~

Supported width presets are `single` (8.5 cm), `half_double` (12.0 cm), and `double` (17.6 cm), plus `custom` with an explicit `width_cm`. DPI is 300 or 600. If `height_cm` is omitted, the toolkit derives the canvas height from the median source-image aspect ratio. Image count must equal `rows × cols`; inputs are assigned in row-major order.

The script calculates layout in pixels, creates a white RGB Photoshop document, places every source through Photoshop's native Place command as an embedded Smart Object, and moves each object into its own `Group_Panel_a`, `Group_Panel_b`, etc. Images are scaled proportionally to fit their cells; they are never stretched. `align` accepts `center` or `top-left`.

When `auto_trim=True`, Pillow inspects each source file without modifying or re-encoding it, then supplies only the required trim steps to Photoshop. If the alpha bounding box is smaller than the source canvas, Photoshop trims transparent pixels first with `Document.trim(TrimType.TRANSPARENT)`. The remaining image is then checked for a same-color border matching its top-left pixel; if such a border is detected, Photoshop uses `Document.trim(TrimType.TOPLEFT)`. Steps that have no detected border are not invoked, avoiding invalid Photoshop trim commands on fully opaque or already-trimmed images. The color-based trim uses exact corner-color matching, not a tolerance-based near-white flood-fill. Inspect results when white scientific signal touches the edge.

The output path must end in `.psd`. Existing files are protected unless `overwrite=True` is explicitly requested. After Photoshop reports success, Python checks the calculated canvas dimensions and verifies that the PSD exists and is non-empty. This save check does not replace reopening the PSD in Photoshop and inspecting its Smart Objects.

## Remaining feature milestones

1. `add_panel_labels`: implemented; CI and real Photoshop Stage 3 integration verified.
2. `add_scale_bar`: implemented; CI and real Photoshop Stage 4 integration verified.
3. `unify_fonts`: implemented; unit tests and real Photoshop Stage 5 recursive integration verified.
4. `export_figure`: implemented; unit tests cover editable PSD copy, RGB LZW TIFF, RGB PNG, DPI, dimensions, and output protection; real Photoshop Stage 6 verification pending.
5. `preflight_figure`: canvas dimensions, effective panel PPI, layer structure, label/font checks, calibration checks, and export metadata checks.

## add_scale_bar: current implementation

The Python API is:

~~~python
from ps_sci_toolkit import add_scale_bar

# Explicit display-width mode: the bar is 120 px wide in the final PSD.
result = add_scale_bar(
    "Figure_1_labeled.psd",
    "Figure_1_scalebar.psd",
    panel_id="a",
    mode="pixels",
    bar_width_px=120,
    label="120 px",
    color_hex="#000000",
    outline_hex="#FFFFFF",
)

# Physical calibration: 10 µm at 0.2 µm per pixel in the source image.
result = add_scale_bar(
    "Figure_1_labeled.psd",
    "Figure_1_scalebar.psd",
    panel_id="a",
    mode="physical",
    physical_length=10,
    pixel_size=0.2,
    unit="µm",
    color_hex="#000000",
    outline_hex="#FFFFFF",
)
~~~

Pixel mode treats `bar_width_px` as pixels in the final Photoshop document. Physical mode treats `pixel_size` as physical units per pixel in the source Smart Object. Photoshop opens the selected embedded Smart Object, reads its current pixel dimensions after any Stage 2 trimming, then calculates:

~~~text
bar width in final document pixels
  = (physical_length / pixel_size)
    × (displayed Smart Object width / Smart Object source width)
~~~

The scale bar is added only to the selected `Group_Panel_<id>`. Its rectangle and optional contrasting outline are distinct editable raster layers; the caption is a separate editable Photoshop text layer. The group-relative placement is bottom-right with a configurable inset. The input PSD is never saved in place; output must be a different `.psd` path.

## unify_fonts: current implementation

The Python API is:

~~~python
from ps_sci_toolkit import unify_fonts

result = unify_fonts(
    "Figure_1_scalebar.psd",
    "Figure_1_fonts.psd",
    font_postscript="ArialMT",
    size_pt=7,
    color_hex="#000000",
)
~~~

At least one normalization attribute is required. Any attribute passed as `None` is left unchanged. The operation traverses every document-level and nested `LayerSet`, finds only native Photoshop `LayerKind.TEXT` layers, and applies the requested font PostScript name, point size, and/or RGB color in one JSX execution. It then reads each changed value back for verification and saves a separate editable PSD. Raster layers and embedded Smart Objects are not modified. The source PSD is never saved in place.

## export_figure: current implementation

The Python API is:

~~~python
from ps_sci_toolkit import export_figure

result = export_figure(
    "Figure_1_fonts.psd",
    "Figure_1_exports",
    basename="Figure_1",
    expected_dpi=300,
)
~~~

The bundle contains:

- `Figure_1.psd`: byte-identical copy of the editable source PSD, retaining its native groups, Smart Objects, and text layers.
- `Figure_1.tif`: flattened RGB TIFF saved by Photoshop with LZW compression, embedded color profile, and IBM byte order for Windows.
- `Figure_1.png`: flattened RGB PNG saved by Photoshop.

The TIFF and PNG are created from two independent duplicates of the source document. Only those duplicates are flattened; the input PSD is never flattened or saved. Canvas dimensions are not resampled and source resolution must be 300 or 600 DPI. An optional `expected_dpi` checks the requested value without changing resolution.

Before publishing the output bundle, Photoshop reopens all three staged files and checks their dimensions, RGB mode, and resolution. Python then reads TIFF/PNG metadata only: RGB mode, pixel dimensions, DPI, and TIFF compression tag 259 (value 5 = LZW). It does not re-encode the images. Output files are staged under temporary names and moved into their final paths only after export and metadata checks succeed; the source PSD's SHA-256 must remain unchanged.

The CLI options are `--unify-font-postscript`, `--unify-font-size-pt`, and `--unify-font-color`; specify at least one. The real Photoshop demo deliberately inserts a text layer two groups deep before normalization to prove the recursive traversal, then reopens the result to check every text layer.
Implement and test each feature separately. Keep the Python public API stable where practical and put native Photoshop operations in JSX helpers.

## Error handling

Every bridge failure must identify:

- operation (connect, status, lock_units, execute_jsx, compose_grid, add_panel_labels, or add_scale_bar);
- Python exception class, if any;
- original exception message;
- JSX preview when applicable;
- whether Photoshop was detected and connected.

A failure must not be converted into a success-shaped return value. CLI exit code 0 means the selected operation passed; non-zero means it failed or prerequisites were missing.

## Validation

Run platform-independent unit tests:

~~~powershell
python -m unittest discover -s tests -v
~~~

Then run the real Photoshop integration checks on the Windows machine:

~~~powershell
python scripts/ps_sci_toolkit.py --action diagnose
python tests/test_compose_grid_demo.py
python tests/test_add_panel_labels_demo.py --input-psd path\\to\\Figure_Grid_Stage2.psd
python tests/test_add_scale_bar_demo.py --input-psd path\\to\\Figure_Grid_Stage2_Stage3_Labeled.psd
python tests/test_unify_fonts_demo.py --input-psd path\\to\\Figure_Grid_Stage4_ScaleBar.psd
~~~

A bridge diagnostic pass requires COM connection, JSX execution, successful unit read-back, and controlled JSX exception capture. Stage 2 and Stage 3 integration passes additionally require the respective saved PSDs to be reopened in Photoshop and their native document/layer structures to be inspected. Unit tests alone are not proof of Photoshop compatibility.

## Confirmed environment baseline

The initial real-machine test for this skill reported:

- Python 3.13.9 at D:\Software\Anacoda\python.exe.
- photoshop-python-api 0.24.2.
- Photoshop 2026, version 27.9.1 (file version 27.9).
- Photoshop.Application COM registration and connection succeeded.
- Application.doJavaScript() succeeded.
- Ruler units verified as Units.PIXELS.
- Type units verified as TypeUnits.POINTS.
- A deliberately raised and caught ExtendScript exception returned its marker correctly.

This baseline is specific to the tested Windows machine. Re-run diagnostics after changing Photoshop, Python, or the package version.

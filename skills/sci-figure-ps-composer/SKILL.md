---
name: sci-figure-ps-composer
description: Automate publication-ready scientific multi-panel figure composition in desktop Adobe Photoshop on Windows, using Python COM automation and native ExtendScript/JSX while preserving editable PSD layers.
---

# sci-figure-ps-composer

## Purpose

Create and maintain publication-ready scientific multi-panel figures in the Windows desktop edition of Adobe Photoshop. Python is the orchestration and geometry layer; Photoshop remains the native document, typography, layer, smart-object, and export engine.

The implementation is staged. Stages 1–8 and the Stage 9A partial-preview workflow have passed real-machine verification on the user's Windows installation running Photoshop 2026 (27.9.1). The Photoshop Bridge, `compose_grid`, editable panel labels, calibrated scale bars, recursive text-layer font normalization, and Stage 6 PSD/TIFF/PNG publication export were verified individually. Stage 7 read-only publication preflight has also passed against the Stage 5 PSD and Stage 6 export bundle, including Smart Object inspection, effective PPI, scale calibration, export metadata, source-hash preservation, and restoration of Photoshop's prior document and preferences. Stage 8 adds a quiet/interactive execution policy and read-only inspection plus approved, structure-only organization plans for existing PSDs. Both Stage 8A inspection and Stage 8B Smart Object reparenting passed real Photoshop 2026 (27.9.1) integration tests; the source PSD hash remained unchanged. Stage 9A adds isolated layer/group PNG previews, a full-composite preview, and a labeled contact sheet for visual review. It passed real Photoshop 2026 (27.9.1) verification on the 19-layer test PSD: 12 selected layer/group previews succeeded, the composite/contact sheet/report were created, and the original PSD SHA-256 remained unchanged. The initial real-machine run intentionally capped previews at 12 of 19 nodes. The Stage 9A integration demo now requires complete coverage of every non-background node by default and fails if the selected limit truncates the layer tree or any isolated preview fails; `--allow-partial` is reserved for explicitly partial diagnostic runs.

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

# Run a non-destructive publication preflight on an editable PSD.
python scripts/ps_sci_toolkit.py --action preflight-figure `
  --input-psd Figure_1_fonts.psd --preflight-dpi 300 `
  --preflight-min-panel-ppi 300 --preflight-export-dir Figure_1_exports `
  --preflight-calibration-json Figure_1_calibration.json

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
~~~

The Stage 5 demo creates a separate fixture PSD with an additional text layer nested two groups deep, then normalizes the font, point size, and color of every editable text layer into a new PSD. It reopens that output and verifies every text layer, including the nested test layer, as well as preservation of the four panel labels and scale-bar label. Both the supplied input PSD and the generated fixture PSD are checked for unchanged hashes.

After Stage 5 passes, run the Stage 6 publication export integration test:

~~~powershell
python tests/test_export_figure_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd
~~~

The Stage 6 test creates a new export folder containing a byte-identical editable PSD copy, a flattened RGB TIFF using LZW compression, and a flattened RGB PNG. Photoshop reopens the PSD/TIFF/PNG and verifies their dimensions, RGB mode, resolution, panel groups, Smart Objects, and text layers. Pillow reads—but does not rewrite—the raster exports to verify RGB mode, dimensions, DPI metadata, and TIFF compression tag 259 = 5 (LZW). The source PSD hash must remain unchanged.

After Stage 6 passes, run the Stage 7 read-only publication preflight on the Stage 5 editable PSD and its Stage 6 export bundle:

~~~powershell
python tests/test_preflight_figure_demo.py `
  --input-psd path\to\Figure_Grid_Stage5_FontsUnified.psd `
  --export-dir path\to\Figure_Grid_Stage5_FontsUnified_exports `
  --export-basename Figure_Grid_Stage2_Stage3_Labeled_Stage4_ScaleBar_Stage5_FontsUnified
~~~

The Stage 7 demo inspects canvas geometry, DPI, color mode/profile, panel groups, embedded Smart Objects, effective PPI, panel-label/text styling, scale-bar geometry, the known synthetic scale calibration, and the Stage 6 PSD/TIFF/PNG bundle. It never saves the inspected PSD and restores the prior active document, dialog setting, ruler units, and type units. The demo intentionally permits warnings for the low-resolution synthetic panel rasters and the font-normalization test's 9 pt Arial text; use `--strict` to turn warnings into failures. This integration test passed on the user's Photoshop 2026 installation.

### Bridge behavior

1. Import photoshop.api and create photoshop.api.Application().
2. Execute JSX through the package's Application.doJavaScript() method.
3. Normalize Python/COM/JSX failures as PhotoshopBridgeError, preserving the operation stage, exception type, message, and a short JSX preview.
4. Lock app.preferences.rulerUnits to Units.PIXELS.
5. Lock app.preferences.typeUnits to TypeUnits.POINTS.
6. Verify both preferences by reading them back from Photoshop.

The unit lock intentionally changes Photoshop's global preferences and leaves them set to Pixels and Points. It does not create, close, or modify a document.

### Execution modes

Every CLI action accepts `--execution-mode quiet|interactive`; the default is `quiet`.

- `quiet` replaces script dialog calls with `DialogModes.NO` so supported operations do not wait for Photoshop dialogs.
- `interactive` permits Photoshop dialogs by using `DialogModes.ALL`; use it when troubleshooting a save/export prompt or intentionally running a dialog-driven script.
- Both modes require a running desktop Photoshop process connected through Windows COM. This is not a true server/headless mode: Photoshop must remain installed and running, ExtendScript/COM execution is synchronous and blocks Photoshop while the script runs, and the toolkit does not guarantee that Photoshop will never become the foreground window.
- The bridge restores the previous `app.displayDialogs` value and unit preferences in document operations that temporarily change them. The explicit `lock-units` action and `diagnose` intentionally leave Pixels/Points set as documented above.

Example:

~~~powershell
python scripts/ps_sci_toolkit.py --action inspect-psd --input-psd Figure_legacy.psd --execution-mode quiet --json
~~~

For background/batch use, keep Photoshop open in the user's Windows desktop session and avoid concurrently interacting with the same document while a synchronous script is running.

## Agent workflow

### Before any document operation

- Confirm the user is on Windows and identify the exact Python interpreter.
- Confirm Photoshop is running. Do not launch or terminate Photoshop automatically.
- Run --action diagnose and inspect its structured result.
- If COM connection fails, stop before attempting document changes. Report the stage and original error.
- If JSX execution works but unit verification fails, stop and report the observed values; do not silently proceed with geometry calculations.
- Never claim real Photoshop integration passed based only on mocked tests.
- For an existing PSD, run `inspect-psd` before proposing any structural edits. Preserve the returned ordered layer tree and source hashes.
- Show the user the exact proposed organizer operations, affected layer names/paths, and any uncertain classifications. Do not infer scientific meaning from names or visual appearance alone.
- Do not run `apply-organize-plan` until the user explicitly approves the displayed plan. The plan must bind to both `source_sha256` and `layer_tree_sha256`; a stale plan must be regenerated.
- Organizer v1 may create groups, rename layers, move/reorder layers, and change visibility. It must not delete, merge, flatten, rasterize, resample, crop, transform, or alter pixel content. It rejects moving background, locked, position-locked, or clipping-mask layers.
- Always write the organized result to a new PSD path. Retain the user's original PSD and verify its SHA-256 is unchanged.

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

## Feature milestones

1. Photoshop Bridge, `compose_grid`, `add_panel_labels`, `add_scale_bar`, `unify_fonts`, `export_figure`, and `preflight_figure`: implemented; Stages 1–7 passed real Photoshop verification on the user's Photoshop 2026 installation.
2. Stage 8 execution modes: `quiet` and `interactive` are implemented; unit coverage is included. `quiet` is not true headless execution.
3. Stage 8A `inspect_psd`: implemented; returns a read-only, mixed-order layer tree and source integrity hashes. Real Photoshop 2026 integration passed against a 19-layer research figure (6 groups, 6 text layers, 4 Smart Objects).
4. Stage 8B `apply_organize_plan`: implemented as an explicitly approved, structure-only plan executor operating on a temporary byte-identical PSD copy and saving to a new PSD. Real Photoshop 2026 integration passed by moving an embedded Smart Object into a new group, reopening the output, and verifying 20 layers, unchanged text/Smart Object counts, and unchanged source SHA-256.
5. Stage 9A `preview_psd_layers`: implemented; exports a composite preview and isolated previews of selected layers/groups from disposable copies, then builds a labeled contact sheet and JSON manifest. Real Photoshop 2026 integration passed: 12/12 selected previews succeeded, 0 failed, and source SHA-256 was unchanged. The tested run selected only the first 12 nodes; full-tree visual coverage still requires a run with a larger `--preview-max-layers` CLI value.
6. Stage 9B evidence-grounded classification and plan drafting: implemented with a model-independent JSON contract and Python validation. It checks full node coverage, unique node keys, source/tree fingerprints, confidence, role compatibility, and evidence against actual inspection/preview artifacts. It can produce an unapproved structure-only organizer plan draft; the user must inspect it and separately approve execution. Real Photoshop is not required for this metadata/JSON validation stage.

## Existing PSD inspection and organization

This workflow is for PSDs created outside this Skill or PSDs whose layer structure has become difficult to maintain. It is separate from `compose_grid`, which creates a new figure from raster inputs.

### Inspect before proposing changes

~~~powershell
python scripts/ps_sci_toolkit.py --action inspect-psd `
  --input-psd Figure_legacy.psd `
  --inspection-report Figure_legacy_inspection.json `
  --execution-mode quiet
~~~

`inspect_psd` is read-only. It first copies the saved on-disk PSD byte-for-byte to a unique temporary file and inspects that copy, then verifies the original source SHA-256 remains unchanged. This deliberately excludes unsaved in-memory changes in an already-open Photoshop document: save the intended version to disk before inspection. The temporary copy name is hidden from layer paths; the report uses the original PSD filename. It records document geometry/profile and recursively enumerates the Photoshop `Document.layers` / `LayerSet.layers` collection so interleaved ArtLayer and LayerSet stacking order is retained. Each node contains a stable-within-this-source `node_key` derived from its ordered tree position, parent key, path, name, kind, visibility, opacity, blend mode, lock/clipping flags, pixel bounds, and editable text properties when applicable. The report includes the source PSD SHA-256 and a layer-tree fingerprint.

Node keys are valid only for the inspected source hash and tree fingerprint. They must not be reused after renaming, moving, adding, deleting, or reordering source layers. Run inspection again after any source change.

### Approved organizer plan

The agent must explain the intended changes, including the exact layer paths and operations, then obtain explicit user approval. The plan is JSON:

~~~json
{
  "schema_version": 1,
  "source_sha256": "<copy source_sha256 from inspection report>",
  "layer_tree_sha256": "<copy layer_tree_sha256 from inspection report>",
  "operations": [
    {
      "action": "create_group",
      "key": "panel_a",
      "name": "Panel_A",
      "parent": "root"
    },
    {
      "action": "rename_layer",
      "target": "root/G0/L1",
      "expected_name": "Layer 5 copy",
      "new_name": "Panel_A_Annotation"
    },
    {
      "action": "move_layer",
      "target": "root/G0/L1",
      "expected_name": "Layer 5 copy",
      "destination": "panel_a",
      "position": "end"
    },
    {
      "action": "reorder_layer",
      "target": "root/L2",
      "expected_name": "Legend",
      "relative_to": "root/L1",
      "relative_expected_name": "Panel_A",
      "position": "before"
    },
    {
      "action": "set_visibility",
      "target": "root/L3",
      "expected_name": "Old notes",
      "visible": false
    }
  ]
}
~~~

Supported operations are:

- `create_group`: create a new named LayerSet under `root`, an existing LayerSet node key, or a group created earlier in the plan. Group keys are unique plan-local identifiers.
- `rename_layer`: rename an existing layer or group after checking `expected_name`.
- `move_layer`: move an existing layer/group into `root`, an existing LayerSet, or a created group; `position` is `start` or `end`.
- `reorder_layer`: move a node `before` or `after` an existing sibling.
- `set_visibility`: explicitly set `visible` to true or false.

All `create_group` operations must appear first. The executor then applies renames, moves, reorders, and visibility changes in that order. Every `expected_name` and `relative_expected_name` refers to the original inspection report, even when the same plan also renames that layer. It rejects stale hashes, unknown node keys, ambiguous or unsafe parent relationships, background/locked/position-locked/clipping-chain moves or reorders, and individual or combined group-parent cycles. New root-level groups and moves are placed above Photoshop's locked Background layer.

Apply only after approval:

~~~powershell
python scripts/ps_sci_toolkit.py --action apply-organize-plan `
  --input-psd Figure_legacy.psd `
  --organize-plan Figure_legacy_organize_plan.json `
  --output Figure_legacy_organized.psd `
  --confirm-plan `
  --execution-mode quiet
~~~

The CLI requires `--confirm-plan`. The API requires `approved=True`. The executor copies the source PSD to a temporary working file, applies only the approved structural operations, saves to a temporary PSD in the output directory, reopens and verifies the edited structure, then atomically moves the verified file to the requested output path. The original PSD is never saved or overwritten. Existing output files are protected unless `--overwrite` is explicit.

Organizer v1 does not delete, merge, flatten, rasterize, resample, crop, transform, or alter pixel data. It cannot reconstruct independent layers from a flattened image. Clipping-mask moves are rejected because reparenting can change their rendering relationship. Unknown scientific meaning remains a human classification decision, not an automatic guess.

## Stage 9B: semantic classification and reviewable plan drafting

The Photoshop Skill does not run a hidden vision model or silently infer scientific meaning. The Antigravity agent (or another explicitly selected vision-capable model) reviews the Stage 9A composite/contact sheet and isolated previews, authors a JSON classification report, and asks the user to resolve ambiguous roles. The Python toolkit validates that report against the exact PSD and evidence artifacts.

### Classification report contract

One record is required for every inspected `node_key`. Use `unknown` when the evidence does not support a reliable role. Do not omit uncertain or hidden nodes.

Allowed roles:

- `panel_group`, `panel_image`, `panel_label`
- `scale_bar`, `scale_label`
- `axis`, `axis_label`, `legend`, `title`, `annotation`
- `decoration`, `background`, `test_artifact`, `unknown`

Each record contains a confidence in [0, 1], a concise rationale, and at least one typed evidence item. Evidence types are `layer_name`, `layer_type`, `layer_metadata`, `text_content`, `geometry`, `visual_preview`, `composite_context`, and `user_instruction`.

### Generate a classification work template

After inspection and preview generation, create an evidence-indexed template before asking the vision-capable agent to classify nodes:

~~~powershell
python scripts/ps_sci_toolkit.py --action create-psd-classification-template `
  --input-psd Figure_legacy.psd `
  --inspection-report Figure_legacy_inspection.json `
  --preview-report Figure_legacy_visual_index/preview_report.json `
  --output Figure_legacy_classification_WORK_TEMPLATE.json `
  --execution-mode quiet --json
~~~

The template includes every inspected node exactly once, the source/tree hashes, exact layer metadata, text and geometry facts, available evidence candidates, and the local isolated-preview file associated with each node. Its summary separately records total nodes, Photoshop Background layers, preview successes/failures, missing preview keys, and whether every non-background node has a successful isolated preview. The template generator does not modify the PSD and refuses stale inspection/preview hashes or missing referenced preview files.

The template deliberately does not finalize scientific roles based on layer names. Non-background nodes remain pending with `role: unknown`, no confidence, and no fabricated evidence until the agent actually reviews them. It may include a separate `role_suggestion` as a navigation hint when conservative rules match multiple independent facts (for example, a panel-label name, native text type, and matching one-letter text content; or a Smart Object name whose suffix matches its parent panel group). These are candidate roles only, explicitly marked `is_final_classification: false`, and require visual confirmation. A name-only or type-only match is not promoted to a candidate role. A Photoshop-native Background layer is the one exception: it can be recorded as structural role `background` with confidence 1.0, supported by exact `layer_metadata` evidence for `is_background: true`. The work template is not itself a valid classification report; the agent must review candidates, complete the entries, and produce a separate classification JSON before validation.

~~~json
{
  "schema_version": 1,
  "source_sha256": "<copy from inspection_report.json>",
  "layer_tree_sha256": "<copy from inspection_report.json>",
  "classifications": [
    {
      "node_key": "root/G0/L0",
      "role": "panel_label",
      "confidence": 0.98,
      "rationale": "The node is an editable text layer named Panel_Label_d and its text is the panel identifier.",
      "evidence": [
        {
          "type": "layer_name",
          "observation": "Panel_Label_d"
        },
        {
          "type": "text_content",
          "observation": "d"
        },
        {
          "type": "visual_preview",
          "preview_node_key": "root/G0/L0",
          "observation": "The isolated preview shows a small lowercase panel marker near the panel corner."
        }
      ]
    },
    {
      "node_key": "root/G0/L1",
      "role": "panel_image",
      "confidence": 0.99,
      "rationale": "The embedded Smart Object is the visual image content inside the corresponding panel group.",
      "evidence": [
        {
          "type": "layer_type",
          "observation": "LayerKind.SMARTOBJECT"
        },
        {
          "type": "visual_preview",
          "preview_node_key": "root/G0/L1",
          "observation": "The isolated preview shows the panel's scientific image."
        },
        {
          "type": "composite_context",
          "observation": "The same image occupies the visual body of panel d in the composite."
        }
      ]
    }
  ]
}
~~~

The example is illustrative only; a usable report must include every node in the inspection JSON exactly once. A partial Stage 9A preview run can still support metadata-based classifications, but it cannot support `visual_preview` evidence for nodes that were not previewed. Run Stage 9A with a sufficiently high `--preview-max-layers` CLI value (up to 200) when full visual coverage is needed.

Evidence validation rules:

- `layer_name` must quote the exact inspected layer name.
- `layer_type` must match the inspected Photoshop kind or typename.
- `layer_metadata` must contain exact, type-matching values from the inspected node (for example, `visible`, `opacity`, lock flags, clipping state, or `is_background`); unsupported fields are rejected.
- A node marked by Photoshop as the native Background layer must use role `background`, supported by validated `layer_metadata` evidence with `is_background: true`. The role cannot be assigned to an ordinary layer merely by naming it “Background”.
- `text_content` must quote text actually present in a native text layer.
- `geometry` must include the exact inspected `bounds_px` array.
- `visual_preview` must refer to a successful isolated preview for the same node key.
- `composite_context` requires an existing composite preview in a matching preview manifest.
- `user_instruction` must explicitly set `source: "user"`.
- Confidence below 0.45 produces a warning; visual roles without visual evidence produce a warning. Warnings do not authorize automatic actions.

Validate a model-authored classification:

~~~powershell
python scripts/ps_sci_toolkit.py --action validate-psd-classification `
  --input-psd Figure_legacy.psd `
  --inspection-report Figure_legacy_inspection.json `
  --preview-report Figure_legacy_visual_index/preview_report.json `
  --classification-report Figure_legacy_classification.json `
  --execution-mode quiet --json
~~~

The validator checks that the current PSD hash still matches the inspection report, that the classification references the same layer-tree fingerprint, that every inspected node is accounted for exactly once, and that all supplied evidence references are valid. It does not change the PSD.

### Create an unapproved organizer-plan draft

The agent may prepare a separate JSON array of proposed structure-only operations after analyzing the classification and discussing the desired structure with the user:

~~~json
[
  {
    "action": "create_group",
    "key": "panel_d_clean",
    "name": "Panel_D",
    "parent": "root"
  },
  {
    "action": "move_layer",
    "target": "root/G0/L1",
    "expected_name": "Image_Panel_d",
    "destination": "panel_d_clean",
    "position": "end"
  }
]
~~~

Create a reviewable draft:

~~~powershell
python scripts/ps_sci_toolkit.py --action create-organize-plan-draft `
  --input-psd Figure_legacy.psd `
  --inspection-report Figure_legacy_inspection.json `
  --preview-report Figure_legacy_visual_index/preview_report.json `
  --classification-report Figure_legacy_classification.json `
  --plan-operations Figure_legacy_proposed_operations.json `
  --output Figure_legacy_organize_plan_DRAFT.json `
  --execution-mode quiet --json
~~~

The draft records `requires_user_approval: true`, `approved: false`, source/tree hashes, classification hash, warnings, and the normalized operations. It is not executable merely because it was generated. The agent must display the exact before/after layer paths and every operation, explain visual/structural consequences, and obtain explicit user approval. Only then may the separate `apply-organize-plan --confirm-plan` workflow be invoked, which still applies changes to a temporary PSD copy and saves a new output PSD.

Never propose moving a layer solely because its name looks like a panel/image. Preserve layer visibility, clipping relationships, masks, and Smart Objects; mark ambiguous scientific purpose as `unknown` and ask the user rather than inventing semantics.

## Existing PSD visual preview index

After `inspect_psd` returns the ordered layer tree, use `preview_psd_layers` to understand the actual visual content before proposing structural changes.

~~~powershell
python scripts/ps_sci_toolkit.py --action preview-psd-layers `
  --input-psd Figure_legacy.psd `
  --preview-dir Figure_legacy_visual_index `
  --preview-max-layers 24 `
  --preview-width-px 512 `
  --execution-mode quiet
~~~

The output directory must be new; existing directories are never overwritten. Outputs include:

- `composite.png`: a downsampled, flattened preview of the entire source figure.
- `layer_001.png`, ...: isolated previews of selected original layers or groups, generated from disposable duplicates. A group preview restores its children's original visibility. A clipped layer preview attempts to include its base layer.
- `layer_contact_sheet.png`: a labeled grid with each node key, original layer name, and layer kind.
- `preview_report.json`: source SHA-256, layer-tree fingerprint, total inspected node count, eligible non-background node count, number actually previewed, explicit `full_visual_coverage`, missing/failed node-key lists, original paths, per-preview status, and output metadata.
- `full_visual_coverage: true` means every non-background node from that exact inspection was previewed successfully. A successful partial run can still have `passed: true`; never treat `passed` alone as proof of complete coverage. Use `missing_preview_node_keys` and `failed_preview_node_keys` to diagnose any gap.

By default, non-background layers and groups are selected in the original stack order. Use repeated `--preview-node-key` options to request a smaller, targeted set. The toolkit default maximum is 80, bounded by Photoshop memory and document complexity. The real-Photoshop integration demo defaults to 200 and requires every non-background node to be previewed successfully; it rejects accidental partial coverage unless `--allow-partial` is explicitly supplied. Layer previews are visual aids, not an exact reconstruction of every blend/adjustment dependency: complex blend modes, clipping chains, masks, and adjustment layers may require contextual interpretation. The composite preview is the authoritative visual reference for the whole figure.

Example selecting a group and two image layers:

~~~powershell
python scripts/ps_sci_toolkit.py --action preview-psd-layers `
  --input-psd Figure_legacy.psd `
  --preview-dir Figure_legacy_selected_layers `
  --preview-node-key root/G1 `
  --preview-node-key root/G2/L0 `
  --preview-node-key root/G3/L1
~~~

Only after reviewing the composite and layer contact sheet should the agent classify likely panels/images/text/legends/decorations. Classification must carry node keys, visual/text evidence, confidence, and `unknown` for uncertain purpose. It must not treat layer names alone as proof of scientific meaning.

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

Before publishing the output bundle, Photoshop reopens all three staged files and checks their dimensions, RGB mode, and resolution. Python then reads TIFF/PNG metadata only: RGB mode, pixel dimensions, DPI, and TIFF compression tag 259 (value 5 = LZW). It does not re-encode the images. Output files are staged under short UUID-based temporary names in the export directory (the staging names deliberately do not repeat the potentially long scientific figure basename) and moved into their final paths only after export and metadata checks succeed. Filesystem failures identify the current export phase. The source PSD's SHA-256 must remain unchanged.

The CLI options are `--unify-font-postscript`, `--unify-font-size-pt`, and `--unify-font-color`; specify at least one. The real Photoshop demo deliberately inserts a text layer two groups deep before normalization to prove the recursive traversal, then reopens the result to check every text layer.

## preflight_figure: current implementation

The Python API is:

~~~python
from ps_sci_toolkit import preflight_figure

result = preflight_figure(
    "Figure_1_fonts.psd",
    expected_dpi=300,
    expected_width_cm=17.6,
    min_panel_ppi=300,
    calibrations={
        "a": {
            "physical_length": 10,
            "pixel_size": 0.2,
            "unit": "µm",
            "tolerance_px": 2.0,
        }
    },
    export_dir="Figure_1_exports",
)
~~~

The operation is read-only: Photoshop opens the PSD, inspects document and layer data, opens each panel Smart Object only to read its embedded pixel dimensions, and closes all inspection documents without saving. It restores the prior active document, dialog setting, ruler units, and type units. The input PSD's SHA-256 is checked before and after preflight.

Checks are returned as individual `PASS`, `WARN`, `FAIL`, or `INFO` entries:

- Canvas geometry, RGB mode, 300/600 DPI, optional journal width, and color-profile presence.
- One image Smart Object per panel group, panel bounds within the canvas, and overlapping panel-image bounds.
- Effective panel PPI calculated from the embedded Smart Object's post-trim pixel dimensions, its on-canvas displayed bounds, and document DPI.
- Panel label presence, editable text, expected Arial Bold 10–12 pt black styling, identifier content, and position.
- Figure text font/size/color and canvas containment.
- Scale-bar rectangle/outline/label geometry. Physical width and label are independently verified only when a calibration manifest is supplied; the PSD alone does not establish physical units.

A calibration manifest is a JSON object keyed by panel ID, for example:

~~~json
{
  "a": {
    "physical_length": 10,
    "pixel_size": 0.2,
    "unit": "µm",
    "tolerance_px": 2
  }
}
~~~

Here `pixel_size` is physical units per embedded Smart Object source pixel. The expected displayed bar width is computed from that calibration and the Smart Object's actual post-trim dimensions and on-canvas transform.

When `export_dir` is supplied, preflight also checks the byte-identical editable PSD copy, RGB TIFF dimensions/DPI/LZW tag, and RGB PNG dimensions/DPI. Without an export directory or calibration manifest, those checks are explicitly marked as not assessed rather than guessed.

Typography and effective-PPI deviations are warnings by default because figures may have journal-specific exceptions. Use `strict=True` or CLI `--preflight-strict` to promote all warnings to failures. Structural corruption, wrong color mode/DPI, invalid geometry, broken scale-bar components, failed supplied calibrations, or invalid exports are failures regardless of strict mode.

The CLI action is `preflight-figure`; useful options include `--preflight-min-panel-ppi`, `--preflight-dpi`, `--preflight-width-cm`, `--preflight-calibration-json`, `--preflight-export-dir`, and `--preflight-strict`. The Stage 7 Windows demo uses the known synthetic Stage 4 calibration (10 µm at 0.2 µm/source px) only for the generated demo figure; real research figures must supply their own manifest.
Implement and test each feature separately. Keep the Python public API stable where practical and put native Photoshop operations in JSX helpers.

## Error handling

Every bridge failure must identify:

- operation (connect, status, lock_units, execute_jsx, compose_grid, add_panel_labels, add_scale_bar, unify_fonts, export_figure, or preflight_figure);
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
python tests/test_export_figure_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd
python tests/test_preflight_figure_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd --export-dir path\\to\\Figure_Grid_Stage5_FontsUnified_exports --export-basename Figure_Grid_Stage2_Stage3_Labeled_Stage4_ScaleBar_Stage5_FontsUnified
python tests/test_inspect_psd_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd
python tests/test_apply_organize_plan_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd
python tests/test_preview_psd_layers_demo.py --input-psd path\\to\\Figure_Grid_Stage5_FontsUnified.psd --preview-dir path\\to\\Figure_Grid_Stage8_Layer_Previews

Both Stage 8A and Stage 8B integration demos passed on Photoshop 2026 (27.9.1). Stage 9A `preview_psd_layers` passed real Photoshop verification: 12/12 selected previews were generated successfully, with a composite, contact sheet, manifest, and unchanged source PSD hash. Stage 9B classification validation and unapproved plan drafting are covered by GitHub Actions Run #158 (84 tests passed on Python 3.13); these are model-independent JSON/metadata operations and do not require a Photoshop integration test. The inspected source PSD contained 19 layers, 6 groups, 6 text layers, and 4 Smart Objects. Stage 8B moved the embedded `Image_Panel_a` Smart Object into a new test group on a separate PSD copy; the output reopened with 20 layers and preserved the original 6 text layers and 4 Smart Objects. The original PSD SHA-256 remained unchanged.
~~~

A bridge diagnostic pass requires COM connection, JSX execution, successful unit read-back, and controlled JSX exception capture. Each staged integration pass requires the generated PSD or export bundle to be reopened in Photoshop and its native document/layer structure or export metadata to be inspected. The Stage 7 integration pass additionally verifies that the inspected source PSD remains unchanged. Unit tests alone are not proof of Photoshop compatibility.

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

The Stage 7 read-only publication preflight was also executed successfully on this installation against the Stage 5 editable PSD and Stage 6 export bundle; its required checks passed and the source PSD hash remained unchanged. This baseline is specific to the tested Windows machine. Re-run diagnostics after changing Photoshop, Python, or the package version.

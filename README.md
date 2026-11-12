# Photoshop Agents

A collection of agent skills and automation tools for publication-oriented Photoshop workflows.

## Skills

### `sci-figure-ps-composer`

Automates scientific multi-panel figure composition and maintenance in desktop Adobe Photoshop on Windows.

Current capabilities include:

- Compose multi-panel figures from raster inputs using editable embedded Smart Objects.
- Add editable panel labels and calibrated scale bars.
- Normalize text-layer typography recursively.
- Export editable PSD, RGB LZW TIFF, and PNG deliverables.
- Run publication preflight checks for geometry, effective panel resolution, typography, calibration, and export metadata.
- Inspect existing PSD layer trees, generate full-coverage isolated previews, classify layers with evidence, and draft risk-assessed organization plans. Low-risk renames/grouping and high-risk moves/reorders are reviewed together; visibility changes are separately acknowledged and require an additional execution flag. Destructive/pixel-changing operations are outside organizer v1.

For an existing PSD, the recommended entry point is `prepare-existing-psd`. It produces a reusable inspection report, full-coverage layer previews, a classification work template, and a workflow manifest while reusing the same inspection inventory rather than inspecting the PSD twice.

## Architecture and safety

- Photoshop remains the native document, rendering, typography, Smart Object, and export engine.
- Python handles orchestration, geometry, evidence indexing, validation, and the CLI.
- ExtendScript/JSX performs native Photoshop document operations.
- Existing PSD edits are applied to a temporary copy and saved to a new PSD path.
- Structure-changing operations require explicit user approval; inspection and classification do not create or apply edits.
- Unit tests run in GitHub Actions. Real Photoshop integration is verified separately on Windows with desktop Photoshop.

## Runtime requirements

- Windows desktop session with Adobe Photoshop installed and running.
- Python 3.10 or later.
- Dependencies listed in `skills/sci-figure-ps-composer/requirements.txt`.

See [the skill instructions](skills/sci-figure-ps-composer/SKILL.md) for CLI usage, workflows, safety rules, and publication-output details.

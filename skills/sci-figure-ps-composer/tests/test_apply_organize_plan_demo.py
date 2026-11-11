#!/usr/bin/env python
"""Real-Photoshop Stage 8 approved organizer integration demo.

Moves one eligible ArtLayer into a newly created test group on a temporary
working copy, saves a separate PSD, reopens it, and verifies source integrity.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import tempfile
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from ps_sci_toolkit import PhotoshopBridge, apply_organize_plan, inspect_psd


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-psd", type=Path, required=True)
    parser.add_argument("--output-psd", type=Path)
    parser.add_argument("--execution-mode", choices=("quiet", "interactive"), default="quiet")
    args = parser.parse_args()

    source = args.input_psd.expanduser().resolve(strict=True)
    source_hash_before = sha256(source)
    bridge = PhotoshopBridge(execution_mode=args.execution_mode)
    inspection = inspect_psd(source, bridge=bridge)

    candidates = [
        layer for layer in inspection["layers"]
        if layer["typename"] == "ArtLayer"
        and not layer["is_background"]
        and not layer["all_locked"]
        and not layer["position_locked"]
        and not layer["clipped_to_below"]
    ]
    if not candidates:
        raise RuntimeError(
            "No movable, unlocked, non-background, non-clipping ArtLayer was found."
        )
    target = candidates[-1]

    plan = {
        "schema_version": 1,
        "source_sha256": inspection["source_sha256"],
        "layer_tree_sha256": inspection["layer_tree_sha256"],
        "operations": [
            {
                "action": "create_group",
                "key": "organizer_test",
                "name": "Organizer_Test_Group",
                "parent": "root",
            },
            {
                "action": "move_layer",
                "target": target["node_key"],
                "expected_name": target["name"],
                "destination": "organizer_test",
                "position": "end",
            },
        ],
    }

    temp_context = None
    if args.output_psd:
        output = args.output_psd.expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
    else:
        temp_context = tempfile.TemporaryDirectory(prefix="sci_org_demo_")
        output = Path(temp_context.name) / "Organizer_Integration_Test.psd"

    try:
        result = apply_organize_plan(
            source,
            plan,
            output,
            approved=True,  # This is a disposable integration-test plan.
            bridge=bridge,
        )
        if not output.is_file() or output.stat().st_size <= 0:
            raise RuntimeError("Organizer did not produce a non-empty output PSD.")
        if sha256(source) != source_hash_before or not result["source_sha256_unchanged"]:
            raise RuntimeError("The source PSD was modified by the organizer.")

        verified = result["verified_references"]
        group = verified["group:organizer_test"]
        moved_layer = verified["source:" + target["node_key"]]
        if group["typename"] != "LayerSet" or group["name"] != "Organizer_Test_Group":
            raise RuntimeError("The new organizer group failed verification.")
        if moved_layer["parent_key"] != group["node_key"]:
            raise RuntimeError("The selected layer was not moved into the new group.")
        if result["layer_count_after"] != result["layer_count_before"] + 1:
            raise RuntimeError("The organizer unexpectedly changed the number of layers.")

        report = {
            "passed": True,
            "input_psd": str(source),
            "output_psd": str(output),
            "moved_layer": target["path"],
            "moved_layer_name": target["name"],
            "new_group": group["name"],
            "source_sha256_unchanged": True,
            "layer_count_before": result["layer_count_before"],
            "layer_count_after": result["layer_count_after"],
            "output_layer_tree_sha256": result["output_layer_tree_sha256"],
        }
        print(json.dumps(report, ensure_ascii=False, indent=2))
        print("STAGE8_APPLY_ORGANIZE_PLAN_REAL_PHOTOSHOP_TEST=PASS")
        return 0
    finally:
        if temp_context is not None:
            temp_context.cleanup()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "STAGE8_APPLY_ORGANIZE_PLAN_REAL_PHOTOSHOP_TEST=FAILED\n"
            + type(exc).__name__ + ": " + str(exc),
            file=sys.stderr,
        )
        raise SystemExit(1)

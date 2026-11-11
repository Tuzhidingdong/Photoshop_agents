# Photoshop Agents

A collection of agent skills and automation tools for desktop Adobe Photoshop workflows.

## Skills

- \`sci-figure-ps-composer\`: publication-oriented scientific multi-panel figure composition using Photoshop, Python, and ExtendScript/JSX.

## Development principles

- Keep Photoshop as the native document and rendering engine.
- Use Python for orchestration, measurement, validation, and command-line interfaces.
- Use ExtendScript/JSX for native Photoshop document operations.
- Preserve editable PSD structure and never claim Photoshop integration tests passed unless they ran against a real Photoshop instance.

## Status

The repository is being bootstrapped. The first implementation milestone is the Photoshop automation bridge: COM connection, JSX execution, unit locking, structured errors, and diagnostics.

---
name: stk-automation
description: Use when creating or modifying STK scenarios through automation, connecting to STK by COM or MCP, generating external ephemeris, building transmitter/receiver links, or exporting STK simulation data.
---

# STK Automation

## Core Rule

Treat STK as the source of truth for exported simulation data. Scripts may generate inputs such as external ephemeris files, but exported position, velocity, acceleration, Doppler, and Doppler-rate tables must come from STK Data Providers or STK report APIs.

## Workflow

1. Clarify the scenario: STK version, object types, trajectory type, frequencies, output fields, output naming, and whether data must be grouped by object/frequency.
2. Prefer external ephemeris for smooth, custom trajectories. Avoid multi-waypoint Great Arc when Doppler-rate quality matters; segment boundaries can create velocity/acceleration discontinuities.
3. Create objects by automation: scenario, aircraft, receiver, ground facility/radiator, B1/L1 transmitters, chains/access.
4. Export from STK Data Providers:
   - Aircraft state: `Cartesian Position`, `Cartesian Velocity`, `Cartesian Acceleration`, usually `Fixed`.
   - Link data: `Link Information -> Freq. Doppler Shift`.
   - Doppler rate: create a VGT scalar from `Link Information / Freq. Doppler Shift`, then read `Scalar Calculations -> <scalar> -> Scalar Rate`.
5. Name outputs so the group, trajectory, and frequency are obvious, e.g. `G02_s_curve_L1_uav_receiver_link.csv`.
6. Validate before declaring success:
   - B1/L1 row counts and timestamps match.
   - No blank values or `NaN`.
   - Speed and altitude ranges match the requested envelope.
   - Acceleration is finite and physically plausible.
   - L1/B1 Doppler ratio is close to the carrier-frequency ratio.
   - Doppler-rate agrees with a finite-difference sanity check away from endpoints/intentional maneuvers.

## MCP Pattern

This skill is self-contained for other agents: the MCP package is bundled at `scripts/stkMCP`. To install it on a Windows machine with STK:

```powershell
cd <installed-skill-folder>\scripts\stkMCP
python -m pip install -e .
```

Then configure the MCP client to run:

```json
{
  "mcpServers": {
    "stkMCP": {
      "command": "python",
      "args": ["-m", "stk_mcp.server"],
      "cwd": "<installed-skill-folder>\\scripts\\stkMCP"
    }
  }
}
```

For a fuller copy-to-another-agent checklist, read `references/portable-install.md`.

Use `stkMCP` when available. It exposes:

- `stk_status`: checks Python/STK COM dependency availability.
- `get_default_uav_link_config`: returns a portable scenario config.
- `create_uav_link_project`: creates the STK project, generates external ephemeris, builds links, exports CSV, and returns validation results.

For local Python automation without MCP, use the same package core:

```python
from stk_mcp.automation import run_automation, default_config

config = default_config()
result = run_automation(config=config, base_dir=".")
```

## Common Pitfalls

- Do not compute final Doppler or state columns in Python and label them STK exports.
- Do not rely on visual smoothness alone; inspect speed, acceleration, and Doppler-rate continuity.
- Do not write logs to stdout in an MCP stdio server; stdout must remain valid MCP JSON-RPC traffic.
- Do not hard-code local project paths. Accept a config file, config JSON, or output base directory.

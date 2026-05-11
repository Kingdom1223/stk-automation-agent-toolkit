"""MCP server for STK automation.

Run with:
    python -m stk_mcp.server
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from .automation import default_config, load_config, run_automation


mcp = FastMCP("stkMCP")


@mcp.tool()
def stk_status() -> dict[str, Any]:
    """Check whether Python can import STK COM dependencies."""
    try:
        import win32com.client  # noqa: F401
    except Exception as exc:
        return {
            "ok": False,
            "reason": "pywin32 is not available",
            "detail": str(exc),
        }
    return {
        "ok": True,
        "reason": "pywin32 is available; STK connection is attempted by create_uav_link_project",
    }


@mcp.tool()
def get_default_uav_link_config() -> dict[str, Any]:
    """Return a portable default UAV link scenario config."""
    return default_config()


@mcp.tool()
def create_uav_link_project(
    config_json: str | None = None,
    config_path: str | None = None,
    base_dir: str | None = None,
) -> dict[str, Any]:
    """Create an STK UAV link scenario and export CSV data.

    Provide either config_json or config_path. If neither is provided, the
    bundled five-group default config is used. base_dir controls where outputs
    are written; the current working directory is used when omitted.
    """
    if config_json and config_path:
        raise ValueError("Use only one of config_json or config_path.")
    if config_json:
        config = json.loads(config_json)
    elif config_path:
        config = load_config(config_path)
    else:
        config = default_config()
    return run_automation(config=config, base_dir=base_dir or Path.cwd())


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

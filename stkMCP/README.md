# stkMCP

Portable MCP server for automating STK 11+ through Windows COM.

## Requirements

- Windows
- STK installed and licensed
- Python 3.10+
- `pywin32`
- `mcp`

Install:

```powershell
cd stkMCP
python -m pip install -e .
```

Run as an MCP stdio server:

```powershell
python -m stk_mcp.server
```

Example Claude Desktop style entry:

```json
{
  "mcpServers": {
    "stkMCP": {
      "command": "python",
      "args": ["-m", "stk_mcp.server"],
      "cwd": "C:\\path\\to\\stkMCP"
    }
  }
}
```

If this package is installed from the `stk-automation` skill, use the skill's
bundled path as `cwd`, for example:

```text
<skill-folder>\\scripts\\stkMCP
```

## Tools

- `stk_status`: check local Python dependency readiness.
- `get_default_uav_link_config`: return a portable default UAV-link scenario config.
- `create_uav_link_project`: create an STK project, generate external ephemeris, build transmitter/receiver links, export CSV, and return validation results.

`create_uav_link_project` accepts:

- `config_json`: JSON string with scenario settings.
- `config_path`: path to a JSON config file.
- `base_dir`: output base directory. If omitted, the server current working directory is used.

Use `examples/sample_uav_link_config.json` as a starting point. The server also exposes the full five-group default through `get_default_uav_link_config`.

## Data Policy

Final exported state and link columns are read from STK Data Providers. The Python code only generates input ephemeris and merges STK-exported tables.

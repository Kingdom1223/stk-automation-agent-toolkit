# Portable Installation

The `stk-automation` skill is designed to be copied to another agent as one folder.

## Skill-only install

Copy the whole folder:

```text
stk-automation/
```

into the target agent's skills directory. The bundled MCP package travels with it:

```text
stk-automation/scripts/stkMCP
```

## Install bundled MCP

On the target Windows machine:

```powershell
cd <installed-skill-folder>\scripts
.\install_stk_mcp.ps1
```

or manually:

```powershell
cd <installed-skill-folder>\scripts\stkMCP
python -m pip install -e .
```

## MCP client config

Use the installed skill path, not the original development path:

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

## Requirements

- Windows
- STK installed and licensed
- Python 3.10+
- Network or local wheel access for `mcp` and `pywin32`


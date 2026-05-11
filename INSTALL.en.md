# Installation and Publishing

## Local install

```powershell
git clone <repo-url>
cd stk-automation-agent-toolkit
cd stkMCP
python -m pip install -e .
```

Check dependencies:

```powershell
python -c "import win32com.client; print('pywin32 ok')"
python -c "import mcp; print('mcp ok')"
```

## Install into another agent

Copy:

```text
skills/stk-automation
```

to the target agent's skills directory, then install the bundled MCP package:

```powershell
cd <installed-skill-folder>\scripts
.\install_stk_mcp.ps1
```

## Publish to GitHub

If GitHub CLI is installed and authenticated:

```powershell
git init
git add .
git commit -m "Initial release: STK automation skill and MCP server"
gh repo create stk-automation-agent-toolkit --private --source . --remote origin --push
```

Use `--public` instead of `--private` if the repository should be public.


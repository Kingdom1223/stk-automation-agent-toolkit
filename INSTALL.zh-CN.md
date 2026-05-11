# 安装与发布说明

## 本地安装

```powershell
git clone <repo-url>
cd stk-automation-agent-toolkit
cd stkMCP
python -m pip install -e .
```

检查依赖：

```powershell
python -c "import win32com.client; print('pywin32 ok')"
python -c "import mcp; print('mcp ok')"
```

## 安装到其他 agent

复制：

```text
skills/stk-automation
```

到目标 agent 的 skills 目录，然后安装内置 MCP：

```powershell
cd <installed-skill-folder>\scripts
.\install_stk_mcp.ps1
```

## 发布 GitHub

如果已安装并登录 GitHub CLI：

```powershell
git init
git add .
git commit -m "Initial release: STK automation skill and MCP server"
gh repo create stk-automation-agent-toolkit --private --source . --remote origin --push
```

如果要公开仓库，把 `--private` 改为 `--public`。


# STK Automation Agent Toolkit / STK 自动化智能体工具包

[English](#english) | [中文](#中文)

---

## English

### Overview

`stk-automation-agent-toolkit` packages two portable components for agents that need to automate AGI/Ansys STK:

1. **`stk-automation` skill**: an installable agent skill that teaches an agent how to connect to STK, build scenarios, create transmitter/receiver links, export STK Data Provider tables, and validate simulation quality.
2. **`stkMCP` MCP server**: a Model Context Protocol server that exposes STK automation as tools over stdio.

The toolkit was designed for workflows such as:

- Creating STK scenarios automatically.
- Generating smooth UAV external ephemeris files.
- Adding aircraft receivers, ground transmitters, chains, and accesses.
- Exporting position, velocity, acceleration, Doppler shift, and Doppler-rate data.
- Producing one CSV per scenario group and carrier frequency.

### Repository Layout

```text
stk-automation-agent-toolkit/
├─ skills/
│  └─ stk-automation/
│     ├─ SKILL.md
│     ├─ agents/openai.yaml
│     ├─ references/
│     └─ scripts/
│        ├─ install_stk_mcp.ps1
│        └─ stkMCP/
└─ stkMCP/
   ├─ pyproject.toml
   ├─ README.md
   ├─ examples/sample_uav_link_config.json
   └─ stk_mcp/
```

The MCP package is included twice intentionally:

- `stkMCP/`: standalone MCP package.
- `skills/stk-automation/scripts/stkMCP/`: bundled MCP package that travels with the skill when installed into another agent.

### Tested Environment

| Item | Version / Requirement |
|---|---|
| OS | Windows 10/11 |
| STK | STK 11 tested; STK 12 is expected to work for the same COM Object Model paths, but should be validated locally |
| Python | Python 3.10+ |
| Python packages | `mcp`, `pywin32` |
| STK license | STK desktop license with Communications capability for transmitter/receiver link reports |
| Transport | MCP stdio |

Useful links:

- STK / Ansys official site: https://www.ansys.com/products/missions/ansys-stk
- MCP official site: https://modelcontextprotocol.io/
- MCP Python SDK: https://github.com/modelcontextprotocol/python-sdk
- Python: https://www.python.org/downloads/windows/
- Git: https://git-scm.com/download/win
- GitHub CLI: https://cli.github.com/

### Install the Skill

Copy the skill folder into the target agent's skills directory:

```powershell
Copy-Item -Recurse .\skills\stk-automation <TARGET_AGENT_SKILLS_DIR>\stk-automation
```

Then install the bundled MCP server:

```powershell
cd <TARGET_AGENT_SKILLS_DIR>\stk-automation\scripts
.\install_stk_mcp.ps1
```

Manual install:

```powershell
cd <TARGET_AGENT_SKILLS_DIR>\stk-automation\scripts\stkMCP
python -m pip install -e .
```

### Install Standalone stkMCP

```powershell
cd stkMCP
python -m pip install -e .
```

Check dependencies:

```powershell
python -c "import win32com.client; print('pywin32 ok')"
python -c "import mcp; print('mcp ok')"
```

### MCP Client Configuration

Use the standalone MCP package:

```json
{
  "mcpServers": {
    "stkMCP": {
      "command": "python",
      "args": ["-m", "stk_mcp.server"],
      "cwd": "C:\\path\\to\\stk-automation-agent-toolkit\\stkMCP"
    }
  }
}
```

Use the MCP package bundled inside the skill:

```json
{
  "mcpServers": {
    "stkMCP": {
      "command": "python",
      "args": ["-m", "stk_mcp.server"],
      "cwd": "C:\\path\\to\\stk-automation\\scripts\\stkMCP"
    }
  }
}
```

### MCP Tools

`stkMCP` exposes:

- `stk_status`: checks whether local Python can import STK COM dependencies.
- `get_default_uav_link_config`: returns a portable five-group UAV link scenario configuration.
- `create_uav_link_project`: creates the STK project, generates external ephemeris, builds links, exports CSV files, and returns validation results.

### Example Usage

Run from an MCP-capable agent:

```text
Use stkMCP to create a UAV B1/L1 link project with the default five-group config.
Export one CSV per group and frequency.
```

Run the automation module directly:

```powershell
cd stkMCP
python - <<'PY'
from stk_mcp.automation import default_config, run_automation

config = default_config()
result = run_automation(config=config, base_dir=".")
print(result["scenario_path"])
print(result["validation"])
PY
```

Use a JSON config:

```powershell
cd stkMCP
python - <<'PY'
from stk_mcp.automation import load_config, run_automation

config = load_config("examples/sample_uav_link_config.json")
result = run_automation(config=config, base_dir=".")
print(result)
PY
```

### Output Files

By default, outputs are written to `stk_uav_outputs/`:

- STK scenario: `*.sc`
- STK object files: `*.ac`, `*.f`, `*.x`, `*.c`, `*.r`, etc.
- External ephemeris: `external_ephemeris/*.e`
- CSV files:
  - `<Group>_<trajectory>_B1_uav_receiver_link.csv`
  - `<Group>_<trajectory>_L1_uav_receiver_link.csv`
  - `<Group>_<trajectory>_all_frequencies_uav_receiver_link.csv`
  - `five_groups_uav_link_doppler_all.csv`

CSV columns:

```text
Group, TimeUTC, Band,
X_m, Y_m, Z_m,
Vx_mps, Vy_mps, Vz_mps,
Ax_mps2, Ay_mps2, Az_mps2,
Doppler_Hz, DopplerRate_Hzps
```

### Data Quality Policy

The Python code may generate input ephemeris, but final exported fields are read from STK Data Providers:

- `Cartesian Position`
- `Cartesian Velocity`
- `Cartesian Acceleration`
- `Link Information -> Freq. Doppler Shift`
- `Scalar Calculations -> Scalar Rate`

Validation checks include:

- Matching B1/L1 timestamps.
- No blank values or `NaN`.
- Speed and acceleration range checks.
- L1/B1 Doppler ratio sanity check.
- Doppler-rate finite-difference sanity check.

### Troubleshooting

If `win32com` is missing:

```powershell
python -m pip install pywin32
```

If `mcp` is missing:

```powershell
python -m pip install mcp
```

If STK cannot be launched:

- Confirm STK is installed.
- Confirm the STK license is available.
- Launch STK manually once and close it.
- Check whether the COM ProgID `STK11.Application` is available.

If MCP stdio does not start:

- Do not print logs to stdout in the server.
- Check the MCP client `cwd`.
- Run `python -m stk_mcp.server` from the configured `cwd`.

---

## 中文

### 项目简介

`stk-automation-agent-toolkit` 提供两个可移植组件，用于让智能体自动化操作 AGI/Ansys STK：

1. **`stk-automation` skill**：可安装到 agent 的技能，指导 agent 连接 STK、自动建工程、建立发射机/接收机链路、导出 STK Data Provider 数据并做质量检查。
2. **`stkMCP` MCP server**：基于 Model Context Protocol 的本地 STK 自动化服务，通过 stdio 暴露工具。

适合以下任务：

- 自动创建 STK 工程。
- 生成平滑无人机外部星历轨迹。
- 创建无人机接收机、地面辐射源发射机、链路和 access。
- 导出位置、速度、加速度、多普勒频率、多普勒频率变化率。
- 按组别和频点分别输出 CSV。

### 仓库结构

```text
stk-automation-agent-toolkit/
├─ skills/
│  └─ stk-automation/
│     ├─ SKILL.md
│     ├─ agents/openai.yaml
│     ├─ references/
│     └─ scripts/
│        ├─ install_stk_mcp.ps1
│        └─ stkMCP/
└─ stkMCP/
   ├─ pyproject.toml
   ├─ README.md
   ├─ examples/sample_uav_link_config.json
   └─ stk_mcp/
```

这里有两份 MCP 是有意设计：

- `stkMCP/`：独立 MCP 包，可单独安装。
- `skills/stk-automation/scripts/stkMCP/`：随 skill 一起分发，安装到其他 agent 后也能直接安装 MCP。

### 测试环境与要求

| 项目 | 版本 / 要求 |
|---|---|
| 操作系统 | Windows 10/11 |
| STK | 已在 STK 11 测试；STK 12 理论上可用，但建议本地验证 |
| Python | Python 3.10+ |
| Python 依赖 | `mcp`, `pywin32` |
| STK 许可 | 需要 STK 桌面许可；发射机/接收机链路报告建议具备 Communications 能力 |
| MCP 传输 | stdio |

常用网址：

- STK / Ansys 官方网站：https://www.ansys.com/products/missions/ansys-stk
- MCP 官方网站：https://modelcontextprotocol.io/
- MCP Python SDK：https://github.com/modelcontextprotocol/python-sdk
- Python 下载：https://www.python.org/downloads/windows/
- Git 下载：https://git-scm.com/download/win
- GitHub CLI：https://cli.github.com/

### 安装 Skill

把 skill 文件夹复制到目标 agent 的 skills 目录：

```powershell
Copy-Item -Recurse .\skills\stk-automation <TARGET_AGENT_SKILLS_DIR>\stk-automation
```

安装 skill 内置 MCP：

```powershell
cd <TARGET_AGENT_SKILLS_DIR>\stk-automation\scripts
.\install_stk_mcp.ps1
```

也可以手动安装：

```powershell
cd <TARGET_AGENT_SKILLS_DIR>\stk-automation\scripts\stkMCP
python -m pip install -e .
```

### 单独安装 stkMCP

```powershell
cd stkMCP
python -m pip install -e .
```

检查依赖：

```powershell
python -c "import win32com.client; print('pywin32 ok')"
python -c "import mcp; print('mcp ok')"
```

### MCP 客户端配置

使用独立 MCP 包：

```json
{
  "mcpServers": {
    "stkMCP": {
      "command": "python",
      "args": ["-m", "stk_mcp.server"],
      "cwd": "C:\\path\\to\\stk-automation-agent-toolkit\\stkMCP"
    }
  }
}
```

使用 skill 内置 MCP 包：

```json
{
  "mcpServers": {
    "stkMCP": {
      "command": "python",
      "args": ["-m", "stk_mcp.server"],
      "cwd": "C:\\path\\to\\stk-automation\\scripts\\stkMCP"
    }
  }
}
```

### MCP 工具

`stkMCP` 暴露三个工具：

- `stk_status`：检查本地 Python 是否能导入 STK COM 依赖。
- `get_default_uav_link_config`：返回默认五组无人机链路仿真配置。
- `create_uav_link_project`：创建 STK 工程、生成外部星历、建立链路、导出 CSV，并返回验证结果。

### 使用示例

在支持 MCP 的 agent 中：

```text
使用 stkMCP 创建默认五组 UAV B1/L1 链路工程。
每组每个频点单独导出一个 CSV。
```

直接调用 Python 自动化模块：

```powershell
cd stkMCP
python - <<'PY'
from stk_mcp.automation import default_config, run_automation

config = default_config()
result = run_automation(config=config, base_dir=".")
print(result["scenario_path"])
print(result["validation"])
PY
```

使用 JSON 配置：

```powershell
cd stkMCP
python - <<'PY'
from stk_mcp.automation import load_config, run_automation

config = load_config("examples/sample_uav_link_config.json")
result = run_automation(config=config, base_dir=".")
print(result)
PY
```

### 输出文件

默认输出目录是 `stk_uav_outputs/`：

- STK 工程：`*.sc`
- STK 对象文件：`*.ac`, `*.f`, `*.x`, `*.c`, `*.r` 等
- 外部星历：`external_ephemeris/*.e`
- CSV 文件：
  - `<Group>_<trajectory>_B1_uav_receiver_link.csv`
  - `<Group>_<trajectory>_L1_uav_receiver_link.csv`
  - `<Group>_<trajectory>_all_frequencies_uav_receiver_link.csv`
  - `five_groups_uav_link_doppler_all.csv`

CSV 字段：

```text
Group, TimeUTC, Band,
X_m, Y_m, Z_m,
Vx_mps, Vy_mps, Vz_mps,
Ax_mps2, Ay_mps2, Az_mps2,
Doppler_Hz, DopplerRate_Hzps
```

### 数据质量原则

Python 可以生成输入星历，但最终导出的数据必须来自 STK Data Providers：

- `Cartesian Position`
- `Cartesian Velocity`
- `Cartesian Acceleration`
- `Link Information -> Freq. Doppler Shift`
- `Scalar Calculations -> Scalar Rate`

自动化结果会检查：

- B1/L1 时间戳是否一致。
- 是否存在空值或 `NaN`。
- 速度和加速度范围。
- L1/B1 多普勒比例是否接近载频比例。
- 多普勒变化率与有限差分结果是否一致。

### 常见问题

缺少 `win32com`：

```powershell
python -m pip install pywin32
```

缺少 `mcp`：

```powershell
python -m pip install mcp
```

STK 启动失败：

- 确认 STK 已安装。
- 确认 STK 许可可用。
- 手动启动一次 STK 后关闭。
- 检查 COM ProgID `STK11.Application` 是否可用。

MCP stdio 启动失败：

- MCP server 不要向 stdout 输出普通日志。
- 检查 MCP 客户端配置中的 `cwd`。
- 在配置的 `cwd` 下手动运行 `python -m stk_mcp.server`。


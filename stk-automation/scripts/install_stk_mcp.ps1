param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$SkillDir = Split-Path -Parent $PSScriptRoot
$McpDir = Join-Path $PSScriptRoot "stkMCP"

if (-not (Test-Path -LiteralPath $McpDir)) {
    throw "Missing bundled MCP directory: $McpDir"
}

Push-Location $McpDir
try {
    & $Python -m pip install -e .
}
finally {
    Pop-Location
}

Write-Output "stkMCP installed from $McpDir"
Write-Output "Use this MCP cwd in your client config: $McpDir"

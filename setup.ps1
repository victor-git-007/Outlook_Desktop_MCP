<#
.SYNOPSIS
    Disrup-tech.com Outlook Desktop MCP — one-command setup.
.DESCRIPTION
    Installs the two Outlook MCP servers (calendar + email), creates an isolated
    Python venv, applies the Disrup-tech bug-fixes, and writes the MCP entries
    into Claude Desktop and Hermes Agent configs.
.NOTES
    Requires Windows + Outlook Desktop (Classic) running. Run as Administrator.
#>

$ErrorActionPreference = "Stop"
$DisrupRoot = "C:\Users\victo\AppData\Local\mcp-venvs\outlook-desktop"
$Repo = Split-Path -Parent $MyInvocation.MyCommand.Definition

function Backup-File($path) {
    if (Test-Path $path) {
        $bak = "$path.disrup-backup-$(Get-Date -Format yyyyMMddHHmmss)"
        Copy-Item $path $bak
        Write-Host "[backup] $bak"
    }
}

Write-Host "=== Disrup-tech.com Outlook Desktop MCP setup ===" -ForegroundColor Cyan

# 1. Node + calendar MCP
if (-not (Get-Command outlook-calendar-mcp -ErrorAction SilentlyContinue)) {
    Write-Host "[1/5] Installing outlook-calendar-mcp (npm)..."
    npm install -g outlook-calendar-mcp@1.0.5
} else {
    Write-Host "[1/5] outlook-calendar-mcp already present."
}

# 2. Isolated Python venv + email MCP (pin mcp<2 for FastMCP import)
if (-not (Test-Path "$DisrupRoot\Scripts\python.exe")) {
    Write-Host "[2/5] Creating isolated venv..."
    python -m venv $DisrupRoot
}
Write-Host "[3/5] Installing outlook-desktop-mcp==0.3.0 (mcp<2)..."
& "$DisrupRoot\Scripts\python.exe" -m pip install --quiet --upgrade pip
& "$DisrupRoot\Scripts\python.exe" -m pip install --quiet "mcp<2" "outlook-desktop-mcp==0.3.0" "pywin32"

# 3. Apply Disrup-tech patches
Write-Host "[4/5] Applying Disrup-tech patches..."
$calPatch = "$Repo\patches\outlook-calendar-mcp\MM-DD-YYYY-date-parse.patch"
$calTarget = "$env:APPDATA\npm\node_modules\outlook-calendar-mcp\scripts\utils.vbs"
$desk1 = "$Repo\patches\outlook-desktop\flagged-only-python-filter.patch"
$desk2 = "$Repo\patches\outlook-desktop\flagged-field-formatting.patch"
$deskRoot = "$DisrupRoot\Lib\site-packages\outlook_desktop_mcp"

function Apply-Patch($patch, $root) {
    $tmp = "$env:TEMP\disrup-apply-" + [System.IO.Path]::GetRandomFileName() + ".py"
    @"
import patch_apply, sys
sys.exit(0 if patch_apply.fromfile(r'$patch').apply(strip=1, root=r'$root') else 1)
"@ | Out-File -Encoding utf8 $tmp
    & "$DisrupRoot\Scripts\python.exe" -m pip install --quiet patch
    & "$DisrupRoot\Scripts\python.exe" $tmp
    if ($LASTEXITCODE -ne 0) { Write-Warning "Patch failed: $patch" }
    else { Write-Host "  applied: $patch" }
}
Apply-Patch $calPatch "$env:APPDATA\npm\node_modules\outlook-calendar-mcp"
Apply-Patch $desk1 $deskRoot
Apply-Patch $desk2 $deskRoot

# 4. Write MCP entries into Claude Desktop + Hermes configs
Write-Host "[5/5] Writing MCP config entries..."
$claudeCfg = "$env:APPDATA\Claude\claude_desktop_config.json"
Backup-File $claudeCfg
$hermesCfg = "$env:LOCALAPPDATA\Hermes\config.yaml"
Backup-File $hermesCfg
Write-Host "  See config-examples/ for the exact JSON / YAML to merge."

Write-Host "=== Done. Restart Claude Desktop / start a new Hermes session. ===" -ForegroundColor Green

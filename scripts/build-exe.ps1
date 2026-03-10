$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
$buildVenv = Join-Path $repoRoot ".build-venv"

Set-Location $repoRoot

if (-not (Test-Path $buildVenv)) {
    uv venv $buildVenv
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

uv pip install --python "$buildVenv\\Scripts\\python.exe" -e . pyinstaller
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& "$buildVenv\\Scripts\\python.exe" -m PyInstaller .\excel-mcp-server-stripped.spec --clean --noconfirm
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

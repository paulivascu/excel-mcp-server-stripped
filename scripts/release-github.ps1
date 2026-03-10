param(
    [string]$Version,
    [string]$Notes,
    [switch]$Draft,
    [switch]$SkipBuild,
    [switch]$IncludePythonArtifacts
)

$ErrorActionPreference = "Stop"

function Fail([string]$Message) {
    Write-Error $Message
    exit 1
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir

Set-Location $repoRoot

if (-not $Version) {
    $versionLine = Select-String -Path ".\pyproject.toml" -Pattern '^version = "(.+)"$' | Select-Object -First 1
    if (-not $versionLine) {
        Fail "Could not determine version from pyproject.toml."
    }

    $Version = $versionLine.Matches[0].Groups[1].Value
}

$normalizedVersion = if ($Version.StartsWith("v")) { $Version.Substring(1) } else { $Version }
$tag = if ($Version.StartsWith("v")) { $Version } else { "v$Version" }

$gitStatus = git status --porcelain=v1
if ($LASTEXITCODE -ne 0) {
    Fail "Failed to read git status."
}

if ($gitStatus) {
    Fail "Working tree is not clean. Commit or stash changes before creating a release."
}

gh auth status *> $null
if ($LASTEXITCODE -ne 0) {
    Fail "GitHub CLI is not authenticated. Run 'gh auth login' first."
}

$currentBranch = git branch --show-current
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($currentBranch)) {
    Fail "Could not determine the current branch."
}

if (-not $SkipBuild) {
    & powershell -ExecutionPolicy Bypass -File ".\scripts\build-exe.ps1"
    if ($LASTEXITCODE -ne 0) {
        Fail "Executable build failed."
    }

    uv build
    if ($LASTEXITCODE -ne 0) {
        Fail "Python package build failed."
    }
}

$exePath = Join-Path $repoRoot "dist\excel-mcp-server-stripped.exe"
if (-not (Test-Path $exePath)) {
    Fail "Missing executable artifact: $exePath"
}

$assets = @($exePath)

if ($IncludePythonArtifacts) {
    $wheelPath = Join-Path $repoRoot "dist\excel_mcp_server_stripped-$normalizedVersion-py3-none-any.whl"
    $sdistPath = Join-Path $repoRoot "dist\excel_mcp_server_stripped-$normalizedVersion.tar.gz"

    if (-not (Test-Path $wheelPath)) {
        Fail "Missing wheel artifact: $wheelPath"
    }

    if (-not (Test-Path $sdistPath)) {
        Fail "Missing source distribution artifact: $sdistPath"
    }

    $assets += $wheelPath
    $assets += $sdistPath
}

git push origin HEAD
if ($LASTEXITCODE -ne 0) {
    Fail "Failed to push the current branch to origin."
}

$localTag = git tag --list $tag
if ($LASTEXITCODE -ne 0) {
    Fail "Failed to query existing tags."
}

if (-not $localTag) {
    git tag $tag
    if ($LASTEXITCODE -ne 0) {
        Fail "Failed to create local tag $tag."
    }
}

git push origin $tag
if ($LASTEXITCODE -ne 0) {
    Fail "Failed to push tag $tag to origin."
}

cmd /c "gh release view $tag >nul 2>&1"
if ($LASTEXITCODE -eq 0) {
    Fail "A GitHub release for $tag already exists."
}

$ghArgs = @("release", "create", $tag)
$ghArgs += $assets
$ghArgs += @("--title", $tag)

if ($Draft) {
    $ghArgs += "--draft"
}

if ($Notes) {
    $ghArgs += @("--notes", $Notes)
} else {
    $ghArgs += "--generate-notes"
}

& gh @ghArgs
if ($LASTEXITCODE -ne 0) {
    Fail "Failed to create GitHub release $tag."
}

Write-Host "Created GitHub release $tag."
Write-Host "Attached assets:"
$assets | ForEach-Object { Write-Host " - $_" }

if (-not $Draft) {
    Write-Host "Publishing this release also triggers the PyPI workflow in .github/workflows/publish.yml."
}

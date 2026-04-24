param(
    [string]$InstallPackagePath = "",
    [string]$UpdatePackagePath = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$labRoot = Join-Path $repoRoot "artifacts\ribbon-feed-test-lab"
$transferRoot = Join-Path $repoRoot "artifacts\transfer"
$transferZipPath = Join-Path $transferRoot "AJTools-Ribbon-Feed-TestLab.zip"
$feedRoot = Join-Path $labRoot "local-feed"

function Resolve-FirstExistingPath {
    param([string[]]$Candidates)

    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        $fullPath = [System.IO.Path]::GetFullPath($candidate)
        if (Test-Path $fullPath) {
            return $fullPath
        }
    }

    return $null
}

$InstallPackagePath = Resolve-FirstExistingPath @(
    $InstallPackagePath,
    (Join-Path $repoRoot "artifacts\runtime-package\Release\AJTools-3.0.0.3.zip"),
    (Join-Path $repoRoot "artifacts\update-test-lab\2-New\AJTools-3.0.0.3.zip"),
    (Join-Path $repoRoot "artifacts\runtime-package\Debug\AJTools-3.0.0.3.zip")
)

$UpdatePackagePath = Resolve-FirstExistingPath @(
    $UpdatePackagePath,
    (Join-Path $repoRoot "artifacts\runtime-package\Release\AJTools-3.2.0.0.zip")
)

if ([string]::IsNullOrWhiteSpace($InstallPackagePath) -or -not (Test-Path $InstallPackagePath)) {
    throw "Install package not found: $InstallPackagePath"
}

if ([string]::IsNullOrWhiteSpace($UpdatePackagePath) -or -not (Test-Path $UpdatePackagePath)) {
    throw "Update package not found: $UpdatePackagePath"
}

function Reset-Directory {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }

    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Get-VersionFromZipName {
    param([string]$ZipPath)

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ZipPath)
    if ($baseName -match '^AJTools-(\d+\.\d+\.\d+\.\d+)$') {
        return $Matches[1]
    }

    throw "Could not determine the AJ Tools version from ZIP name: $ZipPath"
}

Reset-Directory $labRoot
Reset-Directory $feedRoot

$installVersion = Get-VersionFromZipName -ZipPath $InstallPackagePath
$updateVersion = Get-VersionFromZipName -ZipPath $UpdatePackagePath

$installZipName = [System.IO.Path]::GetFileName($InstallPackagePath)
$updateZipName = [System.IO.Path]::GetFileName($UpdatePackagePath)

Copy-Item -LiteralPath $InstallPackagePath -Destination (Join-Path $labRoot $installZipName) -Force
Copy-Item -LiteralPath $UpdatePackagePath -Destination (Join-Path $labRoot $updateZipName) -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "scripts\Install-AJTools.ps1") -Destination (Join-Path $labRoot "Install-AJTools.ps1") -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "scripts\Uninstall-AJTools.ps1") -Destination (Join-Path $labRoot "Uninstall-AJTools.ps1") -Force

$manifest = [ordered]@{
    version         = $updateVersion
    buildDate       = (Get-Date).ToString("yyyy-MM-dd")
    downloadZipFile = "..\$updateZipName"
    downloadUrl     = "..\$updateZipName"
    releaseNotes    = "Ribbon-driven AJ Tools update test."
    releaseNotesUrl = ""
    sha256          = (Get-FileHash -Path $UpdatePackagePath -Algorithm SHA256).Hash
}

$manifest | ConvertTo-Json | Set-Content -Path (Join-Path $feedRoot "version.json") -Encoding UTF8

@"
@echo off
setlocal
echo Installing AJ Tools $installVersion and wiring the ribbon to a local update feed.
echo UAC should not appear for this user-scoped install path.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-AJTools.ps1" -PackageZip "%~dp0$installZipName" -UpdateFeedUrl "%~dp0local-feed\version.json"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
    echo AJ Tools install failed with exit code %EXIT_CODE%.
) else (
    echo AJ Tools install completed.
)
pause
exit /b %EXIT_CODE%
"@ | Set-Content -Path (Join-Path $labRoot "1-Install AJ Tools $installVersion.cmd") -Encoding ASCII

@'
@echo off
setlocal
echo Remove AJ Tools by running the user-scoped uninstall.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-AJTools.ps1"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
    echo AJ Tools uninstall failed with exit code %EXIT_CODE%.
) else (
    echo AJ Tools uninstall completed.
)
pause
exit /b %EXIT_CODE%
'@ | Set-Content -Path (Join-Path $labRoot "3-Uninstall AJ Tools.cmd") -Encoding ASCII

@"
Ribbon update test flow

1. Double-click "1-Install AJ Tools $installVersion.cmd"
2. Open Microsoft Project and confirm the AJ ribbon loads.
3. In the AJ ribbon, click "Check for Updates".
4. The ribbon should discover version $updateVersion from local-feed\version.json.
5. Accept the update and note whether UAC appears.
6. Reopen Microsoft Project and confirm the AJ ribbon still loads.
7. Double-click "3-Uninstall AJ Tools.cmd" when finished.

What this lab is testing
- The ribbon-driven update check, not the manual updater launcher
- A local feed file that points to a local ZIP payload
- Update behavior from $installVersion to $updateVersion without publishing to GitHub first
"@ | Set-Content -Path (Join-Path $labRoot "README.txt") -Encoding ASCII

if (-not (Test-Path $transferRoot)) {
    New-Item -ItemType Directory -Path $transferRoot -Force | Out-Null
}

if (Test-Path $transferZipPath) {
    Remove-Item -LiteralPath $transferZipPath -Force
}

Compress-Archive -Path (Join-Path $labRoot '*') -DestinationPath $transferZipPath

Write-Host ""
Write-Host "Ribbon feed test lab ready:"
Write-Host $labRoot
Write-Host ""
Write-Host "Transfer bundle ready:"
Write-Host $transferZipPath

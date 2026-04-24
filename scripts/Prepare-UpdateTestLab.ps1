param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$testLabRoot = Join-Path $repoRoot "artifacts\update-test-lab"
$transferRoot = Join-Path $repoRoot "artifacts\transfer"
$legacySource = Join-Path $repoRoot "artifacts\legacy-package\Release"
$newSource = Join-Path $repoRoot "artifacts\runtime-package\$Configuration"

if (-not (Test-Path $legacySource)) {
    throw "Legacy package source not found: $legacySource"
}

if (-not (Test-Path $newSource)) {
    throw "New runtime package source not found: $newSource"
}

function Reset-Directory {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }

    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Copy-Tree {
    param(
        [string]$SourceRoot,
        [string]$DestinationRoot
    )

    Get-ChildItem -LiteralPath $SourceRoot -Force | ForEach-Object {
        Copy-Item $_.FullName $DestinationRoot -Recurse -Force
    }
}

function Get-VersionFromZipName {
    param([string]$ZipPath)

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ZipPath)
    if ($baseName -match '^AJTools-(\d+\.\d+\.\d+\.\d+)$') {
        return $Matches[1]
    }

    throw "Could not determine the AJ Tools version from ZIP name: $ZipPath"
}

function Get-IncrementedVersion {
    param([string]$Version)

    $parts = $Version.Split('.')
    if ($parts.Length -ne 4) {
        throw "Expected a four-part version number but got: $Version"
    }

    $parts[3] = ([int]$parts[3] + 1).ToString()
    return ($parts -join '.')
}

Reset-Directory $testLabRoot
$legacyRoot = Join-Path $testLabRoot "1-Legacy"
$newRoot = Join-Path $testLabRoot "2-New"
Reset-Directory $legacyRoot
Reset-Directory $newRoot

Copy-Tree -SourceRoot $legacySource -DestinationRoot $legacyRoot
Copy-Tree -SourceRoot $newSource -DestinationRoot $newRoot

$installZip = Get-ChildItem -LiteralPath $newRoot -Filter 'AJTools-*.zip' | Select-Object -First 1
if ($null -eq $installZip) {
    throw "The new runtime test bundle does not contain an AJTools-*.zip package."
}

$installVersion = Get-VersionFromZipName -ZipPath $installZip.FullName
$updateVersion = Get-IncrementedVersion -Version $installVersion
$updateZipName = "AJTools-$updateVersion.zip"
$updateZipPath = Join-Path $newRoot $updateZipName
Copy-Item -LiteralPath $installZip.FullName -Destination $updateZipPath -Force
$runNewUpdatePath = Join-Path $newRoot "Run-NewUpdate.ps1"

$labReadmePath = Join-Path $testLabRoot "README.txt"
$transferZipPath = Join-Path $transferRoot "AJTools-Update-TestLab.zip"

@'
@echo off
setlocal
echo Launching legacy AJ Tools installer...
echo If UAC appears here, that is expected for the legacy install path.
echo.
start "" "%~dp0AJSetup.exe"
exit /b 0
'@ | Set-Content -Path (Join-Path $legacyRoot "1-Install Legacy AJ Tools.cmd") -Encoding ASCII

@'
@echo off
setlocal
echo Launching legacy AJ Tools update path...
echo If UAC appears here, that is the exact behavior we are testing.
echo.
start "" "%~dp0AJSetup.exe" /update "%~dp0AJAddIn.msi" /version "LegacyTest"
exit /b 0
'@ | Set-Content -Path (Join-Path $legacyRoot "2-Update Legacy AJ Tools.cmd") -Encoding ASCII

@'
@echo off
setlocal
echo Removing legacy AJ Tools usually needs to happen from Windows Installed Apps.
echo Open Settings ^> Apps ^> Installed apps and uninstall AJ Tools there.
pause
exit /b 0
'@ | Set-Content -Path (Join-Path $legacyRoot "3-Uninstall Legacy AJ Tools.cmd") -Encoding ASCII

@'
Legacy AJ Tools test flow

1. Double-click "1-Install Legacy AJ Tools.cmd"
2. Note whether UAC appears
3. Open Microsoft Project and verify the AJ ribbon loads
4. Close Microsoft Project
5. Double-click "2-Update Legacy AJ Tools.cmd"
6. Note whether UAC appears during the update
7. Reopen Microsoft Project and verify the AJ ribbon still loads
8. Use "3-Uninstall Legacy AJ Tools.cmd" for uninstall guidance
'@ | Set-Content -Path (Join-Path $legacyRoot "README.txt") -Encoding ASCII

@"
@echo off
setlocal
echo Launching new AJ Tools install...
echo UAC should not appear for this user-scoped install path.
echo.
call "%~dp0Install-AJTools.cmd" "%~dp0$($installZip.Name)"
exit /b %ERRORLEVEL%
"@ | Set-Content -Path (Join-Path $newRoot "1-Install New AJ Tools.cmd") -Encoding ASCII

@"
`$ErrorActionPreference = 'Stop'
`$root = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$zip = Join-Path `$root '$updateZipName'
if (-not (Test-Path `$zip)) {
    throw 'The update ZIP was not found: ' + `$zip
}

`$updater = Get-ChildItem -LiteralPath `$root -Recurse -Filter 'AJRuntimeUpdater.exe' | Select-Object -First 1
if (-not `$updater) {
    throw 'AJRuntimeUpdater.exe was not found in the new test bundle.'
}

& `$updater.FullName /version '$updateVersion' /zip `$zip
"@ | Set-Content -Path $runNewUpdatePath -Encoding ASCII

@'
@echo off
setlocal
echo Launching new AJ Tools update path...
echo UAC should not appear for this runtime updater path.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-NewUpdate.ps1"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
    echo New AJ Tools update failed with exit code %EXIT_CODE%.
)
pause
exit /b %EXIT_CODE%
'@ | Set-Content -Path (Join-Path $newRoot "2-Update New AJ Tools.cmd") -Encoding ASCII

@'
@echo off
setlocal
echo Launching new AJ Tools uninstall...
echo.
call "%~dp0Uninstall-AJTools.cmd"
exit /b %ERRORLEVEL%
'@ | Set-Content -Path (Join-Path $newRoot "3-Uninstall New AJ Tools.cmd") -Encoding ASCII

@'
New AJ Tools test flow

1. Double-click "1-Install New AJ Tools.cmd"
2. Note whether UAC appears
3. Open Microsoft Project and verify the AJ ribbon loads
4. Close Microsoft Project
5. Double-click "2-Update New AJ Tools.cmd"
6. Note whether UAC appears during the update
7. Reopen Microsoft Project and verify the AJ ribbon still loads
8. Double-click "3-Uninstall New AJ Tools.cmd" to remove the user-scoped runtime
'@ | Set-Content -Path (Join-Path $newRoot "README.txt") -Encoding ASCII

@"
AJ Tools update comparison lab

Legacy folder:
- Install with 1-Install Legacy AJ Tools.cmd
- Test update with 2-Update Legacy AJ Tools.cmd
- Uninstall with 3-Uninstall Legacy AJ Tools.cmd

New folder:
- Install version $installVersion with 1-Install New AJ Tools.cmd
- Update to version $updateVersion with 2-Update New AJ Tools.cmd
- Uninstall with 3-Uninstall New AJ Tools.cmd

What to record:
- Whether UAC appears during install
- Whether UAC appears during update
- Which executable triggered UAC
- Whether Microsoft Project loads the AJ ribbon after install and after update
"@ | Set-Content -Path $labReadmePath -Encoding ASCII

if (-not (Test-Path $transferRoot)) {
    New-Item -ItemType Directory -Path $transferRoot -Force | Out-Null
}

if (Test-Path $transferZipPath) {
    Remove-Item -LiteralPath $transferZipPath -Force
}

Compress-Archive -Path (Join-Path $testLabRoot '*') -DestinationPath $transferZipPath

Write-Host ""
Write-Host "Update test lab ready:"
Write-Host $testLabRoot
Write-Host ""
Write-Host "Transfer bundle ready:"
Write-Host $transferZipPath

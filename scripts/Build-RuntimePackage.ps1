param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repoRoot "Arian Jahandarfards MS Project Add-in.slnx"
$addInProjectDir = Join-Path $repoRoot "Arian Jahandarfards MS Project Add-in"
$outputDir = Join-Path $addInProjectDir "bin\$Configuration"
$artifactRoot = Join-Path $repoRoot "artifacts\runtime-package\$Configuration"
$versionFile = Join-Path $addInProjectDir "Properties\AssemblyInfo.cs"

function Get-MSBuildPath {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $path = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        if ($path) { return $path }
    }

    $fallbacks = @(
        "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($candidate in $fallbacks) {
        if (Test-Path $candidate) { return $candidate }
    }

    throw "MSBuild.exe was not found."
}

function Get-VersionFromAssemblyInfo {
    $content = Get-Content $versionFile -Raw
    if ($content -match '\[assembly: AssemblyVersion\("(\d+\.\d+\.\d+\.\d+)"\)\]') {
        return $Matches[1]
    }

    throw "Could not read AssemblyVersion from $versionFile"
}

function Copy-DirectoryTree {
    param(
        [string]$SourceRoot,
        [string]$DestinationRoot
    )

    Get-ChildItem -Path $SourceRoot -Recurse -File | ForEach-Object {
        $relativePath = $_.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $destinationPath = Join-Path $DestinationRoot $relativePath
        $destinationDir = Split-Path -Parent $destinationPath
        if (-not (Test-Path $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
        }

        Copy-Item $_.FullName $destinationPath -Force
    }
}

$msbuild = Get-MSBuildPath
$version = Get-VersionFromAssemblyInfo
$normalizedVersion = $version.ToLowerInvariant()
$stagingDir = Join-Path $artifactRoot "AJTools_$($version.Replace('.', '_'))"
$zipPath = Join-Path $artifactRoot "AJTools-$normalizedVersion.zip"
$manifestPath = Join-Path $artifactRoot "version.json"
$installPs1Path = Join-Path $repoRoot "scripts\Install-AJTools.ps1"
$uninstallPs1Path = Join-Path $repoRoot "scripts\Uninstall-AJTools.ps1"
$installCmdPath = Join-Path $repoRoot "scripts\Install-AJTools.cmd"
$uninstallCmdPath = Join-Path $repoRoot "scripts\Uninstall-AJTools.cmd"

Write-Host "Building solution..."
& $msbuild $solutionPath /t:Rebuild /p:Configuration=$Configuration /p:Platform="Any CPU"
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed."
}

if (-not (Test-Path $outputDir)) {
    throw "Build output directory not found: $outputDir"
}

if (Test-Path $artifactRoot) {
    Remove-Item -LiteralPath $artifactRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null

$excludedFiles = @(
    "AJAddIn.msi",
    "AJSetup.exe",
    "AJSetup.exe.config",
    "AJSetup.pdb"
)

Get-ChildItem -Path $outputDir -Force | ForEach-Object {
    if ($excludedFiles -contains $_.Name) {
        return
    }

    if ($_.PSIsContainer) {
        Copy-DirectoryTree -SourceRoot $_.FullName -DestinationRoot (Join-Path $stagingDir $_.Name)
    }
    elseif ($_.Extension -ne ".pdb") {
        Copy-Item $_.FullName (Join-Path $stagingDir $_.Name) -Force
    }
}

Compress-Archive -Path (Join-Path $stagingDir '*') -DestinationPath $zipPath
$zipHash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash

$manifest = [ordered]@{
    version         = $version
    buildDate       = (Get-Date).ToString("yyyy-MM-dd")
    downloadZipFile = (Split-Path $zipPath -Leaf)
    downloadUrl     = (Split-Path $zipPath -Leaf)
    releaseNotes    = ""
    releaseNotesUrl = ""
    sha256          = $zipHash
}

$manifest | ConvertTo-Json | Set-Content -Path $manifestPath -Encoding UTF8

Copy-Item $installPs1Path (Join-Path $artifactRoot "Install-AJTools.ps1") -Force
Copy-Item $uninstallPs1Path (Join-Path $artifactRoot "Uninstall-AJTools.ps1") -Force
Copy-Item $installCmdPath (Join-Path $artifactRoot "Install-AJTools.cmd") -Force
Copy-Item $uninstallCmdPath (Join-Path $artifactRoot "Uninstall-AJTools.cmd") -Force

Write-Host ""
Write-Host "Runtime package ready:"
Get-Item $zipPath, $manifestPath | Select-Object FullName, LastWriteTime, Length | Format-Table -AutoSize

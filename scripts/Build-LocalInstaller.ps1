param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repoRoot "Arian Jahandarfards MS Project Add-in.slnx"
$addInProjectDir = Join-Path $repoRoot "Arian Jahandarfards MS Project Add-in"
$addInReleaseDir = Join-Path $addInProjectDir "bin\Release"
$setupReleaseDir = Join-Path $repoRoot "AJSetup\bin\Release"
$wxsPath = Join-Path $repoRoot "AJToolsInstaller\Package.wxs"
$setupExePath = Join-Path $setupReleaseDir "AJSetup.exe"
$logoPath = Join-Path $setupReleaseDir "AJ Logo Final Files-02.png"
$outputMsiPath = Join-Path $setupReleaseDir "AJAddIn.msi"
$versionFile = Join-Path $addInProjectDir "Properties\AssemblyInfo.cs"

function Get-MSBuildPath {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $path = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        if ($path) { return $path }
    }

    $fallbacks = @(
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($candidate in $fallbacks) {
        if (Test-Path $candidate) { return $candidate }
    }

    throw "MSBuild.exe was not found. Open this repo in Visual Studio or install the Visual Studio Build Tools workload."
}

function Get-WixPath {
    $cmd = Get-Command wix -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    throw "WiX v4 CLI ('wix') was not found on PATH. Install it with: dotnet tool install --global wix --version 4.0.5"
}

function Get-VersionFromAssemblyInfo {
    $content = Get-Content $versionFile -Raw
    if ($content -match '\[assembly: AssemblyVersion\("(\d+\.\d+\.\d+\.\d+)"\)\]') {
        return $Matches[1]
    }

    throw "Could not read AssemblyVersion from $versionFile"
}

$msbuild = Get-MSBuildPath
$wix = Get-WixPath
$version = Get-VersionFromAssemblyInfo

Write-Host "Building solution with MSBuild..."
& $msbuild $solutionPath /t:Rebuild /p:Configuration=$Configuration /p:Platform="Any CPU"
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed."
}

Write-Host "Refreshing MSI packaging inputs..."
Copy-Item $setupExePath (Join-Path $addInReleaseDir "AJSetup.exe") -Force
Copy-Item $logoPath (Join-Path $addInReleaseDir "AJ Logo Final Files-02.png") -Force

Write-Host "Building local MSI..."
& $wix build $wxsPath -o $outputMsiPath -d Version=$version -d SourceDir=$repoRoot
if ($LASTEXITCODE -ne 0) {
    throw "WiX build failed."
}

Write-Host ""
Write-Host "Local installer is ready:"
Get-Item $outputMsiPath, $setupExePath | Select-Object FullName, LastWriteTime, Length | Format-Table -AutoSize

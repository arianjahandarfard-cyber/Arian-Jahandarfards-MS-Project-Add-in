param(
    [string]$Configuration = "Debug",
    [switch]$Build
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$registerScript = Join-Path $PSScriptRoot "Register-DebugAddIn.ps1"

& $registerScript -Configuration $Configuration -Build:$Build

$projectRootPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office\16.0\Project\InstallRoot",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Project\InstallRoot"
)

$winprojPath = $null
foreach ($registryPath in $projectRootPaths) {
    try {
        $installRoot = (Get-ItemProperty -Path $registryPath -ErrorAction Stop).Path
        if ($installRoot) {
            $candidate = Join-Path $installRoot "WINPROJ.EXE"
            if (Test-Path $candidate) {
                $winprojPath = $candidate
                break
            }
        }
    }
    catch {
    }
}

if (-not $winprojPath) {
    $winprojPath = "WINPROJ.EXE"
}

Start-Process -FilePath $winprojPath

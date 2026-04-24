param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

Write-Warning "Build-LocalInstaller.ps1 is now a compatibility wrapper. AJ Tools updates are ZIP-based, so this script will build the runtime package instead of the legacy MSI installer."

& (Join-Path $PSScriptRoot "Build-RuntimePackage.ps1") -Configuration $Configuration

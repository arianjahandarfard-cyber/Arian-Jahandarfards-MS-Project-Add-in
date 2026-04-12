param(
    [string]$Configuration = "Debug",
    [switch]$Build
)

$ErrorActionPreference = "Stop"

function Ensure-Elevated {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        return
    }

    $args = @(
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-Configuration", $Configuration
    )

    if ($Build) {
        $args += "-Build"
    }

    Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $args
    exit
}

function Get-MSBuildPath {
    $fallbacks = @(
        "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($candidate in $fallbacks) {
        if (Test-Path $candidate) { return $candidate }
    }

    throw "MSBuild.exe was not found. Open the solution in Visual Studio and build it once, or install the Visual Studio Build Tools."
}

function Remove-AJProjectRegistryKeys {
    $projectKeys = @(
        "HKCU:\Software\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn",
        "HKCU:\Software\Microsoft\Office\MS Project\Addins\Arian Jahandarfards MS Project Add-in"
    )

    foreach ($key in $projectKeys) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Register-DebugProjectAddInKey {
    param(
        [string]$ManifestPath
    )

    $addInRoot = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\Microsoft\Office\MS Project\Addins\Arian Jahandarfards MS Project Add-in")
    if ($null -eq $addInRoot) {
        throw "Could not create the Project add-in registration key for debug mode."
    }

    try {
        $addInRoot.SetValue("Description", "Arian Jahandarfards MS Project Add-in", [Microsoft.Win32.RegistryValueKind]::String)
        $addInRoot.SetValue("FriendlyName", "Arian Jahandarfards MS Project Add-in", [Microsoft.Win32.RegistryValueKind]::String)
        $addInRoot.SetValue("LoadBehavior", 3, [Microsoft.Win32.RegistryValueKind]::DWord)
        $addInRoot.SetValue("Manifest", "$ManifestPath|vstolocal", [Microsoft.Win32.RegistryValueKind]::String)
    }
    finally {
        $addInRoot.Close()
    }
}

function Remove-AJCurrentUserUninstallEntries {
    $uninstallRoot = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall", $true)
    if ($null -eq $uninstallRoot) {
        return
    }

    try {
        foreach ($subKeyName in @($uninstallRoot.GetSubKeyNames())) {
            $subKey = $uninstallRoot.OpenSubKey($subKeyName)
            if ($null -eq $subKey) {
                continue
            }

            try {
                $displayName = $subKey.GetValue("DisplayName", "")
                if ($displayName -like "*Arian Jahandarfards MS Project Add-in*" -or $displayName -like "AJ Tools*") {
                    $uninstallRoot.DeleteSubKeyTree($subKeyName, $false)
                }
            }
            finally {
                $subKey.Close()
            }
        }
    }
    finally {
        $uninstallRoot.Close()
    }
}

function Remove-AJVstoMetadata {
    $rootPath = "Software\Microsoft\VSTO\SolutionMetadata"
    $registryRoot = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($rootPath, $true)
    if ($null -eq $registryRoot) {
        return
    }

    try {
        foreach ($valueName in @($registryRoot.GetValueNames())) {
            if ($valueName -like "*Arian Jahandarfards MS Project Add-in*") {
                $guidValue = $registryRoot.GetValue($valueName)
                $registryRoot.DeleteValue($valueName, $false)

                if ($guidValue -and $registryRoot.OpenSubKey($guidValue)) {
                    $registryRoot.DeleteSubKeyTree($guidValue, $false)
                }
            }
        }

        foreach ($subKeyName in @($registryRoot.GetSubKeyNames())) {
            $subKey = $registryRoot.OpenSubKey($subKeyName)
            if ($null -eq $subKey) {
                continue
            }

            try {
                $addInName = $subKey.GetValue("addInName", "")
                $friendlyName = $subKey.GetValue("friendlyName", "")
                $description = $subKey.GetValue("description", "")

                if ($addInName -eq "Arian Jahandarfards MS Project Add-in" -or
                    $friendlyName -eq "Arian Jahandarfards MS Project Add-in" -or
                    $description -eq "Arian Jahandarfards MS Project Add-in") {
                    $registryRoot.DeleteSubKeyTree($subKeyName, $false)
                }
            }
            finally {
                $subKey.Close()
            }
        }
    }
    finally {
        $registryRoot.Close()
    }
}

function Remove-AJVstoSecurity {
    $rootPath = "Software\Microsoft\VSTO\Security\Inclusion"
    $registryRoot = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($rootPath, $true)
    if ($null -eq $registryRoot) {
        return
    }

    try {
        foreach ($subKeyName in @($registryRoot.GetSubKeyNames())) {
            $subKey = $registryRoot.OpenSubKey($subKeyName)
            if ($null -eq $subKey) {
                continue
            }

            try {
                $url = $subKey.GetValue("Url", "")
                if ($url -like "*Arian Jahandarfards MS Project Add-in*") {
                    $registryRoot.DeleteSubKeyTree($subKeyName, $false)
                }
            }
            finally {
                $subKey.Close()
            }
        }
    }
    finally {
        $registryRoot.Close()
    }
}

Ensure-Elevated

$repoRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repoRoot "Arian Jahandarfards MS Project Add-in.slnx"
$projectDir = Join-Path $repoRoot "Arian Jahandarfards MS Project Add-in"
$outputDir = Join-Path $projectDir "bin\$Configuration"
$vstoPath = Join-Path $outputDir "Arian Jahandarfards MS Project Add-in.vsto"
$installedMachineKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn"
)

if (Get-Process WINPROJ -ErrorAction SilentlyContinue) {
    throw "Close Microsoft Project before switching into debug add-in mode."
}

if ($Build) {
    $msbuild = Get-MSBuildPath
    & $msbuild $solutionPath /t:Build /p:Configuration=$Configuration /p:Platform="Any CPU"
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed while preparing the debug add-in."
    }
}

if (-not (Test-Path $vstoPath)) {
    throw "The VSTO manifest was not found: $vstoPath"
}

foreach ($machineKey in $installedMachineKeys) {
    if (Test-Path $machineKey) {
        try {
            Set-ItemProperty -Path $machineKey -Name "LoadBehavior" -Type DWord -Value 0
        }
        catch {
            Write-Warning "Could not disable installed machine-wide add-in registration at $machineKey"
        }
    }
}

Remove-AJProjectRegistryKeys
Remove-AJVstoMetadata
Remove-AJVstoSecurity
Remove-AJCurrentUserUninstallEntries

Register-DebugProjectAddInKey -ManifestPath $vstoPath

Write-Host ""
Write-Host "Debug add-in registration is active."
Write-Host "Microsoft Project will now load the local $Configuration build from:"
Write-Host $vstoPath
Write-Host ""
Write-Host "Next steps:"
Write-Host "1. Start debugging from Visual Studio, or open Microsoft Project manually."
Write-Host "2. When you want to go back to the installed AJ Tools version, run scripts\\Use-InstalledAddIn.ps1."

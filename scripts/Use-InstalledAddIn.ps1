$ErrorActionPreference = "Stop"

function Ensure-Elevated {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        return
    }

    Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList @(
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`""
    )
    exit
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

function Remove-AJProjectAddInData {
    $addInDataKeys = @(
        "HKCU:\Software\Microsoft\Office\MS Project\AddinsData\ArianJahandarfardsAddIn",
        "HKCU:\Software\Microsoft\Office\MS Project\AddinsData\Arian Jahandarfards MS Project Add-in"
    )

    foreach ($key in $addInDataKeys) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
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

$installedManifest = "C:\Program Files (x86)\AJTools\Arian Jahandarfards MS Project Add-in.vsto|vstolocal"
$installedMachineKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn"
)

if (Get-Process WINPROJ -ErrorAction SilentlyContinue) {
    throw "Close Microsoft Project before switching back to the installed add-in."
}

Remove-AJProjectRegistryKeys
Remove-AJProjectAddInData
Remove-AJVstoMetadata
Remove-AJVstoSecurity
Remove-AJCurrentUserUninstallEntries

foreach ($machineKey in $installedMachineKeys) {
    try {
        if (-not (Test-Path $machineKey)) {
            New-Item -Path $machineKey -Force | Out-Null
        }

        Set-ItemProperty -Path $machineKey -Name "Description" -Type String -Value "AJ Tools"
        Set-ItemProperty -Path $machineKey -Name "FriendlyName" -Type String -Value "AJ Tools"
        Set-ItemProperty -Path $machineKey -Name "LoadBehavior" -Type DWord -Value 3
        Set-ItemProperty -Path $machineKey -Name "Manifest" -Type String -Value $installedManifest
    }
    catch {
        Write-Warning "Could not restore machine-wide add-in registration at $machineKey"
    }
}

Write-Host ""
Write-Host "Installed add-in mode is active again."
Write-Host "Microsoft Project will fall back to the AJ Tools version installed under Program Files on next launch."

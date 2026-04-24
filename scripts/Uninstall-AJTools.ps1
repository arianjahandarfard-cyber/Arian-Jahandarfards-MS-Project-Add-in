param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "AJTools")
)

$ErrorActionPreference = "Stop"

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

if (Get-Process WINPROJ -ErrorAction SilentlyContinue) {
    throw "Close Microsoft Project before uninstalling AJ Tools."
}

Remove-AJProjectRegistryKeys
Remove-AJProjectAddInData
Remove-AJVstoMetadata
Remove-AJVstoSecurity

if (Test-Path $InstallRoot) {
    Remove-Item -LiteralPath $InstallRoot -Recurse -Force
}

Write-Host ""
Write-Host "AJ Tools user-scoped runtime removed from: $InstallRoot"

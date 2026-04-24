param(
    [Parameter(Mandatory = $true)]
    [string]$PackageZip,
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "AJTools"),
    [string]$UpdateFeedUrl
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

function Remove-AJDeploymentCacheEntries {
    $deploymentRoot = "HKCU:\Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0"
    if (-not (Test-Path $deploymentRoot)) {
        return
    }

    Get-ChildItem -Path $deploymentRoot -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $item = Get-Item $_.PSPath -ErrorAction Stop
            foreach ($valueName in @($item.GetValueNames())) {
                $value = $item.GetValue($valueName)
                $text = if ($value -is [byte[]]) {
                    [System.Text.Encoding]::ASCII.GetString($value)
                }
                else {
                    [string]$value
                }

                if ($valueName -like "*Arian*" -or
                    $text -like "*Arian*Project*Add-in*") {
                    $item.DeleteValue($valueName, $false)
                }
            }
        }
        catch {
        }
    }
}

function Remove-AJClickOnceCacheFiles {
    $appsRoot = Join-Path $env:LOCALAPPDATA "Apps\2.0"
    if (-not (Test-Path $appsRoot)) {
        return
    }

    $targets = Get-ChildItem -Path $appsRoot -Recurse -Force -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -like "aria..vsto_*" -or
        $_.FullName -like "*Arian*Jahandarfards*MS*Project*Add-in*"
    }

    foreach ($item in $targets | Sort-Object { $_.FullName.Length } -Descending) {
        try {
            if ($item.PSIsContainer) {
                Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
            else {
                Remove-Item -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
        }
    }
}

function Convert-ManifestPathToVstoLocalUri {
    param(
        [string]$ManifestPath
    )

    return ([System.Uri]$ManifestPath).AbsoluteUri + "|vstolocal"
}

function Register-ProjectAddInKey {
    param(
        [string]$ManifestPath
    )

    $manifestValue = Convert-ManifestPathToVstoLocalUri -ManifestPath $ManifestPath

    $addInRoot = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\Microsoft\Office\MS Project\Addins\ArianJahandarfardsAddIn")
    if ($null -eq $addInRoot) {
        throw "Could not create the Project add-in registration key."
    }

    try {
        $addInRoot.SetValue("Description", "AJ Tools", [Microsoft.Win32.RegistryValueKind]::String)
        $addInRoot.SetValue("FriendlyName", "AJ Tools", [Microsoft.Win32.RegistryValueKind]::String)
        $addInRoot.SetValue("LoadBehavior", 3, [Microsoft.Win32.RegistryValueKind]::DWord)
        $addInRoot.SetValue("Manifest", $manifestValue, [Microsoft.Win32.RegistryValueKind]::String)
    }
    finally {
        $addInRoot.Close()
    }
}

function Trust-ManifestPublisher {
    param(
        [string]$ManifestPath
    )

    [xml]$manifestXml = Get-Content -Path $ManifestPath -Raw
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($manifestXml.NameTable)
    $namespaceManager.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#")

    $certificateNode = $manifestXml.SelectSingleNode("//ds:X509Certificate", $namespaceManager)
    if ($null -eq $certificateNode -or [string]::IsNullOrWhiteSpace($certificateNode.InnerText)) {
        throw "Could not find the signing certificate in the VSTO manifest: $ManifestPath"
    }

    $certificateBytes = [System.Convert]::FromBase64String($certificateNode.InnerText)
    $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$certificateBytes)
    $tempCertificatePath = Join-Path $env:TEMP "AJTools-Publisher.cer"

    try {
        [System.IO.File]::WriteAllBytes($tempCertificatePath, $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))

        if (-not (Get-ChildItem "Cert:\CurrentUser\TrustedPublisher" | Where-Object Thumbprint -eq $certificate.Thumbprint)) {
            Import-Certificate -FilePath $tempCertificatePath -CertStoreLocation "Cert:\CurrentUser\TrustedPublisher" | Out-Null
        }
    }
    finally {
        Remove-Item -Path $tempCertificatePath -Force -ErrorAction SilentlyContinue
    }
}

function Get-VersionFromPackageName {
    param(
        [string]$Path
    )

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    if ($fileName -match 'AJTools-(\d+\.\d+\.\d+\.\d+)$') {
        return $Matches[1]
    }

    throw "Could not determine the AJ Tools version from package name: $Path"
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

if (Get-Process WINPROJ -ErrorAction SilentlyContinue) {
    throw "Close Microsoft Project before installing AJ Tools."
}

$PackageZip = [System.IO.Path]::GetFullPath($PackageZip)
if (-not (Test-Path $PackageZip)) {
    throw "Package ZIP not found: $PackageZip"
}

$Version = Get-VersionFromPackageName -Path $PackageZip
$VersionFolderName = "AJTools_" + $Version.Replace('.', '_')
$ApplicationFilesRoot = Join-Path $InstallRoot "Application Files"
$DataRoot = Join-Path $InstallRoot "Data"
$LogsRoot = Join-Path $DataRoot "Logs"
$StagingRoot = Join-Path $InstallRoot "Staging"
$TargetVersionRoot = Join-Path $ApplicationFilesRoot $VersionFolderName
$ManifestPath = Join-Path $TargetVersionRoot "Arian Jahandarfards MS Project Add-in.vsto"
$StatePath = Join-Path $InstallRoot "state.json"
$UpdateFeedOverridePath = Join-Path $InstallRoot "update-feed-url.txt"

New-Item -ItemType Directory -Path $ApplicationFilesRoot -Force | Out-Null
New-Item -ItemType Directory -Path $LogsRoot -Force | Out-Null
New-Item -ItemType Directory -Path $StagingRoot -Force | Out-Null

Remove-AJProjectRegistryKeys
Remove-AJProjectAddInData
Remove-AJVstoMetadata
Remove-AJVstoSecurity
Remove-AJDeploymentCacheEntries
Remove-AJClickOnceCacheFiles

$stagingPath = Join-Path $StagingRoot $VersionFolderName
if (Test-Path $stagingPath) {
    Remove-Item -LiteralPath $stagingPath -Recurse -Force
}

New-Item -ItemType Directory -Path $stagingPath -Force | Out-Null
Expand-Archive -Path $PackageZip -DestinationPath $stagingPath -Force

if (Test-Path $TargetVersionRoot) {
    Remove-Item -LiteralPath $TargetVersionRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $TargetVersionRoot -Force | Out-Null
Copy-DirectoryTree -SourceRoot $stagingPath -DestinationRoot $TargetVersionRoot

if (-not (Test-Path $ManifestPath)) {
    throw "Installed VSTO manifest not found: $ManifestPath"
}

Trust-ManifestPublisher -ManifestPath $ManifestPath
Register-ProjectAddInKey -ManifestPath $ManifestPath

$state = [ordered]@{
    currentVersion      = $Version
    currentManifestPath = $ManifestPath
    previousVersion     = $null
    previousManifestPath = $null
    pendingValidation   = $false
    lastUpdateUtc       = (Get-Date).ToUniversalTime().ToString("o")
    lastHealthyUtc      = (Get-Date).ToUniversalTime().ToString("o")
    lastPackageSource   = $PackageZip
}

$state | ConvertTo-Json | Set-Content -Path $StatePath -Encoding UTF8

if ([string]::IsNullOrWhiteSpace($UpdateFeedUrl)) {
    Remove-Item -Path $UpdateFeedOverridePath -Force -ErrorAction SilentlyContinue
}
else {
    Set-Content -Path $UpdateFeedOverridePath -Value $UpdateFeedUrl.Trim() -Encoding ASCII
}

Write-Host ""
Write-Host "AJ Tools installed successfully."
Write-Host "Install root: $InstallRoot"
Write-Host "Version folder: $TargetVersionRoot"
Write-Host "Manifest: $ManifestPath"
if (-not [string]::IsNullOrWhiteSpace($UpdateFeedUrl)) {
    Write-Host "Update feed override: $UpdateFeedUrl"
}
Write-Host ""
Write-Host "Next step: open Microsoft Project and confirm the AJ ribbon loads."

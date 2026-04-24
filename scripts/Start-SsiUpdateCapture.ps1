param(
    [string]$OutputDir = $(Join-Path $env:TEMP ("ssi-update-capture-" + (Get-Date -Format "yyyyMMdd-HHmmss"))),
    [int]$DurationMinutes = 20
)

$ErrorActionPreference = "Stop"

$ssiRoot = Join-Path $env:LOCALAPPDATA "SSI_Tools"
$appFilesRoot = Join-Path $ssiRoot "Application Files"
$tempRoot = $env:TEMP
$projectAddinKey = "HKCU:\Software\Microsoft\Office\MS Project\Addins\SSIToolsForMSProject"
$projectAddinDataKey = "HKCU:\Software\Microsoft\Office\MS Project\AddinsData\SSIToolsForMSProject"
$solutionMetadataKey = "HKCU:\Software\Microsoft\VSTO\SolutionMetadata"
$securityInclusionKey = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion"
$uninstallKey = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{8128EB31-5BC6-4647-9A16-FE65C179EC54}_is1"
$updateInfoUrl = "https://ssitools.com/ssitoolsupdates/SSIToolsForMSProject/NoClickOnceSSIToolsForMSProjectUpdateInfo.json"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$registryDir = Join-Path $OutputDir "registry"
$inventoryDir = Join-Path $OutputDir "inventory"
$snapshotsDir = Join-Path $OutputDir "snapshots"
New-Item -ItemType Directory -Force -Path $registryDir | Out-Null
New-Item -ItemType Directory -Force -Path $inventoryDir | Out-Null
New-Item -ItemType Directory -Force -Path $snapshotsDir | Out-Null

$fsLog = Join-Path $OutputDir "fs-events.log"
$tempLog = Join-Path $OutputDir "temp-events.log"
$procLog = Join-Path $OutputDir "process-events.log"
$stateLog = Join-Path $OutputDir "state.log"
$netLog = Join-Path $OutputDir "network.log"
$ssiLogDelta = Join-Path $OutputDir "ssi-log-delta.log"
$toolingLog = Join-Path $OutputDir "tooling.log"

function Write-Log {
    param(
        [string]$Path,
        [string]$Message
    )

    Add-Content -Path $Path -Value ("[{0}] {1}" -f (Get-Date -Format o), $Message)
}

function Get-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Export-JsonFile {
    param(
        [string]$Path,
        $Object,
        [int]$Depth = 8
    )

    $Object | ConvertTo-Json -Depth $Depth | Set-Content -Path $Path -Encoding UTF8
}

function Get-RegistryNode {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return [ordered]@{
            Path = $Path
            Exists = $false
        }
    }

    try {
        $item = Get-Item -Path $Path -ErrorAction Stop
        $properties = @{}
        foreach ($name in $item.Property) {
            try {
                $properties[$name] = $item.GetValue($name)
            }
            catch {
            }
        }

        return [ordered]@{
            Path = $Path
            Exists = $true
            Values = $properties
            SubKeys = @($item.GetSubKeyNames())
        }
    }
    catch {
        return [ordered]@{
            Path = $Path
            Exists = $true
            Error = $_.Exception.Message
        }
    }
}

function Export-RegistrySnapshot {
    param([string]$Label)

    $snapshot = [ordered]@{
        Timestamp = (Get-Date).ToString("o")
        Label = $Label
        ProjectAddin = Get-RegistryNode -Path $projectAddinKey
        ProjectAddinData = Get-RegistryNode -Path $projectAddinDataKey
        SolutionMetadata = Get-RegistryNode -Path $solutionMetadataKey
        SecurityInclusion = Get-RegistryNode -Path $securityInclusionKey
        UninstallEntry = Get-RegistryNode -Path $uninstallKey
    }

    Export-JsonFile -Path (Join-Path $registryDir ("registry-{0}.json" -f $Label)) -Object $snapshot
}

function Get-VersionFolders {
    if (-not (Test-Path $appFilesRoot)) {
        return @()
    }

    return @(Get-ChildItem -Path $appFilesRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "SSIToolsForMSProject_*" } |
        Sort-Object Name |
        ForEach-Object {
            [ordered]@{
                Name = $_.Name
                FullName = $_.FullName
                LastWriteTimeUtc = $_.LastWriteTimeUtc.ToString("o")
            }
        })
}

function Get-LatestSsiLog {
    $logRoot = Join-Path $ssiRoot "Logs"
    if (-not (Test-Path $logRoot)) {
        return $null
    }

    return Get-ChildItem -Path $logRoot -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Get-FileInventory {
    param([string]$Root)

    if (-not (Test-Path $Root)) {
        return @()
    }

    return @(Get-ChildItem -Path $Root -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        ForEach-Object {
            $hash = $null
            try {
                $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
            }
            catch {
            }

            [ordered]@{
                RelativePath = $_.FullName.Substring($Root.Length).TrimStart('\')
                Length = $_.Length
                LastWriteTimeUtc = $_.LastWriteTimeUtc.ToString("o")
                Sha256 = $hash
            }
        })
}

function Export-Inventory {
    param(
        [string]$Label,
        [string]$Root,
        [string]$Name
    )

    $payload = [ordered]@{
        Timestamp = (Get-Date).ToString("o")
        Label = $Label
        Root = $Root
        Files = Get-FileInventory -Root $Root
    }

    Export-JsonFile -Path (Join-Path $inventoryDir ("{0}-{1}.json" -f $Name, $Label)) -Object $payload
}

function Get-InterestingProcesses {
    return @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match "^(WINPROJ|msiexec|powershell|cmd|wscript|cscript|rundll32)\.exe$" -or
            ($_.ExecutablePath -and $_.ExecutablePath -like "*SSI_Tools*") -or
            ($_.CommandLine -and $_.CommandLine -like "*SSI_Tools*")
        } |
        Sort-Object Name, ProcessId |
        ForEach-Object {
            [ordered]@{
                Name = $_.Name
                ProcessId = $_.ProcessId
                ExecutablePath = $_.ExecutablePath
                CommandLine = $_.CommandLine
            }
        })
}

function Get-InterestingConnections {
    param([int[]]$ProcessIds)

    if ($null -eq $ProcessIds -or $ProcessIds.Count -eq 0) {
        return @()
    }

    if (-not (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue)) {
        return @()
    }

    return @(Get-NetTCPConnection -ErrorAction SilentlyContinue |
        Where-Object {
            $_.OwningProcess -in $ProcessIds -and
            $_.State -in @("Established", "SynSent", "CloseWait", "TimeWait")
        } |
        Sort-Object OwningProcess, RemoteAddress, RemotePort |
        ForEach-Object {
            [ordered]@{
                OwningProcess = $_.OwningProcess
                LocalAddress = $_.LocalAddress
                LocalPort = $_.LocalPort
                RemoteAddress = $_.RemoteAddress
                RemotePort = $_.RemotePort
                State = $_.State
            }
        })
}

function Get-DirectoryAcl {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return $null
    }

    try {
        $acl = Get-Acl -LiteralPath $Path
        return [ordered]@{
            Path = $Path
            Owner = $acl.Owner
            Access = @($acl.Access | ForEach-Object {
                [ordered]@{
                    IdentityReference = $_.IdentityReference.ToString()
                    FileSystemRights = $_.FileSystemRights.ToString()
                    AccessControlType = $_.AccessControlType.ToString()
                    IsInherited = $_.IsInherited
                }
            })
        }
    }
    catch {
        return [ordered]@{
            Path = $Path
            Error = $_.Exception.Message
        }
    }
}

function Export-Snapshot {
    param([string]$Label)

    $processes = Get-InterestingProcesses
    $connections = Get-InterestingConnections -ProcessIds @($processes | ForEach-Object { [int]$_.ProcessId })
    $latestLog = Get-LatestSsiLog

    $payload = [ordered]@{
        Timestamp = (Get-Date).ToString("o")
        Label = $Label
        IsAdmin = Get-IsAdmin
        UpdateInfoUrl = $updateInfoUrl
        VersionFolders = Get-VersionFolders
        Processes = $processes
        Connections = $connections
        LatestLog = if ($latestLog) {
            [ordered]@{
                FullName = $latestLog.FullName
                Length = $latestLog.Length
                LastWriteTimeUtc = $latestLog.LastWriteTimeUtc.ToString("o")
            }
        } else {
            $null
        }
        SsiRootAcl = Get-DirectoryAcl -Path $ssiRoot
        TempAcl = Get-DirectoryAcl -Path $tempRoot
    }

    Export-JsonFile -Path (Join-Path $snapshotsDir ("snapshot-{0}.json" -f $Label)) -Object $payload
}

function Try-SaveUpdateInfo {
    try {
        $response = Invoke-WebRequest -Uri $updateInfoUrl -UseBasicParsing -ErrorAction Stop
        Set-Content -Path (Join-Path $OutputDir "update-info.json") -Value $response.Content -Encoding UTF8
        Write-Log -Path $stateLog -Message "Saved live update-info JSON."
    }
    catch {
        Write-Log -Path $stateLog -Message ("Could not save update-info JSON: {0}" -f $_.Exception.Message)
    }
}

function Matches-InterestingTempPath {
    param([string]$Path)

    return $Path -match 'ssi|ssitools|\.zip$|\.vsto$|\.manifest$|clickonce|setup'
}

Write-Log -Path $stateLog -Message ("Capture starting. OutputDir={0}" -f $OutputDir)
Write-Log -Path $toolingLog -Message ("IsAdmin={0}" -f (Get-IsAdmin))
Try-SaveUpdateInfo
Export-RegistrySnapshot -Label "before"
Export-Snapshot -Label "before"
Export-Inventory -Label "before" -Root $ssiRoot -Name "ssi-root"

$watcher = $null
$tempWatcher = $null
$subscriptions = @()
$netshStarted = $false
$netshTraceFile = Join-Path $OutputDir "network-trace.etl"

try {
    if (Get-IsAdmin) {
        try {
            & netsh trace start capture=yes report=no persistent=no tracefile="$netshTraceFile" | Out-Null
            $netshStarted = $true
            Write-Log -Path $toolingLog -Message ("Started netsh trace at {0}" -f $netshTraceFile)
        }
        catch {
            Write-Log -Path $toolingLog -Message ("Could not start netsh trace: {0}" -f $_.Exception.Message)
        }
    }
    else {
        Write-Log -Path $toolingLog -Message "Skipping netsh trace because shell is not elevated."
    }

    if (Test-Path $ssiRoot) {
        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path = $ssiRoot
        $watcher.Filter = "*"
        $watcher.IncludeSubdirectories = $true
        $watcher.NotifyFilter = [System.IO.NotifyFilters]"FileName, DirectoryName, LastWrite, CreationTime, Size"
        $watcher.EnableRaisingEvents = $true

        $subscriptions += Register-ObjectEvent -InputObject $watcher -EventName Created -Action {
            Add-Content -Path $using:fsLog -Value ("[{0}] CREATED {1}" -f (Get-Date -Format o), $Event.SourceEventArgs.FullPath)
        }
        $subscriptions += Register-ObjectEvent -InputObject $watcher -EventName Changed -Action {
            Add-Content -Path $using:fsLog -Value ("[{0}] CHANGED {1}" -f (Get-Date -Format o), $Event.SourceEventArgs.FullPath)
        }
        $subscriptions += Register-ObjectEvent -InputObject $watcher -EventName Deleted -Action {
            Add-Content -Path $using:fsLog -Value ("[{0}] DELETED {1}" -f (Get-Date -Format o), $Event.SourceEventArgs.FullPath)
        }
        $subscriptions += Register-ObjectEvent -InputObject $watcher -EventName Renamed -Action {
            Add-Content -Path $using:fsLog -Value ("[{0}] RENAMED {1} -> {2}" -f (Get-Date -Format o), $Event.SourceEventArgs.OldFullPath, $Event.SourceEventArgs.FullPath)
        }
    }

    if (Test-Path $tempRoot) {
        $tempWatcher = New-Object System.IO.FileSystemWatcher
        $tempWatcher.Path = $tempRoot
        $tempWatcher.Filter = "*"
        $tempWatcher.IncludeSubdirectories = $false
        $tempWatcher.NotifyFilter = [System.IO.NotifyFilters]"FileName, DirectoryName, LastWrite, CreationTime, Size"
        $tempWatcher.EnableRaisingEvents = $true

        $subscriptions += Register-ObjectEvent -InputObject $tempWatcher -EventName Created -Action {
            if ($Event.SourceEventArgs.FullPath -match 'ssi|ssitools|\.zip$|\.vsto$|\.manifest$|clickonce|setup') {
                Add-Content -Path $using:tempLog -Value ("[{0}] CREATED {1}" -f (Get-Date -Format o), $Event.SourceEventArgs.FullPath)
            }
        }
        $subscriptions += Register-ObjectEvent -InputObject $tempWatcher -EventName Changed -Action {
            if ($Event.SourceEventArgs.FullPath -match 'ssi|ssitools|\.zip$|\.vsto$|\.manifest$|clickonce|setup') {
                Add-Content -Path $using:tempLog -Value ("[{0}] CHANGED {1}" -f (Get-Date -Format o), $Event.SourceEventArgs.FullPath)
            }
        }
        $subscriptions += Register-ObjectEvent -InputObject $tempWatcher -EventName Renamed -Action {
            if ($Event.SourceEventArgs.FullPath -match 'ssi|ssitools|\.zip$|\.vsto$|\.manifest$|clickonce|setup' -or
                $Event.SourceEventArgs.OldFullPath -match 'ssi|ssitools|\.zip$|\.vsto$|\.manifest$|clickonce|setup') {
                Add-Content -Path $using:tempLog -Value ("[{0}] RENAMED {1} -> {2}" -f (Get-Date -Format o), $Event.SourceEventArgs.OldFullPath, $Event.SourceEventArgs.FullPath)
            }
        }
    }

    $subscriptions += Register-WmiEvent -Class Win32_ProcessStartTrace -Action {
        $pid = $Event.SourceEventArgs.ProcessID
        $name = $Event.SourceEventArgs.ProcessName
        try {
            $proc = Get-CimInstance Win32_Process -Filter ("ProcessId={0}" -f $pid) -ErrorAction SilentlyContinue
            $path = $proc.ExecutablePath
            $cmd = $proc.CommandLine
        }
        catch {
            $path = $null
            $cmd = $null
        }

        if ($name -match "^(WINPROJ|msiexec|powershell|cmd|wscript|cscript|rundll32)\.exe$" -or
            ($path -and $path -like "*SSI_Tools*") -or
            ($cmd -and $cmd -like "*SSI_Tools*")) {
            Add-Content -Path $using:procLog -Value ("[{0}] START PID={1} Name={2} Path={3} CommandLine={4}" -f (Get-Date -Format o), $pid, $name, $path, $cmd)
        }
    }

    $subscriptions += Register-WmiEvent -Class Win32_ProcessStopTrace -Action {
        $pid = $Event.SourceEventArgs.ProcessID
        $name = $Event.SourceEventArgs.ProcessName
        if ($name -match "^(WINPROJ|msiexec|powershell|cmd|wscript|cscript|rundll32)\.exe$") {
            Add-Content -Path $using:procLog -Value ("[{0}] STOP PID={1} Name={2}" -f (Get-Date -Format o), $pid, $name)
        }
    }

    $deadline = (Get-Date).AddMinutes($DurationMinutes)
    $lastAddinManifest = (Get-RegistryNode -Path $projectAddinKey | ConvertTo-Json -Compress)
    $lastVersionFolders = (Get-VersionFolders | ConvertTo-Json -Compress)
    $lastConnections = ""
    $trackedLogFile = Get-LatestSsiLog
    $trackedLogLength = if ($trackedLogFile) { $trackedLogFile.Length } else { 0 }

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 1

        $addinState = Get-RegistryNode -Path $projectAddinKey
        $addinJson = $addinState | ConvertTo-Json -Compress
        if ($addinJson -ne $lastAddinManifest) {
            Export-RegistrySnapshot -Label ("change-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmssfff"))
            Write-Log -Path $stateLog -Message ("Project add-in registry changed: {0}" -f $addinJson)
            $lastAddinManifest = $addinJson
        }

        $versionFolders = Get-VersionFolders
        $foldersJson = $versionFolders | ConvertTo-Json -Compress
        if ($foldersJson -ne $lastVersionFolders) {
            Write-Log -Path $stateLog -Message ("Version folders changed: {0}" -f $foldersJson)
            $lastVersionFolders = $foldersJson
        }

        $processes = Get-InterestingProcesses
        $connections = Get-InterestingConnections -ProcessIds @($processes | ForEach-Object { [int]$_.ProcessId })
        $connectionJson = $connections | ConvertTo-Json -Compress
        if ($connectionJson -ne $lastConnections) {
            Export-JsonFile -Path (Join-Path $snapshotsDir "connections-latest.json") -Object $connections
            if ($connections.Count -gt 0) {
                Write-Log -Path $netLog -Message ("Connections changed: {0}" -f $connectionJson)
            }
            $lastConnections = $connectionJson
        }

        $latestLog = Get-LatestSsiLog
        if ($latestLog) {
            if ($null -eq $trackedLogFile -or $trackedLogFile.FullName -ne $latestLog.FullName) {
                Write-Log -Path $stateLog -Message ("SSI log file changed to: {0}" -f $latestLog.FullName)
                $trackedLogFile = $latestLog
                $trackedLogLength = 0
            }

            if ($latestLog.Length -gt $trackedLogLength) {
                try {
                    $stream = [System.IO.File]::Open($latestLog.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    try {
                        $stream.Seek($trackedLogLength, [System.IO.SeekOrigin]::Begin) | Out-Null
                        $reader = New-Object System.IO.StreamReader($stream)
                        $delta = $reader.ReadToEnd()
                    }
                    finally {
                        $stream.Dispose()
                    }

                    if (-not [string]::IsNullOrWhiteSpace($delta)) {
                        Add-Content -Path $ssiLogDelta -Value ("[{0}] LOGFILE {1}" -f (Get-Date -Format o), $latestLog.FullName)
                        Add-Content -Path $ssiLogDelta -Value $delta
                    }
                }
                catch {
                    Write-Log -Path $stateLog -Message ("Could not read SSI log delta: {0}" -f $_.Exception.Message)
                }

                $trackedLogLength = $latestLog.Length
            }
        }
    }

    Export-RegistrySnapshot -Label "after"
    Export-Snapshot -Label "after"
    Export-Inventory -Label "after" -Root $ssiRoot -Name "ssi-root"
    Write-Log -Path $stateLog -Message "Capture completed."
}
finally {
    if ($netshStarted) {
        try {
            & netsh trace stop | Out-Null
            Write-Log -Path $toolingLog -Message "Stopped netsh trace."
        }
        catch {
            Write-Log -Path $toolingLog -Message ("Could not stop netsh trace cleanly: {0}" -f $_.Exception.Message)
        }
    }

    foreach ($subscription in $subscriptions) {
        try {
            Unregister-Event -SubscriptionId $subscription.Id -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    if ($watcher -ne $null) {
        try {
            $watcher.Dispose()
        }
        catch {
        }
    }

    if ($tempWatcher -ne $null) {
        try {
            $tempWatcher.Dispose()
        }
        catch {
        }
    }
}

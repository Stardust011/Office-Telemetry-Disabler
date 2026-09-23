<#
Office Privacy and Telemetry Disabler (Office 16.0 baseline)
- Default: apply telemetry/privacy baseline aligned to OC2R_DisableTelemetry.reg
- Optional: disable Office updates and update-related scheduled tasks (default: No)
- Optional: block telemetry domains in hosts file (default: No)
- Supports restore from backup created by this script
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Restore,
    [string]$BackupPath,
    [switch]$DisableOfficeUpdates,
    [switch]$SkipHosts,
    [string]$LogPath
)

$Colors = @{
    Title    = 'Cyan'
    Section  = 'Yellow'
    Success  = 'Green'
    Info     = 'Blue'
    Warning  = 'Yellow'
    Error    = 'Red'
    Gray     = 'Gray'
    Changed  = 'Magenta'
    Skip     = 'DarkGray'
}

$OfficeVersion = '16.0'
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$BackupDir = Join-Path $ScriptRoot 'backup'

$BaselineRegistrySettings = @(
    @{ Path='HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry'; Name='DisableTelemetry'; Type='DWord'; Value=1; Description='Disable common telemetry' },
    @{ Path='HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry'; Name='VerboseLogging'; Type='DWord'; Value=0; Description='Disable verbose logging' },
    @{ Path='HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry'; Name='SendTelemetry'; Type='DWord'; Value=3; Description='Set telemetry level to minimum' },
    @{ Path='HKCU:\SOFTWARE\Policies\Microsoft\Office\Common\ClientTelemetry'; Name='SendTelemetry'; Type='DWord'; Value=3; Description='Set telemetry level to minimum (Policies)' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\ClientTelemetry"; Name='DisableTelemetry'; Type='DWord'; Value=1; Description='Disable telemetry' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\ClientTelemetry"; Name='VerboseLogging'; Type='DWord'; Value=0; Description='Disable verbose logging' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common"; Name='QMEnable'; Type='DWord'; Value=0; Description='Disable CEIP' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common"; Name='sendcustomerdata'; Type='DWord'; Value=0; Description='Disable customer data collection' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common"; Name='linkedin'; Type='DWord'; Value=0; Description='Disable LinkedIn integration' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common"; Name='sendcustomerdata'; Type='DWord'; Value=0; Description='Disable customer data collection (Policies)' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common"; Name='linkedin'; Type='DWord'; Value=0; Description='Disable LinkedIn integration (Policies)' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Feedback"; Name='Enabled'; Type='DWord'; Value=0; Description='Disable feedback' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Feedback"; Name='includescreenshot'; Type='DWord'; Value=0; Description='Disable screenshot in feedback' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Feedback"; Name='includeemail'; Type='DWord'; Value=0; Description='Disable email in feedback' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Feedback"; Name='surveyenabled'; Type='DWord'; Value=0; Description='Disable surveys' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common\Feedback"; Name='includescreenshot'; Type='DWord'; Value=0; Description='Disable screenshot in feedback (Policies)' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common\Feedback"; Name='includeemail'; Type='DWord'; Value=0; Description='Disable email in feedback (Policies)' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common\Feedback"; Name='surveyenabled'; Type='DWord'; Value=0; Description='Disable surveys (Policies)' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Privacy"; Name='UserContentDisabled'; Type='DWord'; Value=2; Description='Disable content analysis' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Privacy"; Name='DownloadContentDisabled'; Type='DWord'; Value=2; Description='Disable online content download' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Privacy"; Name='disconnectedstate'; Type='DWord'; Value=2; Description='Set disconnected state' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Privacy"; Name='controllerconnectedservicesenabled'; Type='DWord'; Value=2; Description='Disable controller connected services' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common\Privacy"; Name='disconnectedstate'; Type='DWord'; Value=2; Description='Set disconnected state (Policies)' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common\Privacy"; Name='controllerconnectedservicesenabled'; Type='DWord'; Value=2; Description='Disable controller connected services (Policies)' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Internet"; Name='UseOnlineContent'; Type='DWord'; Value=0; Description='Disable online content' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Internet"; Name='serviceleveloptions'; Type='DWord'; Value=0; Description='Disable service level options' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common\Internet"; Name='serviceleveloptions'; Type='DWord'; Value=0; Description='Disable service level options (Policies)' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\PTWatson"; Name='PTWOptIn'; Type='DWord'; Value=0; Description='Disable PTWatson opt-in' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Outlook\Options\Mail"; Name='EnableLogging'; Type='DWord'; Value=0; Description='Disable Outlook mail logging' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Word\Options"; Name='EnableLogging'; Type='DWord'; Value=0; Description='Disable Word logging' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\OSM"; Name='EnableLogging'; Type='DWord'; Value=0; Description='Disable OSM logging' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\OSM"; Name='EnableUpload'; Type='DWord'; Value=0; Description='Disable OSM upload' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\OSM"; Name='EnableLogging'; Type='DWord'; Value=0; Description='Disable OSM logging (Policies)' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\OSM"; Name='EnableUpload'; Type='DWord'; Value=0; Description='Disable OSM upload (Policies)' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Lync"; Name='disableautomaticsendtracing'; Type='DWord'; Value=1; Description='Disable Lync automatic tracing' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Lync"; Name='disableautomaticsendtracing'; Type='DWord'; Value=1; Description='Disable Lync automatic tracing (Policies)' },

    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\Security\FileValidation"; Name='disablereporting'; Type='DWord'; Value=1; Description='Disable file validation reporting' },
    @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Common\Security\FileValidation"; Name='disablereporting'; Type='DWord'; Value=1; Description='Disable file validation reporting (Policies)' }
)

$ClickToRunPrivacySettings = @(
    @{ Path='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name='SendTelemetry'; Type='DWord'; Value=3; Description='Set telemetry level to minimum (ClickToRun)' },
    @{ Path='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name='UserContentDisabled'; Type='DWord'; Value=2; Description='Disable content analysis (ClickToRun)' },
    @{ Path='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name='DownloadContentDisabled'; Type='DWord'; Value=2; Description='Disable online content download (ClickToRun)' },
    @{ Path='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name='DisconnectedState'; Type='DWord'; Value=2; Description='Set disconnected state (ClickToRun)' },
    @{ Path='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name='ControllerConnectedServicesEnabled'; Type='DWord'; Value=2; Description='Disable controller connected services (ClickToRun)' },
    @{ Path='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name='skydrivesigninoption'; Type='DWord'; Value=0; Description='Disable OneDrive sign-in (ClickToRun)' }
)

$OptionalUpdateRegistrySettings = @(
    @{ Path='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name='UpdatesEnabled'; Type='String'; Value='False'; Description='Disable automatic updates (ClickToRun)' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Update"; Name='OfficeMgmtCOM'; Type='DWord'; Value=0; Description='Disable Office management COM' },
    @{ Path="HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Update"; Name='EnableAutomaticUpdates'; Type='DWord'; Value=0; Description='Disable automatic updates' }
)

$OptionalUpdateTasks = @(
    @{ Name='Microsoft\Office\Office Automatic Updates'; Description='Office Automatic Updates' },
    @{ Name='Microsoft\Office\Office Automatic Updates 2.0'; Description='Office Automatic Updates 2.0' },
    @{ Name='Microsoft\Office\Office Feature Updates'; Description='Office Feature Updates' },
    @{ Name='Microsoft\Office\Office Feature Updates Logon'; Description='Office Feature Updates Logon' },
    @{ Name='Microsoft\Office\Office 16 Subscription Heartbeat'; Description='Office 16 Subscription Heartbeat' },
    @{ Name='Microsoft\Office\Office Subscription Maintenance'; Description='Office Subscription Maintenance' },
    @{ Name='Microsoft\Office\Office ClickToRun Service Monitor'; Description='Office ClickToRun Service Monitor' }
)

$TelemetryDomains = @(
    'vortex.data.microsoft.com',
    'vortex-win.data.microsoft.com',
    'telecommand.telemetry.microsoft.com',
    'oca.telemetry.microsoft.com',
    'sqm.telemetry.microsoft.com',
    'watson.telemetry.microsoft.com'
)

function Test-Admin {
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Convert-RegistryKindToType {
    param([string]$Kind)
    switch ($Kind) {
        'String' { 'String' }
        'ExpandString' { 'ExpandString' }
        'Binary' { 'Binary' }
        'DWord' { 'DWord' }
        'QWord' { 'QWord' }
        'MultiString' { 'MultiString' }
        default { 'String' }
    }

    function Get-RegistryKindFromType {
        param([string]$Type)

        switch ($Type) {
            'String' { 'String' }
            'ExpandString' { 'ExpandString' }
            'Binary' { 'Binary' }
            'DWord' { 'DWord' }
            'QWord' { 'QWord' }
            'MultiString' { 'MultiString' }
            default { 'String' }
        }
    }
}

function Get-RegistryState {
    param(
        [string]$Path,
        [string]$Name
    )

    $result = @{ Existed = $false; Value = $null; Kind = $null }

    if (-not (Test-Path $Path)) {
        return $result
    }

    try {
        $key = Get-Item -Path $Path -ErrorAction Stop
        if ($key.GetValueNames() -contains $Name) {
            $result.Existed = $true
            $result.Value = $key.GetValue($Name)
            $result.Kind = [string]$key.GetValueKind($Name)
        }
    }
    catch {
    }

    return $result
}

function Add-RegistryBackupRecord {
    param(
        [hashtable]$Backup,
        [string]$Path,
        [string]$Name,
        [string]$Type,
        $NewValue,
        [string]$Description
    )

    $state = Get-RegistryState -Path $Path -Name $Name

    $Backup.Registry += @{
        Path        = $Path
        Name        = $Name
        Existed     = $state.Existed
        Value       = $state.Value
        Kind        = $state.Kind
        NewType     = $Type
        NewValue    = $NewValue
        Description = $Description
    }
}

function Set-RegistryValueWithBackup {
    param(
        [hashtable]$Setting,
        [hashtable]$Backup
    )

    $path = $Setting.Path
    $name = $Setting.Name
    $type = $Setting.Type
    $value = $Setting.Value

    try {
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force -ErrorAction Stop | Out-Null
        }

        $state = Get-RegistryState -Path $path -Name $name
        $targetKind = Get-RegistryKindFromType -Type $type
        if ($state.Existed -and $state.Kind -eq $targetKind -and $state.Value -eq $value) {
            Write-Host "  [SKIP] Already set: $name -> $value" -ForegroundColor $Colors.Skip
            return
        }

        Add-RegistryBackupRecord -Backup $Backup -Path $path -Name $name -Type $type -NewValue $value -Description $Setting.Description
        if (-not $PSCmdlet.ShouldProcess("$path\$name", "Set registry value to $value ($type)")) {
            Write-Host "  [SKIP] WhatIf/ShouldProcess skipped: $path\$name" -ForegroundColor $Colors.Skip
            return
        }
        Set-ItemProperty -Path $path -Name $name -Type $type -Value $value -Force -ErrorAction Stop

        Write-Host "  [OK] Changed: $name -> $value" -ForegroundColor $Colors.Success
        if ($Setting.Description) {
            Write-Host "    Description: $($Setting.Description)" -ForegroundColor $Colors.Info
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to set: $path\$name" -ForegroundColor $Colors.Error
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor $Colors.Error
    }
}

function Get-TaskEnabledState {
    param([string]$TaskName)

    $taskPath = "\$TaskName"
    $xmlResult = & schtasks.exe /Query /TN $taskPath /XML 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $xmlResult) {
        return $null
    }

    try {
        $taskXml = [xml]($xmlResult -join [Environment]::NewLine)
        $enabledNode = $taskXml.Task.Settings.Enabled
        if ($enabledNode -eq $null) {
            return $true
        }

        return $enabledNode -eq 'true'
    }
    catch {
        return $null
    }
}

function Disable-TaskWithBackup {
    param(
        [hashtable]$Task,
        [hashtable]$Backup
    )

    $enabledState = Get-TaskEnabledState -TaskName $Task.Name
    if ($enabledState -eq $null) {
        Write-Host "  [SKIP] Task not found: $($Task.Name)" -ForegroundColor $Colors.Skip
        return
    }

    if (-not $enabledState) {
        Write-Host "  [SKIP] Task already disabled: $($Task.Name)" -ForegroundColor $Colors.Skip
        return
    }

    $taskPath = "\$($Task.Name)"
    if (-not $PSCmdlet.ShouldProcess($Task.Name, 'Disable scheduled task')) {
        Write-Host "  [SKIP] WhatIf/ShouldProcess skipped: $($Task.Name)" -ForegroundColor $Colors.Skip
        return
    }

    & schtasks.exe /Change /TN $taskPath /DISABLE 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $Backup.Tasks += @{ Name = $Task.Name; PreviousEnabled = $true; Description = $Task.Description }
        Write-Host "  [OK] Disabled task: $($Task.Name)" -ForegroundColor $Colors.Success
    }
    else {
        Write-Host "  [ERROR] Failed to disable task: $($Task.Name)" -ForegroundColor $Colors.Error
    }
}

function Save-Backup {
    param([hashtable]$Backup)

    if (-not (Test-Path $BackupDir)) {
        New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null
    }

    $file = Join-Path $BackupDir ("office_telemetry_backup_{0}.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $Backup | ConvertTo-Json -Depth 8 | Set-Content -Path $file -Encoding UTF8
    Write-Host "`n[INFO] Backup saved: $file" -ForegroundColor $Colors.Info
    return $file
}

function Get-LatestBackupFile {
    if (-not (Test-Path $BackupDir)) {
        return $null
    }

    return Get-ChildItem -Path $BackupDir -Filter 'office_telemetry_backup_*.json' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Restore-FromBackup {
    param([string]$InputBackupPath)

    $resolvedPath = $InputBackupPath
    if (-not $resolvedPath) {
        $latest = Get-LatestBackupFile
        if ($latest) {
            $resolvedPath = $latest.FullName
        }
    }

    if (-not $resolvedPath -or -not (Test-Path $resolvedPath)) {
        Write-Host "[ERROR] No backup file found for restore." -ForegroundColor $Colors.Error
        return
    }

    Write-Host "[INFO] Restoring from backup: $resolvedPath" -ForegroundColor $Colors.Info
    $backup = Get-Content -Path $resolvedPath -Raw | ConvertFrom-Json

    foreach ($record in $backup.Registry) {
        try {
            if ($record.Existed) {
                if (-not (Test-Path $record.Path)) {
                    New-Item -Path $record.Path -Force | Out-Null
                }
                $restoreType = if ($record.Kind) { Convert-RegistryKindToType -Kind $record.Kind } else { 'String' }
                if (-not $PSCmdlet.ShouldProcess("$($record.Path)\$($record.Name)", 'Restore registry value from backup')) {
                    Write-Host "  [SKIP] WhatIf/ShouldProcess skipped: $($record.Path)\$($record.Name)" -ForegroundColor $Colors.Skip
                    continue
                }
                Set-ItemProperty -Path $record.Path -Name $record.Name -Type $restoreType -Value $record.Value -Force -ErrorAction Stop
                Write-Host "  [OK] Restored value: $($record.Path)\$($record.Name)" -ForegroundColor $Colors.Success
            }
            else {
                if (Test-Path $record.Path) {
                    if (-not $PSCmdlet.ShouldProcess("$($record.Path)\$($record.Name)", 'Remove registry value created by script')) {
                        Write-Host "  [SKIP] WhatIf/ShouldProcess skipped: $($record.Path)\$($record.Name)" -ForegroundColor $Colors.Skip
                        continue
                    }
                    Remove-ItemProperty -Path $record.Path -Name $record.Name -ErrorAction SilentlyContinue
                    Write-Host "  [OK] Removed value created by script: $($record.Path)\$($record.Name)" -ForegroundColor $Colors.Success
                }
            }
        }
        catch {
            Write-Host "  [ERROR] Failed to restore: $($record.Path)\$($record.Name)" -ForegroundColor $Colors.Error
        }
    }

    foreach ($task in $backup.Tasks) {
        if (-not $task.PreviousEnabled) {
            continue
        }

        $taskPath = "\$($task.Name)"
        if (-not $PSCmdlet.ShouldProcess($task.Name, 'Re-enable scheduled task from backup')) {
            Write-Host "  [SKIP] WhatIf/ShouldProcess skipped: $($task.Name)" -ForegroundColor $Colors.Skip
            continue
        }

        & schtasks.exe /Change /TN $taskPath /ENABLE 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  [OK] Re-enabled task: $($task.Name)" -ForegroundColor $Colors.Success
        }
    }

    if ($backup.Hosts -and $backup.Hosts.Applied -and $backup.Hosts.BackupFile -and (Test-Path $backup.Hosts.BackupFile)) {
        try {
            if (-not $PSCmdlet.ShouldProcess($backup.Hosts.HostsFile, "Restore hosts file from backup $($backup.Hosts.BackupFile)")) {
                Write-Host "  [SKIP] WhatIf/ShouldProcess skipped hosts restore." -ForegroundColor $Colors.Skip
                return
            }
            Copy-Item -Path $backup.Hosts.BackupFile -Destination $backup.Hosts.HostsFile -Force -ErrorAction Stop
            Write-Host "  [OK] Restored hosts file from backup" -ForegroundColor $Colors.Success
        }
        catch {
            Write-Host "  [ERROR] Failed to restore hosts file: $($_.Exception.Message)" -ForegroundColor $Colors.Error
        }
    }

    Write-Host "`n[OK] Restore complete." -ForegroundColor $Colors.Success
}

function Apply-HostsBlocking {
    param([hashtable]$Backup)

    $hostsFile = "$env:windir\System32\drivers\etc\hosts"
    $backupFile = "$hostsFile.backup.$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    try {
        if (-not $PSCmdlet.ShouldProcess($hostsFile, 'Add telemetry domain blocking entries')) {
            Write-Host "  [SKIP] WhatIf/ShouldProcess skipped hosts modification." -ForegroundColor $Colors.Skip
            return
        }

        Copy-Item -Path $hostsFile -Destination $backupFile -Force -ErrorAction Stop

        $hostsContent = Get-Content -Path $hostsFile -Raw -ErrorAction SilentlyContinue
        if (-not $hostsContent) { $hostsContent = '' }

        $addedCount = 0
        foreach ($domain in $TelemetryDomains) {
            if ($hostsContent -notmatch "0\.0\.0\.0\s+$([regex]::Escape($domain))" -and
                $hostsContent -notmatch "127\.0\.0\.1\s+$([regex]::Escape($domain))") {
                Add-Content -Path $hostsFile -Value "0.0.0.0 $domain" -Force -ErrorAction Stop
                $addedCount++
            }
        }

        $Backup.Hosts = @{
            Applied = $true
            HostsFile = $hostsFile
            BackupFile = $backupFile
            AddedCount = $addedCount
        }

        Write-Host "  [OK] Hosts blocking applied. Added entries: $addedCount" -ForegroundColor $Colors.Success
    }
    catch {
        Write-Host "  [ERROR] Failed to modify hosts file: $($_.Exception.Message)" -ForegroundColor $Colors.Error
    }
}

Write-Host "--- Office Privacy and Telemetry Disabler (Office 16.0) ---" -ForegroundColor $Colors.Title
$transcriptStarted = $false

if (-not (Test-Admin)) {
    Write-Host "[ERROR] Administrator privileges are required." -ForegroundColor $Colors.Error
    Read-Host 'Press Enter to exit'
    exit 1
}

if ($LogPath) {
    try {
        Start-Transcript -Path $LogPath -Append -ErrorAction Stop | Out-Null
        Write-Host "[INFO] Logging enabled: $LogPath" -ForegroundColor $Colors.Info
        $transcriptStarted = $true
    }
    catch {
        Write-Host "[WARN] Failed to start transcript logging: $($_.Exception.Message)" -ForegroundColor $Colors.Warning
    }
}

if ($Restore) {
    Restore-FromBackup -InputBackupPath $BackupPath
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
    Read-Host 'Press Enter to exit'
    exit 0
}

Write-Host "[INFO] HKCU settings apply to the current user context running this script." -ForegroundColor $Colors.Warning
Write-Host "[INFO] In domain-managed environments, Group Policy may override local settings (check gpresult /r)." -ForegroundColor $Colors.Warning

$officeRunning = Get-Process WINWORD, EXCEL, POWERPNT, OUTLOOK -ErrorAction SilentlyContinue
if ($officeRunning) {
    $runningNames = ($officeRunning | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
    Write-Host "[WARN] Office processes detected: $runningNames. Close Office apps before applying for best reliability." -ForegroundColor $Colors.Warning
}

$hasClickToRun = Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$backup = @{
    Metadata = @{
        Tool = 'Office Privacy and Telemetry Disabler'
        OfficeVersion = $OfficeVersion
        Timestamp = (Get-Date).ToString('o')
    }
    Registry = @()
    Tasks = @()
    Hosts = @{ Applied = $false }
}

Write-Host "`n--- Applying Office 16.0 telemetry/privacy baseline ---" -ForegroundColor $Colors.Section
foreach ($setting in $BaselineRegistrySettings) {
    Set-RegistryValueWithBackup -Setting $setting -Backup $backup
}

if ($hasClickToRun) {
    Write-Host "`n--- Applying Click-to-Run privacy settings ---" -ForegroundColor $Colors.Section
    foreach ($setting in $ClickToRunPrivacySettings) {
        Set-RegistryValueWithBackup -Setting $setting -Backup $backup
    }
}
else {
    Write-Host "`n[SKIP] Click-to-Run not detected. Skipping Click-to-Run privacy keys." -ForegroundColor $Colors.Skip
}

$disableUpdates = $false
if ($DisableOfficeUpdates) {
    $disableUpdates = $true
}
else {
    $disableUpdatesAnswer = Read-Host "`nOptional: disable Office updates and update-related tasks? (y/N)"
    $disableUpdates = $disableUpdatesAnswer -match '^[Yy]$'
}
if ($disableUpdates) {
    Write-Host "`n--- Applying optional update disable settings ---" -ForegroundColor $Colors.Section
    foreach ($setting in $OptionalUpdateRegistrySettings) {
        if ($setting.Path -like 'HKLM:*' -and -not $hasClickToRun) {
            continue
        }
        Set-RegistryValueWithBackup -Setting $setting -Backup $backup
    }

    foreach ($task in $OptionalUpdateTasks) {
        Disable-TaskWithBackup -Task $task -Backup $backup
    }
}
else {
    Write-Host "[SKIP] Optional update disable settings were not selected." -ForegroundColor $Colors.Skip
}

$applyHostsBlocking = $false
if ($SkipHosts) {
    Write-Host "[SKIP] Hosts blocking skipped by parameter." -ForegroundColor $Colors.Skip
}
else {
    $hostsAnswer = Read-Host "`nOptional: block telemetry domains in hosts file? (y/N)"
    $applyHostsBlocking = $hostsAnswer -match '^[Yy]$'
}

if ($applyHostsBlocking) {
    Write-Host "`n--- Applying optional hosts blocking ---" -ForegroundColor $Colors.Section
    Apply-HostsBlocking -Backup $backup
}
else {
    Write-Host "[SKIP] Optional hosts blocking was not selected." -ForegroundColor $Colors.Skip
}

$backupFile = Save-Backup -Backup $backup

Write-Host "`n--- Summary ---" -ForegroundColor $Colors.Section
Write-Host "  > Office target version: $OfficeVersion" -ForegroundColor $Colors.Info
Write-Host "  > Click-to-Run detected: $hasClickToRun" -ForegroundColor $Colors.Info
Write-Host "  > Optional update disable selected: $disableUpdates" -ForegroundColor $Colors.Info
Write-Host "  > Optional hosts blocking selected: $applyHostsBlocking" -ForegroundColor $Colors.Info
Write-Host "  > Backup file: $backupFile" -ForegroundColor $Colors.Info
Write-Host "`n[OK] Completed." -ForegroundColor $Colors.Success

if ($transcriptStarted) {
    Stop-Transcript | Out-Null
}

Read-Host 'Press Enter to exit'

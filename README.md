# Office Privacy and Telemetry Disabler (Office 16.0)

PowerShell tool for Office **16.0** telemetry/privacy hardening (Office 2016/2019/2021/2024/Microsoft 365 family).

## What changed

- The repository now targets **Office 16.0 only**.
- `OC2R_DisableTelemetry.reg` is provided as the baseline profile.
- Update-disabling behavior is split into **optional prompts** and is **off by default**.
- A backup/restore flow is included. Restore only reverts values/tasks/hosts entries recorded by this tool backup.

## Files

- `script/office_privacy_telemetry_disabler.ps1` – main script (Office 16.0 only)
- `OC2R_DisableTelemetry.reg` – baseline registry profile (UTF-16 LE, `.reg`)
- `Launcher.bat` – elevated launcher

## Default behavior (no optional selections)

The script applies Office telemetry/privacy settings aligned with `OC2R_DisableTelemetry.reg`, including:

- Client telemetry / CEIP / feedback related keys
- Privacy + online content policy keys
- OSM/PTWatson/logging related keys
- Click-to-Run privacy keys (if Click-to-Run exists)

By default, it **does not**:

- disable Office updates
- disable Office update tasks
- modify update policy values (`UpdatesEnabled`, `EnableAutomaticUpdates`, etc.)

## Optional actions (default OFF)

When prompted, you can optionally enable:

1. **Disable Office updates and update-related tasks**
   - Sets update-related registry values
   - Disables update-related scheduled tasks
2. **Hosts blocking**
   - Adds telemetry domains to hosts

If you do not opt in, those actions are skipped.

## Parameters

- `-DisableOfficeUpdates`: non-interactive opt-in to disable update policy keys and update tasks.
- `-SkipHosts`: skip hosts modification without prompt.
- `-Restore [-BackupPath <file>]`: restore from latest/specified backup.
- `-LogPath <file>`: enable transcript logging.
- `-WhatIf`: preview changes without applying them (`SupportsShouldProcess`).

## Restore / rollback

Each run creates a backup JSON in `script/backup/`.

Restore command:

```powershell
.\script\office_privacy_telemetry_disabler.ps1 -Restore
```

Or restore a specific backup:

```powershell
.\script\office_privacy_telemetry_disabler.ps1 -Restore -BackupPath "C:\path\to\office_telemetry_backup_YYYYMMDD_HHMMSS.json"
```

Restore logic is symmetric to tracked changes only:

- restore previous registry values (or remove values created by this tool)
- re-enable only tasks this tool disabled from enabled state
- restore hosts file from backup if this tool modified hosts

## Requirements

- Windows with Administrator privileges
- PowerShell 5.1+ or PowerShell 7+
- Office 16.0 family

## Usage

### Option 1: Launcher

Run `Launcher.bat` as Administrator.

### Option 2: PowerShell directly

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\script\office_privacy_telemetry_disabler.ps1
```

Example with switches:

```powershell
.\script\office_privacy_telemetry_disabler.ps1 -DisableOfficeUpdates -SkipHosts -LogPath .\run.log
```

## Baseline `.reg` import (optional)

`OC2R_DisableTelemetry.reg` uses standard `.reg` format and UTF-16 LE encoding.

Manual import example:

```cmd
reg import OC2R_DisableTelemetry.reg
```

## Notes

- Restart Office apps after applying changes.
- HKCU keys apply to the user context running the script (elevated credentials may target a different profile).
- In domain-managed environments, Group Policy may override local settings (`gpresult /r`).
- Hosts blocking list is intentionally conservative and excludes activation/update/login endpoint domains.
- Always create your own system/registry backup before running system-modifying scripts.
- This project reduces configured Office telemetry/privacy exposure but does **not** claim to completely block all Microsoft data collection in every environment.

## License

MIT

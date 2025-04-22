# CronTask Scheduler PowerShell Module

## Overview

The **CronTask Scheduler** module provides a PowerShell‐native way to create, update and manage Windows Scheduled Tasks using:

- **Cron expressions** (five‑field syntax)  
- **Relative date targets** (e.g. “Last Wednesday of the month”)  
- **Event log triggers** (filter by log name, Event ID, provider)  

All tasks live under the `\cron` folder in Task Scheduler and run under the SYSTEM account by default.

## Prerequisites

- Windows PowerShell 5.1 or later (Windows)  
- Or PowerShell Core (7.x) on Windows  
- Administrative rights to register/unregister scheduled tasks and event sources  

## Installation

You can install `win-cron` either directly from the PowerShell Gallery or manually:

### 1. Install from PowerShell Gallery

```powershell
# Requires PowerShellGet module (built‑in on PS 5.1+)
Install-Module -Name win-cron -Scope AllUsers
# then import if needed
Import-Module win-cron
````

### 2. Manual Installation

1. Clone or download the repository (or just the `.psm1` file) to your machine:
    
    ```none
    git clone https://github.com/BlueConfetti/win‑cron.git C:\Modules\win-cron
    ```
    
2. From an elevated PowerShell prompt:
    
    ```powershell
    Import-Module -Name 'C:\Modules\win-cron\WinCron.psm1'
    ```
    
3. Verify availability:
    
    ```powershell
    Get-Command -Module win-cron
    ```
    
## Public Cmdlets

### Get-RelativeTargetDate

Calculates a date in a month based on an occurrence (“First”, “Second”, “Third”, “Fourth”, “Last”) and weekday.

```powershell
Get-RelativeTargetDate
  -Occurrence   <First|Second|Third|Fourth|Last>
  -DayOfWeek    <Sunday|Monday|…|Saturday>
  [-Month       <DateTime>]
```

### New-CronTask

Creates and registers a new scheduled task combining:

- A PowerShell `ScriptBlock` action  
- A cron‐based calendar trigger (`-CronExpression`)  
- Optional relative date adjustments (`-RelativeTargets`)  
- Optional one‑shot behavior (`-RunOnce`)  
- Optional event log trigger (`-EventLogTrigger` / `-EventLogTriggerOverrides`)

```powershell
New-CronTask
  -Name                     <string>           # Task name (also event source if used)
  -ScriptBlock              <ScriptBlock>      # Code to execute
  [-CronExpression          <string>]          # e.g. "0 12 * * *"
  [-RelativeTargets         <string[]>]        # e.g. @("Last,Wednesday,02:00:00")
  [-BaseDate                <DateTime>]        # Defaults to (Get-Date)
  [-RunOnce]                                  # Switch → single‐shot TimeTrigger
  [-EventLogTrigger]                          # Switch → subscribe to event log
  [-EventLogTriggerOverrides <string[]>]      # e.g. @("Application,1000,MyApp")
```

### Update-CronTask

Appends new triggers and/or replaces the action script of an existing task without removing existing triggers.

```powershell
Update-CronTask
  -Name                     <string>
  [-ScriptBlock             <ScriptBlock>]
  [-CronExpression          <string>]
  [-RelativeTargets         <string[]>]
  [-BaseDate                <DateTime>]
  [-RunOnce]
  [-EventLogTrigger]
  [-EventLogTriggerOverrides <string[]>]
```

### Get-CronTask

Lists all tasks under `\cron` (or filters by name).

```powershell
# All cron tasks
Get-CronTask

# By name
Get-CronTask -Name "MyTask"
```

### Remove-CronTask

Removes one or more tasks (and their event log sources) by piped input or name.

```powershell
# By pipeline
Get-CronTask -Name "MyTask" | Remove-CronTask

# By name
Remove-CronTask -Name "MyTask"
```

### Trigger-CronEvent

Manually writes an informational event (`EventId 1006`) to the `win-cron` log to kick off any event‑triggered tasks.

```powershell
Trigger-CronEvent -Name "MyTask"
```

## Examples

1. **Daily at noon**  
   ```powershell
   New-CronTask -Name "DailyNoon" `
     -ScriptBlock { Write-Output "Hello at noon" } `
     -CronExpression "0 12 * * *"
   ```

2. **Last Wednesday at 18:00 + 2h offset**  
   ```powershell
   New-CronTask -Name "PayrollReminder" `
     -ScriptBlock { Send-MailMessage ... } `
     -CronExpression "0 0 * * *" `
     -RelativeTargets @("Last,Wednesday,02:00:00")
   ```

3. **One-shot run next matching**  
   ```powershell
   New-CronTask -Name "OneTimeReport" `
     -ScriptBlock { Export-Csv ... } `
     -CronExpression "30 9 * * 1" `
     -RunOnce
   ```

4. **Trigger on event log entry**  
   ```powershell
   New-CronTask -Name "DiskFailWatcher" `
     -ScriptBlock { Restart-Service -Name 'Spooler' } `
     -EventLogTrigger `
     -EventLogTriggerOverrides @("System,51,disk")
   ```

5. **Update an existing task**  
   ```powershell
   Update-CronTask -Name "DailyNoon" `
     -ScriptBlock { Write-Output "Updated script!" } `
     -CronExpression "0 15 * * *"  # also append 3:15 PM trigger
   ```

6. **List & remove**  
   ```powershell
   Get-CronTask | Format-Table TaskName,TaskPath
   Get-CronTask -Name "DailyNoon" | Remove-CronTask
   ```

7. **Manual event kick**  
   ```powershell
   Trigger-CronEvent -Name "DiskFailWatcher"
   ```

## Troubleshooting

- **Insufficient privileges** – run PowerShell as administrator.  
- **Cron syntax errors** – ensure a valid five‑field expression.  
- **Task never runs** – use `-Verbose` on `New-CronTask` / `Update-CronTask` to see XML generation and registration logs.  

## Contributing

Feel free to fork, improve the module, and submit pull requests. Please write unit tests for new features and update this README accordingly.

## License

MIT License — see [LICENSE](LICENSE.md) for details.
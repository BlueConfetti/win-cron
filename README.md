# CronTask Scheduler PowerShell Module

## Overview

CronTask Scheduler provides a PowerShell-friendly interface for creating and managing Windows Scheduled Tasks. It supports:

- Cron expressions (five fields)
- Relative date targets (for example, "Last Wednesday")
- Event log triggers

Tasks are created in the `\cron` folder and run under the SYSTEM account by default.

## Prerequisites

- PowerShell 5.1 or later
- Administrator rights to register tasks and event log sources

## Installation

### From the PowerShell Gallery

```powershell
Install-Module -Name win-cron -Scope AllUsers
Import-Module win-cron   # if not auto-imported
```

### Manual Install

1. Clone or download this repository or copy `win-cron.psm1` somewhere:

   ```powershell
   git clone https://github.com/BlueConfetti/win-cron.git C:\Modules\win-cron
   ```

2. Import the module:

   ```powershell
   Import-Module 'C:\Modules\win-cron\win-cron.psm1'
   ```

3. Verify the commands:

   ```powershell
   Get-Command -Module win-cron
   ```

## Public Cmdlets

### New-CronTask

```powershell
New-CronTask -Name <string> -ScriptBlock <ScriptBlock> [-CronExpression <string>] `
             [-RelativeTargets <string[]>] [-BaseDate <DateTime>] [-RunOnce] `
             [-EventLogTrigger] [-EventLogTriggerOverrides <string[]>]
```

Creates a task with one or more triggers.

### Update-CronTask

```powershell
Update-CronTask -Name <string> [-ScriptBlock <ScriptBlock>] `
                [-CronExpression <string>] [-RelativeTargets <string[]>] `
                [-BaseDate <DateTime>] [-RunOnce] [-EventLogTrigger] `
                [-EventLogTriggerOverrides <string[]>]
```

Adds additional triggers or updates the action for an existing task.

### Get-CronTask

```powershell
Get-CronTask [-Name <string>]
```

Lists cron tasks under `\cron`.

### Remove-CronTask

```powershell
Remove-CronTask -Name <string>
```

Removes cron tasks and their event sources.

### Trigger-CronEvent

```powershell
Trigger-CronEvent -Name <string>
```

Writes an event to the `win-cron` log to trigger event-based tasks.

## Helper Function

### Get-RelativeTargetDate

```powershell
Get-RelativeTargetDate -Occurrence <First|Second|Third|Fourth|Last> `
                       -DayOfWeek  <Sunday|Monday|...|Saturday> `
                       [-Month <DateTime>]
```

Returns the target date in a month, such as the last Wednesday.

## Examples

### Daily at noon

```powershell
New-CronTask -Name "DailyNoon" `
  -ScriptBlock { Write-Output "Hello at noon" } `
  -CronExpression "0 12 * * *"
```

### Last Wednesday at 18:00 plus 2 hours

```powershell
New-CronTask -Name "PayrollReminder" `
  -ScriptBlock { Send-MailMessage ... } `
  -CronExpression "0 0 * * *" `
  -RelativeTargets @("Last,Wednesday,02:00:00")
```

### One-shot task on the next Monday at 09:30

```powershell
New-CronTask -Name "OneTimeReport" `
  -ScriptBlock { Export-Csv ... } `
  -CronExpression "30 9 * * 1" `
  -RunOnce
```

### Triggered by event log entry

```powershell
New-CronTask -Name "DiskFailWatcher" `
  -ScriptBlock { Restart-Service Spooler } `
  -EventLogTrigger `
  -EventLogTriggerOverrides @("System,51,disk")
```

### Update an existing task

```powershell
Update-CronTask -Name "DailyNoon" `
  -ScriptBlock { Write-Output "Updated script!" } `
  -CronExpression "0 15 * * *"   # also run at 3:15 PM
```

### List and remove tasks

```powershell
Get-CronTask | Format-Table TaskName,TaskPath
Remove-CronTask -Name "DailyNoon"
```

### Manually fire an event

```powershell
Trigger-CronEvent -Name "DiskFailWatcher"
```

## Troubleshooting

- **Insufficient privileges** – run PowerShell as administrator.
- **Cron syntax errors** – verify the five-field cron expression.
- **Task never runs** – run `New-CronTask` or `Update-CronTask` with `-Verbose` for details.

## Contributing

Feel free to fork the project, add improvements and open pull requests.

## License

MIT License – see [LICENSE](LICENSE.md) for details.

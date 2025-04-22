# CronTask Scheduler Module

## Overview

The CronTask Scheduler Module is a PowerShell toolkit for creating, updating, and managing Windows scheduled tasks using both time‐based triggers (via cron expressions) and event log–based triggers. It allows you to build tasks that run on a schedule or in response to specific events—with support for relative date calculations (e.g. "Last Wednesday of the month") and event filtering (by Event ID and provider).

## Features
- **Time-based Scheduling:**
  Create tasks using standard five-field cron expressions.
- **Relative Date Scheduling:**
  Use relative target dates (like "First" or "Last" occurrence of a weekday) as a base for scheduling.
- **Event Log Triggers:**
  Subscribe to event log entries (with filtering by log name, Event ID, and optionally provider) so that tasks run when a matching event is logged.
- **Task Management:**
  Easily register, query, and remove tasks via built-in functions.

## Functions
- **Get-RelativeTargetDate**
  Calculates a target date within a month based on an occurrence (e.g. "Last") and a day-of-week.
- **Convert-CronToSchedule**
  Converts a five-field cron string into a schedule object.
- **Parse-CronField**
  Parses individual fields in a cron expression.
- **Expand-CronSchedule**
  Expands a cron expression into one or more schedule objects.
- **Generate-ScheduleXmlBlocks**
  Creates XML trigger blocks from schedule objects for time-based tasks.
- **Register-CronTask**
  Connects to the Task Scheduler service and registers a task using a complete XML definition.
- **New-CronTask**
  Combines a PowerShell ScriptBlock with scheduling parameters (cron, relative, and event triggers) to generate the XML definition and register a new scheduled task.
- **Get-CronTask**
  Retrieves scheduled tasks created using this module.
- **Remove-CronTask**
  Removes a scheduled task by name or input object.

## Installation
1. **Clone or Download the Repository:**
   Download or clone the repository to your local system.
2. **Import the Module:**
   Open PowerShell (preferably with administrative privileges) and run:

   ```powershell
   Import-Module -Path "C:\Path\To\YourModule.psm1"
   ```

## Usage

### Creating a Time-Based Task

Create a scheduled task that runs at a specific time using a cron expression:

```powershell
New-CronTask -Name "DailyTask" -ScriptBlock { Write-Output "Task executed" } -CronExpression "0 12 * * *"
```

This command creates a task named "DailyTask" that runs every day at 12:00 PM.

### Creating a Task with a Relative Target Date

Create a task that runs relative to a target date (e.g. the last Wednesday of the month):

```powershell
New-CronTask -Name "RelativeTask" -ScriptBlock { Write-Output "Relative schedule executed" } `
    -CronExpression "0 12 * * *" -RelativeTarget -Occurrence Last -DayOfWeek Wednesday -BaseDate (Get-Date)
```

### Creating an Event Log Trigger Task

Create a task that triggers when a specific event is logged:

```powershell
New-CronTask -Name "EventTask" -ScriptBlock { Write-Output "Event triggered task executed" } `
    -EventLogTrigger -EventLogName "System" -EventId 105 -ProviderName "Microsoft-Windows-Kernel-Power"
```

This creates a task named "EventTask" that triggers when an event with ID 105 is recorded in the System log from the provider "Microsoft-Windows-Kernel-Power".

### Managing Tasks
- **List Tasks:**

  ```powershell
  Get-CronTask
  ```

- **Remove a Task:**

  ```powershell
  Get-CronTask -Name "DailyTask" | Remove-CronTask
  ```

## Customization
- **Scheduling Options:**
  Adjust default parameter values and add further validations as needed.

## Troubleshooting
- **Administrative Rights:**
  Ensure you run PowerShell with the necessary privileges when creating or modifying scheduled tasks.
- **Cron Expression Validation:**
  Validate your cron expressions if tasks do not fire as expected.
- **Verbose Logging:**
  Use the -Verbose switch with functions to see detailed output for troubleshooting.

## Contributing

Contributions are welcome! Please fork the repository, create your changes, and submit a pull request. Issues and feature requests can be submitted via the repository's issue tracker.

## License

This project is licensed under the MIT License.


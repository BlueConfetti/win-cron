function Get-RelativeTargetDate {
  <#
  .SYNOPSIS
      Returns a target date in a given month based on a relative occurrence and day-of-week.
  
  .DESCRIPTION
      Given an occurrence (First, Second, Third, Fourth, or Last) and a day-of-week (e.g. Wednesday),
      this function returns the corresponding date in the month. By default the current month is used,
      but an alternate base date (to extract the year/month) can be provided.
  
  .PARAMETER Occurrence
      One of "First", "Second", "Third", "Fourth", or "Last".
  
  .PARAMETER DayOfWeek
      A day of the week (Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday).
  
  .PARAMETER Month
      A DateTime whose Year and Month are used for the calculation. Defaults to the current date.
  
  .EXAMPLE
      Get-RelativeTargetDate -Occurrence Last -DayOfWeek Wednesday
      Returns the last Wednesday of the current month.
  
  .EXAMPLE
      Get-RelativeTargetDate -Occurrence Third -DayOfWeek Thursday -Month (Get-Date "2025-03-01")
      Returns the third Thursday in March 2025.
  #>
  [CmdletBinding()]
  param(
      [Parameter(Mandatory = $true)]
      [ValidateSet("First", "Second", "Third", "Fourth", "Last")]
      [string]$Occurrence,
      
      [Parameter(Mandatory = $true)]
      [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
      [string]$DayOfWeek,
      
      [Parameter(Mandatory = $false)]
      [datetime]$Month = (Get-Date)
  )
  
  $year = $Month.Year
  $monthNumber = $Month.Month

  if ($Occurrence -eq "Last") {
      # Find the last day of the month then work backwards.
      $daysInMonth = [DateTime]::DaysInMonth($year, $monthNumber)
      $date = Get-Date -Year $year -Month $monthNumber -Day $daysInMonth
      while ($date.DayOfWeek.ToString() -ne $DayOfWeek) {
          $date = $date.AddDays(-1)
      }
      return $date
  }
  else {
      # Convert occurrence string to numeric index.
      $occurrenceIndex = switch ($Occurrence) {
          "First"  { 1 }
          "Second" { 2 }
          "Third"  { 3 }
          "Fourth" { 4 }
          default  { throw "Invalid occurrence: $Occurrence" }
      }
      
      # Start at the first day of the month and count matches.
      $date = Get-Date -Year $year -Month $monthNumber -Day 1
      $matchCount = 0
      while ($date.Month -eq $monthNumber) {
          if ($date.DayOfWeek.ToString() -eq $DayOfWeek) {
              $matchCount++
              if ($matchCount -eq $occurrenceIndex) {
                  return $date
              }
          }
          $date = $date.AddDays(1)
      }
      throw "Could not find the $Occurrence $DayOfWeek in the month."
  }
}

# Parse-CronField: Determines whether a cron field is a wildcard, interval (using slash), or a fixed value.
function Parse-CronField {
  param(
      [Parameter(Mandatory=$true)]
      [string]$Field,
      [Parameter(Mandatory=$true)]
      [string]$FieldName
  )
  $result = @{
      IsWildcard    = $false;
      IsInterval    = $false;
      FixedValue    = $null;
      IntervalValue = $null;
      Original      = $Field
  }
  if ($Field -eq "*") {
      $result.IsWildcard = $true
      return $result
  }
  if ($Field -match '^\*/(\d+)$') {
      $result.IsInterval    = $true
      $result.IntervalValue = [int]$matches[1]
      return $result
  }
  if ($Field -match '^\d+$') {
      $result.FixedValue = [int]$Field
      return $result
  }
  throw "Unsupported field format for $($FieldName): $Field"
}

# Expand-CronSchedule: Parses a five-field cron string and expands it into one or more schedule objects.
function Expand-CronSchedule {
  param(
      [Parameter(Mandatory=$true)]
      [string]$CronExpression
  )
  # Split the expression into its five fields: minute, hour, day-of-month, month, day-of-week.
  $fields = $CronExpression.Trim() -split '\s+'
  if ($fields.Count -ne 5) {
      throw "Invalid cron expression. Expected 5 fields, got $($fields.Count)."
  }
  $minParsed   = Parse-CronField -Field $fields[0] -FieldName "minute"
  $hourParsed  = Parse-CronField -Field $fields[1] -FieldName "hour"
  $domParsed   = Parse-CronField -Field $fields[2] -FieldName "day-of-month"
  $monthParsed = Parse-CronField -Field $fields[3] -FieldName "month"
  $dowParsed   = Parse-CronField -Field $fields[4] -FieldName "day-of-week"
  
  $now = Get-Date
  $schedules = @()
  
  # Determine the minute settings.
  if ($minParsed.IsInterval) {
      $minInterval = $minParsed.IntervalValue
      $startMinute = 0
      $useMinuteRepetition = $true
  } elseif ($minParsed.IsWildcard) {
      $startMinute = 0
      $minInterval = 1
      $useMinuteRepetition = $true
  } else {
      $startMinute = $minParsed.FixedValue
      $useMinuteRepetition = $false
  }
  
  # Determine valid hours.
  if ($hourParsed.IsInterval) {
      $hourInterval = $hourParsed.IntervalValue
      $validHours = @()
      for ($h = 0; $h -lt 24; $h += $hourInterval) {
          $validHours += $h
      }
  } elseif ($hourParsed.IsWildcard) {
      $validHours = 0..23
  } else {
      $validHours = @($hourParsed.FixedValue)
  }
  
  # Determine valid day-of-month.
  if ($domParsed.IsInterval) {
      $domInterval = $domParsed.IntervalValue
      $validDays = @()
      for ($d = 1; $d -le 31; $d++) {
          # Assume starting at 1; valid if (d-1) mod interval == 0.
          if ((($d - 1) % $domInterval) -eq 0) {
              $validDays += $d
          }
      }
  } elseif ($domParsed.IsWildcard) {
      $validDays = $null  # No specific day restriction.
  } else {
      $validDays = @($domParsed.FixedValue)
  }
  
  # For month and day-of-week, we won’t expand further here.
  if (-not $monthParsed.IsWildcard -and $monthParsed.FixedValue) {
      $fixedMonth = $monthParsed.FixedValue
  } else {
      $fixedMonth = $null
  }
  if (-not $dowParsed.IsWildcard -and $dowParsed.FixedValue) {
      $fixedDOW = $dowParsed.FixedValue
  } else {
      $fixedDOW = $null
  }
  
  # Now generate schedule objects.
  if ($validDays) {
      # For each valid day-of-month and for each valid hour, create a schedule.
      foreach ($day in $validDays) {
          foreach ($hour in $validHours) {
              # Use fixedMonth if specified; otherwise, use the current month.
              if ($fixedMonth) {
                  $monthVal = $fixedMonth
              } else {
                  $monthVal = $now.Month
              }
              # Build a candidate start boundary.
              try {
                  $candidate = Get-Date -Year $now.Year -Month $monthVal -Day $day -Hour $hour -Minute $startMinute -Second 0
              } catch {
                  continue  # Skip invalid dates (e.g. February 30)
              }
              if ($candidate -lt $now) {
                  # If candidate is past, advance (for simplicity, add one month if month is fixed; otherwise, one day).
                  if ($fixedMonth) {
                      $candidate = $candidate.AddMonths(1)
                  } else {
                      $candidate = $candidate.AddDays(1)
                  }
              }
              $sched = @{
                  TriggerType        = "CalendarTrigger";
                  StartBoundary      = $candidate.ToString("yyyy-MM-ddTHH:mm:ss");
                  UseMinuteRepetition = $useMinuteRepetition;
                  MinuteInterval     = $minInterval;
                  RepetitionDuration = "PT60M";  # Cover the valid hour block.
                  ScheduleBy         = "Monthly";  # Because we have a day-of-month restriction.
                  DayOfMonth         = $day;
              }
              $schedules += $sched
          }
      }
  }
  else {
      # No day-of-month restriction; create one schedule per valid hour.
      foreach ($hour in $validHours) {
          $candidate = Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour $hour -Minute $startMinute -Second 0
          if ($candidate -lt $now) { $candidate = $candidate.AddDays(1) }
          $sched = @{
              TriggerType        = "CalendarTrigger";
              StartBoundary      = $candidate.ToString("yyyy-MM-ddTHH:mm:ss");
              UseMinuteRepetition = $useMinuteRepetition;
              MinuteInterval     = $minInterval;
              RepetitionDuration = "PT60M";
              ScheduleBy         = "Daily";
              DaysInterval       = 1;
          }
          $schedules += $sched
      }
  }
  return $schedules
}

function Generate-EventTriggerXml {
  [CmdletBinding()]
  param(
      [Parameter(Mandatory=$false)]
      [string]$LogName = "win-cron",
      
      [Parameter(Mandatory=$true)]
      [int]$EventId,
      
      [Parameter(Mandatory=$false)]
      [string]$ProviderName
  )
  
  # Build the XML query string
  $query = "<QueryList><Query Id='0' Path='$LogName'><Select Path='$LogName'>"
  if ($ProviderName) {
      $query += "*[System[Provider[@Name='$ProviderName'] and (EventID=$EventId)]]"
  }
  else {
      $query += "*[System[(EventID=$EventId)]]"
  }
  $query += "</Select></Query></QueryList>"
  
  # Build the XML block for the event trigger.
  $xmlBlock = @"
  <EventTrigger>
    <Enabled>true</Enabled>
    <Subscription>
      <![CDATA[
        $query
      ]]>
    </Subscription>
  </EventTrigger>
"@
  return $xmlBlock
}

# Generate-ScheduleXmlBlocks: Builds XML trigger blocks from an array of schedule objects.
function Generate-ScheduleXmlBlocks {
  param(
      [Parameter(Mandatory = $true)]
      [array]$Schedules
  )
  $xmlBlocks = @()
  foreach ($sched in $Schedules) {
      $repeatXml = ""
      if ($sched.UseMinuteRepetition) {
          $repeatXml = @"
    <Repetition>
      <Interval>PT$($sched.MinuteInterval)M</Interval>
      <Duration>$($sched.RepetitionDuration)</Duration>
    </Repetition>
"@
      }
      $calendarXml = ""
      if ($sched.ScheduleBy -eq "Monthly") {
          $calendarXml = @"
    <ScheduleByMonth>
      <DaysOfMonth>
        <Day>$($sched.DayOfMonth)</Day>
      </DaysOfMonth>
    </ScheduleByMonth>
"@
      } elseif ($sched.ScheduleBy -eq "Daily") {
          $calendarXml = @"
    <ScheduleByDay>
      <DaysInterval>$($sched.DaysInterval)</DaysInterval>
    </ScheduleByDay>
"@
      }
      $xmlBlock = @"
  <CalendarTrigger>
    <StartBoundary>$($sched.StartBoundary)</StartBoundary>
$repeatXml
$calendarXml
  </CalendarTrigger>
"@
      $xmlBlocks += $xmlBlock
  }
  return ($xmlBlocks -join "`n")
}

function Register-CronTask {
  [CmdletBinding()]
  param(
      # The name to register the task as.
      [Parameter(Mandatory = $true)]
      [string]$Name,
      # The full XML definition for the task.
      [Parameter(Mandatory = $true)]
      [string]$Xml,
      # Optional folder path in Task Scheduler where the task will be created.
      [Parameter(Mandatory = $false)]
      [string]$FolderPath = "\cron"
  )
  
  # Connect to the Task Scheduler service.
  try {
      Write-Verbose "Connecting to Task Scheduler service..."
      $taskService = New-Object -ComObject "Schedule.Service"
      $taskService.Connect()
  }
  catch {
      throw "Failed to connect to Task Scheduler service: $($_.Exception.Message)"
  }
  
  # Create the task definition from the provided XML.
  try {
      Write-Verbose "Creating new task definition..."
      $taskDef = $taskService.NewTask(0)
      $taskDef.XmlText = $Xml
  }
  catch {
      throw "Error creating task definition: $($_.Exception.Message)"
  }
  
  # Get (or create) the folder where the task will be registered.
  try {
      Write-Verbose "Retrieving folder '$FolderPath'..."
      $folder = $taskService.GetFolder($FolderPath)
  }
  catch {
      Write-Verbose "Folder '$FolderPath' does not exist. Creating it..."
      try {
          $rootFolder = $taskService.GetFolder("\")
          # Remove the leading '\' (if any) for CreateFolder.
          $folderName = $FolderPath.TrimStart("\")
          $folder = $rootFolder.CreateFolder($folderName)
      }
      catch {
          throw "Failed to create the folder '$FolderPath': $($_.Exception.Message)"
      }
  }
  
  # Register the task in the folder.
  try {
      Write-Verbose "Registering the task '$Name' in folder '$FolderPath'..."
      # The flag value "6" indicates creation or update. Here we use logon type 5 (SERVICE_ACCOUNT).
      $folder.RegisterTaskDefinition($Name, $taskDef, 6, $null, $null, 5, $null)
      Write-Output "Task '$Name' created successfully in folder '$FolderPath'."
  }
  catch {
      throw "Error creating task: $($_.Exception.Message)"
  }
}
function Trigger-CronEvent {
    <#
    .SYNOPSIS
        Triggers an event log entry for a specified cron task.

    .DESCRIPTION
        This function writes an informational event to the "win-cron" event log using the specified task's name as the event source.
        It ensures that the custom event source exists by checking and creating it if necessary.
        The event is written with a fixed EventId (1006) and a message indicating that the cron task has been triggered.
        Additionally, the function verifies that a cron task with the specified name exists before writing the event.

    .PARAMETER Name
        The name of the cron task. This name is used as the event source in the "win-cron" event log.

    .EXAMPLE
        Trigger-CronEvent -Name "MyDailyTask"
        Writes an informational event to the "win-cron" event log from the event source "MyDailyTask", indicating that the task was triggered.

    .NOTES
        If no cron task with the specified name is found, the function returns an error.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    try {
        if (!$(Get-CronTask -name $Name)) {
            Write-Error "No crontask with name: $Name"
            return
        }
    }
    catch {
        throw "Error finding crontask: $($_.Exception.Message)"
    }
    $logName = "win-cron"
    $source  = $Name
    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$logName\$source"

    # Create the custom event log source if it doesn't exist.
    if (-not (Test-Path $registryPath)) {
        Write-Verbose "Registering new source to win-cron eventlog: $source"
        New-EventLog -LogName $logName -Source $source
    }

    # Write the event with a fixed EventId and message.
    Write-Verbose "Submitting event to win-cron log from $source"
    Write-EventLog -LogName $logName -Source $source -EventId 1006 -EntryType Information -Message "Crontask: $source triggered"
}

# New-CronTask: Combines the script and schedule parts to produce the full Task Scheduler XML.
function New-CronTask {
    <#
    .SYNOPSIS
        Creates or updates a scheduled task using a cron expression and/or an event log trigger.

    .DESCRIPTION
        This function builds a Task Scheduler XML definition from a PowerShell ScriptBlock along with a cron expression.
        It supports three types of triggers:
        - Time-based scheduling: Using a five-field cron expression which can optionally be adjusted relative to a target date.
          To adjust the target date, use the -RelativeTargets parameter which accepts an array of definitions.
          Each definition must be in the format: "<Occurrence>,<DayOfWeek>,<Offset>"
          where:
            - Occurrence: One of "First", "Second", "Third", "Fourth", or "Last".
            - DayOfWeek: A day-of-week (e.g. Monday, Tuesday, etc.).
            - Offset: A timespan (in hh:mm:ss format) to be added to the computed target date.
        - Optionally, if -RunOnce is specified, a one-shot trigger is created instead of a repeating CalendarTrigger.
        - Event log–based trigger: Subscribing to events from a specified event log.
          Use the -EventLogTrigger switch to enable this trigger.
          Additionally, you may supply -EventLogTriggerOverrides as a string array (with one element)
          to override the default event log properties.
          The expected format is "<event_log>,<event_id>,<event_source/provider>".
          If not provided, the defaults are used: "win-cron" for the log, 1006 for the Event ID,
          and the task name for the provider.
        The function generates the complete XML definition and registers the task in the "\cron" folder
        of Task Scheduler. Both time-based and event log triggers can be combined in one task definition.

    .PARAMETER Name
        The name of the scheduled task. This name is also used as the event source name when EventLogTrigger is enabled.

    .PARAMETER ScriptBlock
        The script code to be executed by the task, provided as a ScriptBlock.

    .PARAMETER CronExpression
        A five-field cron expression (e.g. "0 12 * * *") that defines the schedule for the task.
        When provided, one or more calendar triggers will be generated based on the expression.

    .PARAMETER RelativeTargets
        An optional array of relative target definitions used to adjust the schedule.
        Each element should be a string in the format "<Occurrence>,<DayOfWeek>,<Offset>".
        For example, "Last,Wednesday,02:00:00" sets the date to the last Wednesday (with a default time-of-day of 18:00)
        then adds 2 hours to it.

    .PARAMETER BaseDate
        An optional base date from which the year and month are taken to calculate the relative target.
        Defaults to the current date if not provided.

    .PARAMETER RunOnce
        When specified, the task is scheduled to run once on the computed start boundary.
        This uses a <TimeTrigger> element (a one-shot trigger) instead of a repeating <CalendarTrigger>.

    .PARAMETER EventLogTrigger
        When specified, adds an event log–based trigger to the task.

    .PARAMETER EventLogTriggerOverrides
        An optional string array to override event trigger defaults.
        The first element should be in the format "<event_log>,<event_id>,<event_source/provider>".
        For example: @("CustomLog,2001,MyProvider")
        If not provided, defaults are used: "win-cron", 1006, and the task name.

    .EXAMPLE
        New-CronTask -Name "MyDailyTask" -ScriptBlock { Write-Output "Hello" } -CronExpression "0 12 * * *"
        Creates a task named "MyDailyTask" that runs daily at 12:00 PM.

    .EXAMPLE
        New-CronTask -Name "RelativeTask" -ScriptBlock { Write-Output "Relative schedule" } `
                     -CronExpression "* */2 * * *" `
                     -RelativeTargets @("Last,Wednesday,02:00:00")
        Creates a task that adjusts its start boundary based on the last Wednesday (with 18:00 as baseline)
        plus an additional 2-hour offset, using a repeating CalendarTrigger.

    .EXAMPLE
        New-CronTask -Name "OneShotTask" -ScriptBlock { Write-Output "Run once" } `
                     -CronExpression "0 12 * * *" -RunOnce
        Creates a task that runs only once using a TimeTrigger.

    .EXAMPLE
        New-CronTask -Name "EventTask" -ScriptBlock { Write-Output "Event log trigger" } `
                     -EventLogTrigger -EventLogTriggerOverrides @("CustomLog,2001,MyProvider")
        Creates a task with an event log trigger that monitors CustomLog with Event ID 2001 and uses "MyProvider" as the event source.
    #>

  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
      [Parameter(Mandatory = $true, Position = 0)]
      [ValidateNotNullOrEmpty()]
      [string]$Name,
      
      [Parameter(Mandatory = $true, Position = 1)]
      [ValidateNotNullOrEmpty()]
      [ScriptBlock]$ScriptBlock,
      
      [Parameter(Mandatory = $false, Position = 2)]
      [ValidateNotNullOrEmpty()]
      [string]$CronExpression,
      
      # Relative target definitions in the form "<Occurrence>,<DayOfWeek>,<Offset>"
      [Parameter(Mandatory = $false)]
      [string[]]$RelativeTargets,
      
      [Parameter(Mandatory = $false)]
      [datetime]$BaseDate = (Get-Date),

      # Run once switch.
      [Parameter(Mandatory = $false)]
      [switch]$RunOnce,

      # Event log trigger parameters.
      [Parameter(Mandatory = $false)]
      [switch]$EventLogTrigger,
      
      # This parameter optionally overrides the default event log values.
      [Parameter(Mandatory = $false)]
      [string[]]$EventLogTriggerOverrides
  )
  
  # Convert the ScriptBlock to string and encode it.
  try {
      $scriptText = $ScriptBlock.ToString()
  }
  catch {
      throw "Failed to convert ScriptBlock to string: $_"
  }
  $encodedScriptBytes = [System.Text.Encoding]::Unicode.GetBytes("& { $scriptText }")
  $encodedScriptBase64 = [Convert]::ToBase64String($encodedScriptBytes)
  
  # Expand the cron expression into schedule objects if provided.
  $triggerXml = ""
  if ($CronExpression) {
      $baseSchedules = Expand-CronSchedule -CronExpression $CronExpression
      $allSchedules = @()
      
      # Process RelativeTargets if provided.
      if ($RelativeTargets) {
          foreach ($relDef in $RelativeTargets) {
              $parts = $relDef -split ','
              if ($parts.Count -ne 3) {
                  throw "Invalid relative target format: '$relDef'. Expected format: <Occurrence>,<DayOfWeek>,<Offset>"
              }
              $occurrence = $parts[0].Trim()
              $dayOfWeek = $parts[1].Trim()
              $offsetStr = $parts[2].Trim()
              try {
                  # Compute the base target date.
                  $targetDate = Get-RelativeTargetDate -Occurrence $occurrence -DayOfWeek $dayOfWeek -Month $BaseDate
                  # Override time-of-day to default 18:00.
                  $targetDate = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day $targetDate.Day -Hour 18 -Minute 0 -Second 0
                  # Parse and apply the offset.
                  $offset = [timespan]::Parse($offsetStr)
                  $targetDate = $targetDate.Add($offset)
              }
              catch {
                  throw "Error processing relative target '$relDef': $_"
              }
              Write-Verbose "Relative target date calculated as: $targetDate for definition '$relDef'"
              
              foreach ($sched in $baseSchedules) {
                  $schedClone = @{}
                  foreach ($key in $sched.Keys) {
                      $schedClone[$key] = $sched[$key]
                  }
                  $originalTime = [datetime]::ParseExact($schedClone.StartBoundary, 'yyyy-MM-ddTHH:mm:ss', $null).TimeOfDay
                  $newStart = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day $targetDate.Day `
                                -Hour $originalTime.Hours -Minute $originalTime.Minutes -Second $originalTime.Seconds
                  $schedClone.StartBoundary = $newStart.ToString("yyyy-MM-ddTHH:mm:ss")
                  $allSchedules += $schedClone
              }
          }
      }
      else {
          $allSchedules = $baseSchedules
      }
      
      # Choose trigger type based on RunOnce switch.
      if ($RunOnce) {
          $firstSchedule = $allSchedules | Sort-Object StartBoundary | Select-Object -First 1
          $triggerXml = @"
  <TimeTrigger>
    <StartBoundary>$($firstSchedule.StartBoundary)</StartBoundary>
    <Enabled>true</Enabled>
  </TimeTrigger>
"@
      }
      else {
          $triggerXml = Generate-ScheduleXmlBlocks -Schedules $allSchedules
      }
  }
  
  # Process event log trigger if requested.
  if ($EventLogTrigger) {
      # Set event log properties, either from overrides or defaults.
      if ($EventLogTriggerOverrides) {
          # Expect a single element with comma separated values.
          $overrideParts = $EventLogTriggerOverrides[0] -split ','
          if ($overrideParts.Count -ne 3) {
              throw "Invalid EventLogTriggerOverrides format. Expected format: <event_log>,<event_id>,<event_source/provider>"
          }
          $eventLog = $overrideParts[0].Trim()
          $eventIdValue = [int]($overrideParts[1].Trim())
          $providerNameValue = $overrideParts[2].Trim()
      }
      else {
          $eventLog = "win-cron"
          $eventIdValue = 1006
          $providerNameValue = $Name
      }
      
      # Create event source if necessary.
      $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$eventLog\$Name"
      if (-not (Test-Path $registryPath)) {
          New-EventLog -LogName $eventLog -Source $Name
      }
      $eventTriggerXml = Generate-EventTriggerXml -LogName $eventLog -EventId $eventIdValue -ProviderName $providerNameValue
      $triggerXml += "`n" + $eventTriggerXml
  }
  
  # Build the complete Task Scheduler XML.
  $xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
<RegistrationInfo>
  <URI>\cron\$Name</URI>
</RegistrationInfo>
<Principals>
  <Principal id="Author">
    <UserId>S-1-5-18</UserId>
    <RunLevel>HighestAvailable</RunLevel>
  </Principal>
</Principals>
<Settings>
  <AllowHardTerminate>false</AllowHardTerminate>
  <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
  <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
  <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
  <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
  <StartWhenAvailable>true</StartWhenAvailable>
  <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
  <IdleSettings>
    <StopOnIdleEnd>false</StopOnIdleEnd>
    <RestartOnIdle>false</RestartOnIdle>
  </IdleSettings>
  <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
</Settings>
<Triggers>
$triggerXml
</Triggers>
<Actions Context="Author">
  <Exec>
    <Command>C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
    <Arguments>-NonInteractive -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedScriptBase64</Arguments>
  </Exec>
</Actions>
</Task>
"@
  Write-Verbose "Task XML: $xml"
  
  # Register the task.
  Register-CronTask -Name $Name -Xml $xml
}



function Get-CronTask {
  [CmdletBinding()]
  param(
      # Optional: specify a task name to filter by.
      [Parameter(Mandatory = $false)]
      [string]$Name
  )

  try {
      # Retrieve all scheduled tasks and filter by TaskPath that starts with "\cron"
      $cronTasks = Get-ScheduledTask | Where-Object {
          # Normalize TaskPath (it may include a trailing backslash)
          $_.TaskPath.TrimEnd('\') -ieq "\cron"
      }
  }
  catch {
      throw "Error retrieving scheduled tasks: $($_.Exception.Message)"
  }

  # If a specific task name was provided, filter further.
  if ($Name) {
      $cronTasks = $cronTasks | Where-Object { $_.TaskName -ieq $Name }
  }

  return $cronTasks
}

function Remove-CronTask {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
    # Accept piped input from Get-CronTask.
    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ParameterSetName = "InputObject",
        Position = 0
    )]
    [Object]$InputObject,

    # Alternatively, remove by specifying the task name directly.
    [Parameter(
        Mandatory = $true,
        ParameterSetName = "Name",
        Position = 0
    )]
    [string]$Name
    )

    begin {
    $tasksToRemove = @()
    }
    process {
    if ($PSCmdlet.ParameterSetName -eq "InputObject") {
        if ($null -eq $InputObject) { continue }
        # If the piped object is an array, add all items.
        if ($InputObject -is [System.Array]) {
            $tasksToRemove += $InputObject
        }
        else {
            $tasksToRemove += $InputObject
        }
    }
    else {
        # Retrieve the cron task(s) by name using Get-CronTask.
        $cronTasks = Get-CronTask -Name $Name
        if (-not $cronTasks) {
            Write-Warning "No cron task found with the name '$Name'."
        }
        else {
            $tasksToRemove += $cronTasks
        }
    }
    }
    end {
        foreach ($task in $tasksToRemove) {
            if ($null -ne $task) {
                if ($PSCmdlet.ShouldProcess("$($task.TaskName)", "Remove cron task from $($task.TaskPath)")) {
                    try {
                        Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
                        Write-Verbose "Removed task '$($task.TaskName)' from '$($task.TaskPath)'."
                        $eventLogPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\win-cron\$($task.TaskName)"
                        if (Test-Path $eventLogPath) {
                            try {
                                Remove-Item $eventLogPath -Force
                                Write-Verbose "Removed custom event log source for task '$($task.TaskName)'."
                            }
                            catch {
                                Write-Warning "Failed to remove event log source for task '$($task.TaskName)': $_"
                            }
                        }
                    }
                    catch {
                        Write-Warning "Failed to remove task '$($task.TaskName)': $_"
                    }
                }
            }
        }
    }
}

function Update-CronTask {
    <#
    .SYNOPSIS
         Updates an existing cron task by optionally replacing its script block and/or appending new triggers.
    
    .DESCRIPTION
         This function retrieves the full XML definition for an existing cron task using Get-CronTask
         and Export-ScheduledTask. It then can update the action (i.e. the script block) and/or append new trigger(s).
         New triggers may be time-based (using a cron expression, with optional relative adjustments and a run-once switch)
         and/or an event log trigger (with optional override properties).
         Existing triggers remain intact.
         
    .PARAMETER Name
         The name of the cron task to update.
    
    .PARAMETER ScriptBlock
         If provided, replaces the existing action script block with this new script.
    
    .PARAMETER CronExpression
         A new five-field cron expression (e.g. "0 12 * * *") for generating additional time-based triggers.
    
    .PARAMETER RelativeTargets
         (Optional) An array of relative target definitions in the format "<Occurrence>,<DayOfWeek>,<Offset>".
         These definitions are applied to the new cron expression. For example: "Last,Friday,00:30:00".
    
    .PARAMETER BaseDate
         (Optional) Base date from which to compute relative target dates. Defaults to the current date.
    
    .PARAMETER RunOnce
         If specified, any new time trigger generated from the cron expression is created as a one‑shot (TimeTrigger) trigger.
    
    .PARAMETER EventLogTrigger
         If specified, an additional event log trigger is added.
    
    .PARAMETER EventLogTriggerOverrides
         (Optional) A string array where the first element is a comma‑separated string in the format
         "<event_log>,<event_id>,<event_source/provider>". For example: @("CustomLog,2001,MyProvider").
         If not provided, defaults are used: "win-cron", 1006, and the task name.
    
    .EXAMPLE
         Update-CronTask -Name "MyTask" -ScriptBlock { Write-Output "New Script" }
         Replaces the current script block of "MyTask" with the new script.
    
    .EXAMPLE
         Update-CronTask -Name "MyTask" -CronExpression "*/15 * * * *" -RunOnce
         Appends a new one-shot time trigger (based on the given cron expression) to "MyTask".
    
    .EXAMPLE
         Update-CronTask -Name "MyTask" -EventLogTrigger -EventLogTriggerOverrides @("CustomLog,2001,MyProvider")
         Appends an event log trigger to "MyTask" with the provided override values.
    
    .EXAMPLE
         Update-CronTask -Name "MyTask" -ScriptBlock { Write-Output "Updated" } -CronExpression "0 18 * * *" `
                          -RelativeTargets @("Last,Friday,00:30:00")
         Updates the script block and appends a new time trigger adjusted to the last Friday with a 30-minute offset.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $false)]
        [ScriptBlock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [string]$CronExpression,
        
        [Parameter(Mandatory = $false)]
        [string[]]$RelativeTargets,
        
        [Parameter(Mandatory = $false)]
        [datetime]$BaseDate = (Get-Date),
        
        [Parameter(Mandatory = $false)]
        [switch]$RunOnce,
        
        [Parameter(Mandatory = $false)]
        [switch]$EventLogTrigger,
        
        [Parameter(Mandatory = $false)]
        [string[]]$EventLogTriggerOverrides
    )
    
    # Ensure that at least one update parameter is provided.
    if (-not $ScriptBlock -and -not $CronExpression -and -not $EventLogTrigger) {
        throw "No update parameters provided. Provide a new script block and/or trigger parameters."
    }
    
    Write-Verbose "Retrieving current XML for task '$Name'..."
    try {
        $currentTaskXml = Get-CronTask -Name $Name | Export-ScheduledTask
        if (-not $currentTaskXml) {
            throw "Task '$Name' not found."
        }
    }
    catch {
        throw "Error retrieving current task: $_"
    }
    
    # Parse the existing task XML into a DOM object.
    [xml]$xmlDoc = $currentTaskXml
    
    # Replace the script block (action) if a new one is provided.
    if ($ScriptBlock) {
        try {
            $newScriptText = $ScriptBlock.ToString()
        }
        catch {
            throw "Error converting ScriptBlock to string: $_"
        }
        $newEncodedScriptBytes = [System.Text.Encoding]::Unicode.GetBytes("& { $newScriptText }")
        $newEncodedScriptBase64 = [Convert]::ToBase64String($newEncodedScriptBytes)
        
        # Locate the <Arguments> element inside <Actions>/<Exec>.
        $argsNode = $xmlDoc.Task.Actions.Exec.Arguments
        if ($argsNode) {
            $argsNode.'#text' = "-NonInteractive -NoProfile -ExecutionPolicy Bypass -EncodedCommand $newEncodedScriptBase64"
            Write-Verbose "Script block updated in task '$Name'."
        }
        else {
            throw "Could not locate the Arguments node in the task XML."
        }
    }
    
    # Prepare new trigger XML fragments (empty if no new triggers are provided).
    $newTriggersXml = ""
    
    if ($CronExpression) {
        # Generate schedule objects from the cron expression.
        $baseSchedules = Expand-CronSchedule -CronExpression $CronExpression
        $allSchedules = @()
        
        if ($RelativeTargets) {
            foreach ($relDef in $RelativeTargets) {
                $parts = $relDef -split ','
                if ($parts.Count -ne 3) {
                    throw "Invalid relative target format '$relDef'. Expected format: <Occurrence>,<DayOfWeek>,<Offset>"
                }
                $occurrence = $parts[0].Trim()
                $dayOfWeek = $parts[1].Trim()
                $offsetStr = $parts[2].Trim()
                try {
                    $targetDate = Get-RelativeTargetDate -Occurrence $occurrence -DayOfWeek $dayOfWeek -Month $BaseDate
                    # Default time-of-day is set to 18:00.
                    $targetDate = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day $targetDate.Day -Hour 18 -Minute 0 -Second 0
                    $offset = [timespan]::Parse($offsetStr)
                    $targetDate = $targetDate.Add($offset)
                }
                catch {
                    throw "Error processing relative target '$relDef': $_"
                }
                foreach ($sched in $baseSchedules) {
                    $schedClone = @{}
                    foreach ($key in $sched.Keys) {
                        $schedClone[$key] = $sched[$key]
                    }
                    $originalTime = [datetime]::ParseExact($schedClone.StartBoundary, 'yyyy-MM-ddTHH:mm:ss', $null).TimeOfDay
                    $newStart = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day $targetDate.Day `
                                   -Hour $originalTime.Hours -Minute $originalTime.Minutes -Second $originalTime.Seconds
                    $schedClone.StartBoundary = $newStart.ToString("yyyy-MM-ddTHH:mm:ss")
                    $allSchedules += $schedClone
                }
            }
        }
        else {
            $allSchedules = $baseSchedules
        }
        
        if ($RunOnce) {
            # For a one-shot trigger, select the earliest schedule.
            $firstSchedule = $allSchedules | Sort-Object StartBoundary | Select-Object -First 1
            $newTriggersXml += @"
  <TimeTrigger>
    <StartBoundary>$($firstSchedule.StartBoundary)</StartBoundary>
    <Enabled>true</Enabled>
  </TimeTrigger>
"@
        }
        else {
            $newTriggersXml += (Generate-ScheduleXmlBlocks -Schedules $allSchedules)
        }
    }
    
    if ($EventLogTrigger) {
        if ($EventLogTriggerOverrides) {
            $overrideParts = $EventLogTriggerOverrides[0] -split ','
            if ($overrideParts.Count -ne 3) {
                throw "Invalid EventLogTriggerOverrides format. Expected: <event_log>,<event_id>,<event_source/provider>"
            }
            $eventLog = $overrideParts[0].Trim()
            $eventIdValue = [int]($overrideParts[1].Trim())
            $providerNameValue = $overrideParts[2].Trim()
        }
        else {
            $eventLog = "win-cron"
            $eventIdValue = 1006
            $providerNameValue = $Name
        }
        
        $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$eventLog\$Name"
        if (-not (Test-Path $registryPath)) {
            New-EventLog -LogName $eventLog -Source $Name
        }
        $eventTriggerXml = Generate-EventTriggerXml -LogName $eventLog -EventId $eventIdValue -ProviderName $providerNameValue
        $newTriggersXml += "`n" + $eventTriggerXml
    }
    
    if ($newTriggersXml -ne "") {
        Write-Verbose "Appending new trigger(s) to task '$Name'."
        # Get the default namespace from the current XML document.
        $defaultNs = $xmlDoc.DocumentElement.NamespaceURI
        # Wrap the new triggers XML in a temporary root element that declares the default namespace.
        $tempXml = [xml]("<NewTriggers xmlns='$defaultNs'>" + $newTriggersXml + "</NewTriggers>")
        $newTriggerNodes = $tempXml.NewTriggers.ChildNodes
        $triggersNode = $xmlDoc.Task.Triggers
        foreach ($node in $newTriggerNodes) {
            $importedNode = $xmlDoc.ImportNode($node, $true)
            $triggersNode.AppendChild($importedNode) | Out-Null
        }
    }
    
    # Convert the updated XML document back to a string.
    $updatedXml = $xmlDoc.OuterXml
    Write-Verbose "Re-registering task '$Name' with updated XML..."
    Register-CronTask -Name $Name -Xml $updatedXml
}



# === Example usage ===
# The following call creates a task named "TestIntervalTask" that uses the cron expression
# "*/3 */3 */4 * *" – meaning “at every 3rd minute (within the hour) during every 3rd hour on every 4th day-of-month.”
#
# To test, run:
#
#New-CronTask -Name "TestIntervalTask" -ScriptBlock { Write-Output "Running task" } -CronExpression "*/22 */12 * * *"
#Get-CronTask | Remove-CronTask
#New-CronTask
#Get-CronTask
#Get-RelativeTargetDate -Occurrence Last -DayOfWeek Thursday

#New-CronTask -Name "RelativeTask" -ScriptBlock { Write-Output "Relative schedule" } `
          #-CronExpression "* */2 * * *" -RelativeTarget -Occurrence Last -DayOfWeek Wednesday
#New-CronTask -Name "RelativeTask" -ScriptBlock { Write-Output "Relative schedule" } -EventLogTrigger -EventId 2 -ProviderName "MyProvider"

Export-ModuleMember -Function New-CronTask,Update-CronTask,Remove-CronTask,Get-CronTask,Trigger-CronEvent,Get-RelativeTargetDate
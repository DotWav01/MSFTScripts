#Requires -Version 5.1
<#
.SYNOPSIS
Runs PowerShell scripts continuously with flexible scheduling options.

.DESCRIPTION
This script provides a continuous execution wrapper for any PowerShell script.
It supports interval-based scheduling, specific days/times, and one-time execution.

.PARAMETER ScriptPath
Full path to the PowerShell script to execute (required)

.PARAMETER IntervalHours
Hours between executions when using interval mode (default: 1)

.PARAMETER IntervalMinutes
Minutes between executions when using interval mode (can be combined with IntervalHours)

.PARAMETER ScheduledDays
Array of days to run the script (Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday)
When specified, enables scheduled mode instead of interval mode

.PARAMETER ScheduledTimes
Array of times to run the script each scheduled day (24-hour format, e.g., "09:00", "14:30")
Required when using ScheduledDays

.PARAMETER LogPath
Path for the runner log file (optional, defaults to ContinuousRunnerLogs folder)

.PARAMETER ScriptParameters
Hashtable of parameters to pass to the target script

.PARAMETER RunOnce
Run the script once and exit (useful for testing)

.PARAMETER MaxLogFiles
Maximum number of log files to retain (default: 30)

.PARAMETER LogLevel
Logging level: INFO, WARN, ERROR, DEBUG (default: INFO)

.PARAMETER TimeZone
Time zone for scheduled execution (default: local time zone)

.PARAMETER StopOnError
Stop the runner if the target script fails (default: false)

.EXAMPLE
# Basic usage - run every hour
.\ContinuousScriptRunner.ps1 -ScriptPath "C:\Scripts\MyScript.ps1"

.EXAMPLE
# Run every 2 hours with parameters
$params = @{
    UseInteractiveAuth = $true
    Verbose = $true
}
.\ContinuousScriptRunner.ps1 -ScriptPath "C:\Scripts\MyScript.ps1" -IntervalHours 2 -ScriptParameters $params

.EXAMPLE
# Run on weekdays at 9:00 AM and 2:00 PM
.\ContinuousScriptRunner.ps1 -ScriptPath "C:\Scripts\MyScript.ps1" -ScheduledDays @("Monday","Tuesday","Wednesday","Thursday","Friday") -ScheduledTimes @("09:00","14:00")

.EXAMPLE
# Run every 30 minutes with debug logging
.\ContinuousScriptRunner.ps1 -ScriptPath "C:\Scripts\MyScript.ps1" -IntervalMinutes 30 -LogLevel DEBUG

.EXAMPLE
# Test run (execute once)
.\ContinuousScriptRunner.ps1 -ScriptPath "C:\Scripts\MyScript.ps1" -RunOnce

.NOTES
This script runs indefinitely until manually stopped (Ctrl+C) or the PowerShell session ends.
For production use, consider using Windows Task Scheduler for more robust scheduling.

Author: Alexander
Version: 2.0
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="Full path to the PowerShell script to execute")]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$ScriptPath,
    
    [Parameter(Mandatory=$false, HelpMessage="Hours between executions (interval mode)")]
    [ValidateRange(0, 8760)] # 0 to 1 year in hours
    [int]$IntervalHours = 1,
    
    [Parameter(Mandatory=$false, HelpMessage="Minutes between executions (interval mode)")]
    [ValidateRange(0, 59)]
    [int]$IntervalMinutes = 0,
    
    [Parameter(Mandatory=$false, HelpMessage="Days of the week to run (scheduled mode)")]
    [ValidateSet("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")]
    [string[]]$ScheduledDays,
    
    [Parameter(Mandatory=$false, HelpMessage="Times to run each scheduled day (HH:mm format)")]
    [ValidatePattern("^([01]?[0-9]|2[0-3]):[0-5][0-9]$")]
    [string[]]$ScheduledTimes,
    
    [Parameter(Mandatory=$false, HelpMessage="Path for the runner log file")]
    [string]$LogPath,
    
    [Parameter(Mandatory=$false, HelpMessage="Parameters to pass to the target script")]
    [hashtable]$ScriptParameters = @{},
    
    [Parameter(Mandatory=$false, HelpMessage="Run the script once and exit")]
    [switch]$RunOnce,
    
    [Parameter(Mandatory=$false, HelpMessage="Maximum number of log files to retain")]
    [ValidateRange(1, 365)]
    [int]$MaxLogFiles = 30,
    
    [Parameter(Mandatory=$false, HelpMessage="Logging level")]
    [ValidateSet("INFO","WARN","ERROR","DEBUG")]
    [string]$LogLevel = "INFO",
    
    [Parameter(Mandatory=$false, HelpMessage="Time zone for scheduled execution")]
    [string]$TimeZone = [System.TimeZone]::CurrentTimeZone.StandardName,
    
    [Parameter(Mandatory=$false, HelpMessage="Stop runner on script error")]
    [switch]$StopOnError
)

# Global variables
$script:CancelPressed = $false
$script:LogLevels = @{
    "ERROR" = 1
    "WARN"  = 2
    "INFO"  = 3
    "DEBUG" = 4
}

#region Functions

function Initialize-Environment {
    <#
    .SYNOPSIS
    Initializes the runner environment and validates parameters.
    #>
    
    # Validate script exists
    if (!(Test-Path $ScriptPath)) {
        throw "Target script not found: $ScriptPath"
    }
    
    # Validate scheduled mode parameters
    if ($ScheduledDays -and !$ScheduledTimes) {
        throw "ScheduledTimes parameter is required when using ScheduledDays"
    }
    if ($ScheduledTimes -and !$ScheduledDays) {
        throw "ScheduledDays parameter is required when using ScheduledTimes"
    }
    
    # Validate interval
    if (!$ScheduledDays -and $IntervalHours -eq 0 -and $IntervalMinutes -eq 0) {
        throw "Must specify either interval (Hours/Minutes) or scheduled execution (Days/Times)"
    }
    
    # Set up logging
    Initialize-Logging
    
    # Set up signal handlers
    Initialize-SignalHandlers
}

function Initialize-Logging {
    <#
    .SYNOPSIS
    Initializes the logging system.
    #>
    
    if ([string]::IsNullOrEmpty($LogPath)) {
        $ScriptDirectory = Split-Path $MyInvocation.ScriptName
        $LogsFolder = Join-Path $ScriptDirectory "ContinuousRunnerLogs"
        if (!(Test-Path $LogsFolder)) {
            New-Item -ItemType Directory -Path $LogsFolder -Force | Out-Null
        }
        
        $ScriptName = [System.IO.Path]::GetFileNameWithoutExtension((Split-Path $ScriptPath -Leaf))
        $script:LogPath = Join-Path $LogsFolder "Runner_${ScriptName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    } else {
        $script:LogPath = $LogPath
        $LogDirectory = Split-Path $LogPath -Parent
        if (!(Test-Path $LogDirectory)) {
            New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        }
    }
    
    # Clean up old log files
    if ($MaxLogFiles -gt 0) {
        Cleanup-LogFiles
    }
}

function Initialize-SignalHandlers {
    <#
    .SYNOPSIS
    Sets up signal handlers for graceful shutdown.
    #>
    
    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        $global:script:CancelPressed = $true
        Write-Host "`nShutdown requested - stopping after current execution..." -ForegroundColor Yellow
    } | Out-Null
    
    # Handle Ctrl+C
    [Console]::TreatControlCAsInput = $false
    $null = Register-EngineEvent -SourceIdentifier "PowerShell.Exiting" -Action {
        $script:CancelPressed = $true
    }
}

function Write-RunnerLog {
    <#
    .SYNOPSIS
    Writes log entries with level filtering.
    #>
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    # Check if we should log this level
    if ($script:LogLevels[$Level] -gt $script:LogLevels[$LogLevel]) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Color-code console output
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "INFO"  { "White" }
        "DEBUG" { "Gray" }
        default { "White" }
    }
    
    Write-Host $logEntry -ForegroundColor $color
    
    try {
        Add-Content -Path $script:LogPath -Value $logEntry -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Failed to write to log file: $($_.Exception.Message)"
    }
}

function Cleanup-LogFiles {
    <#
    .SYNOPSIS
    Removes old log files based on MaxLogFiles setting.
    #>
    
    try {
        $LogDirectory = Split-Path $script:LogPath -Parent
        $LogFiles = Get-ChildItem -Path $LogDirectory -Filter "Runner_*.log" | 
                   Sort-Object LastWriteTime -Descending
        
        if ($LogFiles.Count -gt $MaxLogFiles) {
            $FilesToDelete = $LogFiles | Select-Object -Skip $MaxLogFiles
            foreach ($file in $FilesToDelete) {
                Remove-Item $file.FullName -Force
                Write-RunnerLog "Deleted old log file: $($file.Name)" "DEBUG"
            }
        }
    }
    catch {
        Write-RunnerLog "Failed to cleanup log files: $($_.Exception.Message)" "WARN"
    }
}

function Invoke-TargetScript {
    <#
    .SYNOPSIS
    Executes the target PowerShell script with error handling.
    #>
    
    Write-RunnerLog "=== Starting script execution ===" "INFO"
    
    try {
        $ScriptName = Split-Path $ScriptPath -Leaf
        Write-RunnerLog "Executing: $ScriptName" "INFO"
        Write-RunnerLog "PowerShell Version: $($PSVersionTable.PSVersion)" "DEBUG"
        Write-RunnerLog "Full Path: $ScriptPath" "DEBUG"
        
        if ($ScriptParameters.Count -gt 0) {
            Write-RunnerLog "Script parameters provided: $($ScriptParameters.Keys -join ', ')" "DEBUG"
            foreach ($key in $ScriptParameters.Keys) {
                Write-RunnerLog "  $key = $($ScriptParameters[$key])" "DEBUG"
            }
        }
        
        # Execute the script
        $ErrorActionPreference = "Continue"
        if ($ScriptParameters.Count -gt 0) {
            & $ScriptPath @ScriptParameters
        } else {
            & $ScriptPath
        }
        
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0 -or $exitCode -eq $null) {
            Write-RunnerLog "Script completed successfully" "INFO"
            return $true
        } else {
            Write-RunnerLog "Script completed with exit code: $exitCode" "WARN"
            return !$StopOnError
        }
    }
    catch {
        Write-RunnerLog "Error executing script: $($_.Exception.Message)" "ERROR"
        Write-RunnerLog "Stack trace: $($_.ScriptStackTrace)" "DEBUG"
        return !$StopOnError
    }
    finally {
        Write-RunnerLog "=== Script execution completed ===" "INFO"
    }
}

function Get-NextScheduledRun {
    <#
    .SYNOPSIS
    Calculates the next scheduled run time.
    #>
    
    $now = Get-Date
    $nextRuns = @()
    
    foreach ($day in $ScheduledDays) {
        foreach ($time in $ScheduledTimes) {
            $timeSpan = [TimeSpan]::Parse($time)
            
            # Find the next occurrence of this day
            $dayOfWeek = [DayOfWeek]$day
            $daysUntil = ($dayOfWeek - $now.DayOfWeek + 7) % 7
            
            $nextDate = $now.Date.AddDays($daysUntil).Add($timeSpan)
            
            # If it's today but the time has passed, move to next week
            if ($nextDate -le $now) {
                $nextDate = $nextDate.AddDays(7)
            }
            
            $nextRuns += $nextDate
        }
    }
    
    return ($nextRuns | Sort-Object)[0]
}

function Start-IntervalMode {
    <#
    .SYNOPSIS
    Runs the script in interval mode.
    #>
    
    $totalMinutes = ($IntervalHours * 60) + $IntervalMinutes
    Write-RunnerLog "Running in interval mode: $totalMinutes minutes" "INFO"
    
    do {
        $startTime = Get-Date
        Write-RunnerLog "Current time: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" "DEBUG"
        
        # Execute the script
        $success = Invoke-TargetScript
        if (!$success -and $StopOnError) {
            Write-RunnerLog "Stopping runner due to script error" "ERROR"
            break
        }
        
        # Exit if running once
        if ($RunOnce) {
            Write-RunnerLog "Single execution completed, exiting" "INFO"
            break
        }
        
        # Calculate next run time
        $nextRunTime = $startTime.AddMinutes($totalMinutes)
        Write-RunnerLog "Next execution scheduled for: $($nextRunTime.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO"
        
        # Sleep until next execution (check every minute for cancellation)
        while ((Get-Date) -lt $nextRunTime -and -not $script:CancelPressed) {
            Start-Sleep -Seconds 60
        }
        
        if ($script:CancelPressed) {
            Write-RunnerLog "Cancellation requested, exiting..." "WARN"
            break
        }
        
    } while (-not $script:CancelPressed)
}

function Start-ScheduledMode {
    <#
    .SYNOPSIS
    Runs the script in scheduled mode.
    #>
    
    Write-RunnerLog "Running in scheduled mode" "INFO"
    Write-RunnerLog "Scheduled days: $($ScheduledDays -join ', ')" "INFO"
    Write-RunnerLog "Scheduled times: $($ScheduledTimes -join ', ')" "INFO"
    
    do {
        $now = Get-Date
        $nextRun = Get-NextScheduledRun
        
        Write-RunnerLog "Current time: $($now.ToString('yyyy-MM-dd HH:mm:ss'))" "DEBUG"
        Write-RunnerLog "Next scheduled run: $($nextRun.ToString('yyyy-MM-dd HH:mm:ss dddd'))" "INFO"
        
        # If we're at or past the scheduled time, run now
        if ($now -ge $nextRun.AddMinutes(-1)) { # 1-minute grace period
            $success = Invoke-TargetScript
            if (!$success -and $StopOnError) {
                Write-RunnerLog "Stopping runner due to script error" "ERROR"
                break
            }
            
            # Exit if running once
            if ($RunOnce) {
                Write-RunnerLog "Single execution completed, exiting" "INFO"
                break
            }
            
            # Wait a bit to avoid immediate re-execution
            Start-Sleep -Seconds 120
            continue
        }
        
        # Sleep until next scheduled time (check every 5 minutes)
        while ((Get-Date) -lt $nextRun -and -not $script:CancelPressed) {
            Start-Sleep -Seconds 300 # 5 minutes
        }
        
        if ($script:CancelPressed) {
            Write-RunnerLog "Cancellation requested, exiting..." "WARN"
            break
        }
        
    } while (-not $script:CancelPressed)
}

#endregion

#region Main Execution

try {
    # Initialize environment
    Initialize-Environment
    
    Write-RunnerLog "=== CONTINUOUS SCRIPT RUNNER STARTED ===" "INFO"
    Write-RunnerLog "Target Script: $(Split-Path $ScriptPath -Leaf)" "INFO"
    Write-RunnerLog "Full Path: $ScriptPath" "DEBUG"
    Write-RunnerLog "Log file: $script:LogPath" "INFO"
    Write-RunnerLog "Log level: $LogLevel" "DEBUG"
    Write-RunnerLog "Run once mode: $RunOnce" "INFO"
    
    if ($ScriptParameters.Count -gt 0) {
        Write-RunnerLog "Script parameters:" "DEBUG"
        foreach ($key in $ScriptParameters.Keys) {
            Write-RunnerLog "  $key = $($ScriptParameters[$key])" "DEBUG"
        }
    }
    
    # Choose execution mode
    if ($ScheduledDays) {
        Start-ScheduledMode
    } else {
        Start-IntervalMode
    }
}
catch {
    Write-RunnerLog "Critical error in continuous runner: $($_.Exception.Message)" "ERROR"
    Write-RunnerLog "Stack trace: $($_.ScriptStackTrace)" "DEBUG"
    exit 1
}
finally {
    Write-RunnerLog "=== CONTINUOUS SCRIPT RUNNER STOPPED ===" "INFO"
    
    # Clean up event handlers
    Get-EventSubscriber | Unregister-Event
}

#endregion

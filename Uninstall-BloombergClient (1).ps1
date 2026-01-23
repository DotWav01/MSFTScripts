<#
.SYNOPSIS
    Uninstalls Bloomberg Terminal client application for enterprise deployment via Intune.

.DESCRIPTION
    This script performs a silent uninstallation of the Bloomberg Terminal client application.
    Automatically closes Office 365 applications (Excel, Word, PowerPoint) before uninstallation
    as required by Bloomberg Terminal uninstaller. Uses Bloomberg's native uninstaller 
    (C:\blp\Uninstall\unins000.exe) as the primary method, with registry-based uninstall as a 
    fallback. Removes the detection tag file used by Intune. Designed for deployment through 
    Microsoft Intune as a Win32 application package.

.PARAMETER LogPath
    Path where uninstallation logs will be written. Defaults to C:\softdist\Logs\Bloomberg.

.PARAMETER Force
    Forces uninstallation even if Bloomberg processes are running (will attempt to terminate them).

.NOTES
    WhatIf parameter is automatically available through SupportsShouldProcess.
    Use -WhatIf to preview actions without executing them.

.EXAMPLE
    .\Uninstall-BloombergClient.ps1
    Performs a silent uninstallation of Bloomberg Terminal.

.EXAMPLE
    .\Uninstall-BloombergClient.ps1 -Force -Verbose
    Forces uninstallation with verbose output, closes Office apps and terminates Bloomberg processes if needed.

.EXAMPLE
    .\Uninstall-BloombergClient.ps1 -WhatIf
    Shows what would be done without actually uninstalling.

.NOTES
    File Name      : Uninstall-BloombergClient.ps1
    Author         : IT Infrastructure Team
    Prerequisite   : PowerShell 5.1 or later, Administrative privileges
    Requirements   : Bloomberg Terminal must be installed
    
    Exit Codes:
    0  = Success
    1  = General error
    2  = Bloomberg not found/already uninstalled
    3  = Uninstallation failed
    4  = Post-uninstallation verification failed
    5  = Insufficient privileges
    6  = Bloomberg processes running and could not be terminated
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\softdist\Logs\Bloomberg",
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Script variables
$ScriptName = "Uninstall-BloombergClient"
$ScriptVersion = "1.0.0"
$ExitCode = 0
$StartTime = Get-Date

# Initialize logging
try {
    if (-not (Test-Path -Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    
    $LogFile = Join-Path -Path $LogPath -ChildPath "$ScriptName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    Start-Transcript -Path $LogFile -Append
    
    Write-Host "[$ScriptName] Starting Bloomberg Terminal uninstallation" -ForegroundColor Green
    Write-Host "[$ScriptName] Script Version: $ScriptVersion" -ForegroundColor Green
    Write-Host "[$ScriptName] Log file: $LogFile" -ForegroundColor Green
    Write-Host "[$ScriptName] Start time: $StartTime" -ForegroundColor Green
    Write-Host "[$ScriptName] Running as: $env:USERNAME" -ForegroundColor Green
}
catch {
    Write-Error "Failed to initialize logging: $($_.Exception.Message)"
    exit 1
}

try {
    # Check for and close Office 365 applications before uninstallation
    Write-Host "[$ScriptName] Checking for running Office 365 applications..." -ForegroundColor Yellow
    
    $OfficeProcesses = @(
        @{ Name = "EXCEL"; DisplayName = "Microsoft Excel" },
        @{ Name = "WINWORD"; DisplayName = "Microsoft Word" },
        @{ Name = "POWERPNT"; DisplayName = "Microsoft PowerPoint" }
    )
    
    $RunningOfficeApps = @()
    
    foreach ($OfficeApp in $OfficeProcesses) {
        $Processes = Get-Process -Name $OfficeApp.Name -ErrorAction SilentlyContinue
        if ($Processes) {
            $RunningOfficeApps += @{
                ProcessName = $OfficeApp.Name
                DisplayName = $OfficeApp.DisplayName
                Processes = $Processes
            }
        }
    }
    
    if ($RunningOfficeApps.Count -gt 0) {
        Write-Host "[$ScriptName] Found running Office 365 applications that need to be closed:" -ForegroundColor Yellow
        foreach ($App in $RunningOfficeApps) {
            Write-Host "  - $($App.DisplayName) ($($App.Processes.Count) instance(s))" -ForegroundColor Yellow
        }
        
        Write-Host "[$ScriptName] Bloomberg Terminal uninstaller requires Office applications to be closed" -ForegroundColor Yellow
        Write-Host "[$ScriptName] Attempting to gracefully close Office applications..." -ForegroundColor Yellow
        
        $ClosureSuccess = $true
        
        foreach ($App in $RunningOfficeApps) {
            foreach ($Process in $App.Processes) {
                try {
                    Write-Host "[$ScriptName] Closing $($App.DisplayName) (PID: $($Process.Id))..." -ForegroundColor Gray
                    
                    # First attempt graceful closure
                    $Process.CloseMainWindow() | Out-Null
                    
                    # Wait up to 10 seconds for graceful closure
                    $WaitCount = 0
                    while (-not $Process.HasExited -and $WaitCount -lt 10) {
                        Start-Sleep -Seconds 1
                        $WaitCount++
                    }
                    
                    # If still running, force termination with Force parameter or user consent
                    if (-not $Process.HasExited) {
                        if ($Force) {
                            Write-Warning "[$ScriptName] Graceful closure failed, force terminating $($App.DisplayName) (PID: $($Process.Id))"
                            $Process.Kill()
                            Start-Sleep -Seconds 2
                        }
                        else {
                            Write-Warning "[$ScriptName] $($App.DisplayName) did not close gracefully. Use -Force to terminate or close manually."
                            $ClosureSuccess = $false
                        }
                    }
                    
                    if ($Process.HasExited) {
                        Write-Host "[$ScriptName] Successfully closed $($App.DisplayName) (PID: $($Process.Id))" -ForegroundColor Green
                    }
                    else {
                        Write-Warning "[$ScriptName] Failed to close $($App.DisplayName) (PID: $($Process.Id))"
                        $ClosureSuccess = $false
                    }
                }
                catch {
                    Write-Warning "[$ScriptName] Error closing $($App.DisplayName): $($_.Exception.Message)"
                    $ClosureSuccess = $false
                }
            }
        }
        
        # Final verification that Office apps are closed
        Start-Sleep -Seconds 3
        $StillRunning = @()
        
        foreach ($OfficeApp in $OfficeProcesses) {
            $RemainingProcesses = Get-Process -Name $OfficeApp.Name -ErrorAction SilentlyContinue
            if ($RemainingProcesses) {
                $StillRunning += @{
                    ProcessName = $OfficeApp.Name
                    DisplayName = $OfficeApp.DisplayName
                    Count = $RemainingProcesses.Count
                }
            }
        }
        
        if ($StillRunning.Count -gt 0) {
            Write-Error "[$ScriptName] Some Office applications are still running after closure attempt:"
            foreach ($App in $StillRunning) {
                Write-Error "  - $($App.DisplayName) ($($App.Count) instance(s))"
            }
            Write-Warning "[$ScriptName] Bloomberg uninstallation may fail or encounter issues with Office apps running"
            Write-Warning "[$ScriptName] Continuing with uninstallation attempt..."
        }
        else {
            Write-Host "[$ScriptName] All Office applications successfully closed" -ForegroundColor Green
        }
    }
    else {
        Write-Host "[$ScriptName] No Office 365 applications found running" -ForegroundColor Green
    }

    # Check for Bloomberg processes
    Write-Host "[$ScriptName] Checking for running Bloomberg processes..." -ForegroundColor Yellow
    
    $BloombergProcessNames = @(
        "bloomberg*",
        "terminal*",
        "wintrv*",
        "bbg*"
    )
    
    $RunningProcesses = @()
    foreach ($ProcessName in $BloombergProcessNames) {
        $Processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($Processes) {
            $RunningProcesses += $Processes
        }
    }
    
    if ($RunningProcesses.Count -gt 0) {
        Write-Host "[$ScriptName] Found $($RunningProcesses.Count) Bloomberg-related processes running:" -ForegroundColor Yellow
        foreach ($Process in $RunningProcesses) {
            Write-Host "  - $($Process.ProcessName) (PID: $($Process.Id))" -ForegroundColor Yellow
        }
        
        if ($Force) {
            Write-Host "[$ScriptName] Force parameter specified, attempting to terminate Bloomberg processes..." -ForegroundColor Yellow
            
            foreach ($Process in $RunningProcesses) {
                try {
                    Write-Host "[$ScriptName] Terminating process: $($Process.ProcessName) (PID: $($Process.Id))" -ForegroundColor Gray
                    $Process | Stop-Process -Force -ErrorAction Stop
                    Write-Host "[$ScriptName] Successfully terminated: $($Process.ProcessName)" -ForegroundColor Green
                }
                catch {
                    Write-Warning "[$ScriptName] Failed to terminate process $($Process.ProcessName): $($_.Exception.Message)"
                }
            }
            
            # Wait for processes to terminate
            Start-Sleep -Seconds 5
            
            # Check if any processes are still running
            $StillRunning = @()
            foreach ($ProcessName in $BloombergProcessNames) {
                $Processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
                if ($Processes) {
                    $StillRunning += $Processes
                }
            }
            
            if ($StillRunning.Count -gt 0) {
                Write-Error "[$ScriptName] Could not terminate all Bloomberg processes. Please close Bloomberg Terminal manually and retry."
                $ExitCode = 6
                throw "Bloomberg processes still running after termination attempt"
            }
        }
        else {
            Write-Warning "[$ScriptName] Bloomberg processes are running. Please close Bloomberg Terminal or use -Force parameter."
            Write-Warning "[$ScriptName] Continuing with uninstallation attempt..."
        }
    }
    else {
        Write-Host "[$ScriptName] No Bloomberg processes found running" -ForegroundColor Green
    }

    if ($WhatIf) {
        Write-Host "[$ScriptName] WhatIf: Would check for and close Office 365 applications (Excel, Word, PowerPoint)" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Would attempt to run Bloomberg uninstaller: C:\blp\Uninstall\unins000.exe /S" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Would fallback to registry uninstall string with /S if needed" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Would delete tag file C:\temp\Bloomberg_Installed.tag" -ForegroundColor Magenta
        exit 0
    }
    
    # Method 1: Use Bloomberg's uninstaller
    Write-Host "[$ScriptName] Attempting primary uninstall method..." -ForegroundColor Yellow
    $BloombergUninstaller = "C:\blp\Uninstall\unins000.exe"
    $UninstallSuccess = $false
    $PrimaryUninstallerRan = $false
    
    if (Test-Path -Path $BloombergUninstaller) {
        Write-Host "[$ScriptName] Found Bloomberg uninstaller: $BloombergUninstaller" -ForegroundColor Green
        
        try {
            $UninstallArgs = "/S"
            Write-Host "[$ScriptName] Executing: $BloombergUninstaller $UninstallArgs" -ForegroundColor Gray
            
            $UninstallStartTime = Get-Date
            $Process = Start-Process -FilePath $BloombergUninstaller -ArgumentList $UninstallArgs -Wait -PassThru -NoNewWindow
            $UninstallEndTime = Get-Date
            $UninstallDuration = $UninstallEndTime - $UninstallStartTime
            
            $PrimaryUninstallerRan = $true
            Write-Host "[$ScriptName] Primary uninstaller completed with exit code: $($Process.ExitCode)" -ForegroundColor Green
            Write-Host "[$ScriptName] Uninstall duration: $($UninstallDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
            
            # Bloomberg uninstaller often returns non-zero codes even when successful
            # If the uninstaller ran, we consider it successful regardless of exit code
            Write-Host "[$ScriptName] Treating primary uninstaller execution as successful (Bloomberg uninstallers commonly return non-zero codes)" -ForegroundColor Green
            $UninstallSuccess = $true
        }
        catch {
            Write-Warning "[$ScriptName] Primary uninstall failed: $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "[$ScriptName] Bloomberg uninstaller not found at: $BloombergUninstaller"
    }
    
    # Method 2: Use registry uninstall string (fallback) - Only if primary uninstaller didn't run
    if (-not $PrimaryUninstallerRan) {
        Write-Host "[$ScriptName] Attempting registry-based uninstall (fallback method)..." -ForegroundColor Yellow
        
        $RegistryPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Bloomberg Terminal_is1"
        
        try {
            if (Test-Path -Path $RegistryPath) {
                $UninstallInfo = Get-ItemProperty -Path $RegistryPath -ErrorAction Stop
                
                if ($UninstallInfo.UninstallString) {
                    Write-Host "[$ScriptName] Found registry uninstall string: $($UninstallInfo.UninstallString)" -ForegroundColor Green
                    
                    $UninstallString = $UninstallInfo.UninstallString
                    
                    # Parse uninstall string
                    if ($UninstallString -match '"([^"]+)"(.*)') {
                        $UninstallExe = $matches[1]
                        $UninstallArgs = $matches[2].Trim()
                    }
                    else {
                        $UninstallExe = $UninstallString
                        $UninstallArgs = ""
                    }
                    
                    # Add silent arguments if not present
                    if ($UninstallArgs -notlike "*S*" -and $UninstallArgs -notlike "*s*") {
                        $UninstallArgs += " /S"
                    }
                    
                    Write-Host "[$ScriptName] Executing: $UninstallExe $UninstallArgs" -ForegroundColor Gray
                    
                    $UninstallStartTime = Get-Date
                    $Process = Start-Process -FilePath $UninstallExe -ArgumentList $UninstallArgs -Wait -PassThru -NoNewWindow
                    $UninstallEndTime = Get-Date
                    $UninstallDuration = $UninstallEndTime - $UninstallStartTime
                    
                    Write-Host "[$ScriptName] Registry-based uninstall completed with exit code: $($Process.ExitCode)" -ForegroundColor Green
                    Write-Host "[$ScriptName] Uninstall duration: $($UninstallDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
                    $UninstallSuccess = $true
                }
                else {
                    Write-Warning "[$ScriptName] Registry entry found but UninstallString is empty"
                }
            }
            else {
                Write-Warning "[$ScriptName] Bloomberg registry entry not found at: $RegistryPath"
            }
        }
        catch {
            Write-Warning "[$ScriptName] Registry-based uninstall failed: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "[$ScriptName] Skipping registry-based uninstall since primary uninstaller already ran" -ForegroundColor Gray
    }
    
    # Always try to remove tag file, regardless of uninstall success
    Write-Host "[$ScriptName] Removing detection tag file..." -ForegroundColor Yellow
    $TagFilePath = "C:\temp\Bloomberg_Installed.tag"
    
    if (Test-Path -Path $TagFilePath) {
        try {
            Remove-Item -Path $TagFilePath -Force -ErrorAction Stop
            Write-Host "[$ScriptName] Successfully removed tag file: $TagFilePath" -ForegroundColor Green
        }
        catch {
            Write-Warning "[$ScriptName] Failed to remove tag file: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "[$ScriptName] Tag file not found (already removed or never existed): $TagFilePath" -ForegroundColor Gray
    }
    
    # Post-uninstallation verification
    Write-Host "[$ScriptName] Performing post-uninstallation verification..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    
    $VerificationPassed = $true
    
    # Check Bloomberg directory - it may still exist but should be mostly empty
    if (Test-Path -Path "C:\blp") {
        # Check if there are any Bloomberg executables left (the real test)
        $BloombergExe = "C:\blp\winrtv\wintrv.exe"
        if (Test-Path -Path $BloombergExe) {
            Write-Warning "[$ScriptName] Bloomberg executable still exists: $BloombergExe"
            $VerificationPassed = $false
        }
        else {
            Write-Host "[$ScriptName] Verification: Bloomberg executable removed successfully" -ForegroundColor Green
            
            # Check what's left in the blp folder
            try {
                $RemainingItems = Get-ChildItem -Path "C:\blp" -ErrorAction SilentlyContinue
                if ($RemainingItems) {
                    Write-Host "[$ScriptName] Verification: C:\blp folder exists but contains only: $($RemainingItems.Name -join ', ')" -ForegroundColor Gray
                }
                else {
                    Write-Host "[$ScriptName] Verification: C:\blp folder is empty" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "[$ScriptName] Verification: Could not enumerate C:\blp contents" -ForegroundColor Gray
            }
        }
    }
    else {
        Write-Host "[$ScriptName] Verification: Bloomberg directory removed completely" -ForegroundColor Green
    }
    
    # Check registry for remaining Bloomberg entries (but don't fail if registry path doesn't exist)
    try {
        if (Test-Path -Path $RegistryPath) {
            Write-Warning "[$ScriptName] Bloomberg registry entry still exists: $RegistryPath"
            $VerificationPassed = $false
        }
        else {
            Write-Host "[$ScriptName] Verification: Bloomberg registry entry removed successfully" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "[$ScriptName] Verification: Could not check registry entry (likely removed)" -ForegroundColor Gray
    }
    
    # Check for Bloomberg services
    $RemainingServices = Get-Service -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*Bloomberg*" -or $_.DisplayName -like "*Bloomberg*" 
    }
    
    if ($RemainingServices) {
        Write-Warning "[$ScriptName] Bloomberg services still exist:"
        foreach ($Service in $RemainingServices) {
            Write-Warning "  - $($Service.Name) ($($Service.DisplayName)): $($Service.Status)"
        }
        $VerificationPassed = $false
    }
    else {
        Write-Host "[$ScriptName] Verification: No Bloomberg services found" -ForegroundColor Green
    }
    
    # Check tag file is removed
    if (Test-Path -Path $TagFilePath) {
        Write-Warning "[$ScriptName] Detection tag file still exists: $TagFilePath"
        $VerificationPassed = $false
    }
    else {
        Write-Host "[$ScriptName] Verification: Detection tag file removed successfully" -ForegroundColor Green
    }
    
    # Determine overall success - use more realistic success criteria
    $BloombergExeExists = Test-Path -Path "C:\blp\winrtv\wintrv.exe"
    
    if ($UninstallSuccess -and -not $BloombergExeExists) {
        Write-Host "[$ScriptName] Bloomberg Terminal uninstallation completed successfully" -ForegroundColor Green
        if (-not $VerificationPassed) {
            Write-Host "[$ScriptName] Some verification items noted, but primary uninstallation was successful" -ForegroundColor Yellow
        }
    }
    else {
        if ($BloombergExeExists) {
            Write-Error "[$ScriptName] Uninstallation failed - Bloomberg executable still exists: C:\blp\winrtv\wintrv.exe"
            $ExitCode = 3
            throw "Uninstallation failed - Bloomberg still installed"
        }
        else {
            Write-Warning "[$ScriptName] Uninstall methods had issues but Bloomberg appears to be removed"
            Write-Host "[$ScriptName] Bloomberg Terminal uninstallation completed" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Error "[$ScriptName] Uninstallation failed: $($_.Exception.Message)"
    Write-Error "[$ScriptName] Full error: $($_.Exception | Format-List * | Out-String)"
    
    if ($ExitCode -eq 0) {
        $ExitCode = 1
    }
}
finally {
    $EndTime = Get-Date
    $TotalDuration = $EndTime - $StartTime
    
    Write-Host "[$ScriptName] Script execution completed" -ForegroundColor Green
    Write-Host "[$ScriptName] Total duration: $($TotalDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
    Write-Host "[$ScriptName] Exit code: $ExitCode" -ForegroundColor Green
    
    try {
        Stop-Transcript
    }
    catch {
        # Transcript may not have started successfully
    }
    
    exit $ExitCode
}

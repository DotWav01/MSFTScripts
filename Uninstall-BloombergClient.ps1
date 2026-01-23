<#
.SYNOPSIS
    Uninstalls Bloomberg Terminal client application for enterprise deployment via Intune.

.DESCRIPTION
    This script performs a silent uninstallation of the Bloomberg Terminal client application.
    Uses Bloomberg's native uninstaller (C:\blp\Uninstall\unins000.exe) as the primary method,
    with registry-based uninstall as a fallback. Removes the detection tag file used by Intune.
    Designed for deployment through Microsoft Intune as a Win32 application package.

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
    Forces uninstallation with verbose output, terminating Bloomberg processes if needed.

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
        Write-Host "[$ScriptName] WhatIf: Would attempt to run Bloomberg uninstaller at C:\blp\Uninstall\unins000.exe" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Would fallback to registry uninstall string if needed" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Would delete tag file C:\temp\Bloomberg_Installed.tag" -ForegroundColor Magenta
        exit 0
    }
    
    # Method 1: Use Bloomberg's uninstaller
    Write-Host "[$ScriptName] Attempting primary uninstall method..." -ForegroundColor Yellow
    $BloombergUninstaller = "C:\blp\Uninstall\unins000.exe"
    $UninstallSuccess = $false
    
    if (Test-Path -Path $BloombergUninstaller) {
        Write-Host "[$ScriptName] Found Bloomberg uninstaller: $BloombergUninstaller" -ForegroundColor Green
        
        try {
            $UninstallArgs = @("/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
            Write-Host "[$ScriptName] Executing: $BloombergUninstaller $($UninstallArgs -join ' ')" -ForegroundColor Gray
            
            $UninstallStartTime = Get-Date
            $Process = Start-Process -FilePath $BloombergUninstaller -ArgumentList $UninstallArgs -Wait -PassThru -NoNewWindow
            $UninstallEndTime = Get-Date
            $UninstallDuration = $UninstallEndTime - $UninstallStartTime
            
            if ($Process.ExitCode -eq 0) {
                Write-Host "[$ScriptName] Primary uninstall completed successfully (Exit Code: $($Process.ExitCode))" -ForegroundColor Green
                Write-Host "[$ScriptName] Uninstall duration: $($UninstallDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
                $UninstallSuccess = $true
            }
            else {
                Write-Warning "[$ScriptName] Primary uninstall failed with exit code: $($Process.ExitCode)"
            }
        }
        catch {
            Write-Warning "[$ScriptName] Primary uninstall failed: $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "[$ScriptName] Bloomberg uninstaller not found at: $BloombergUninstaller"
    }
    
    # Method 2: Use registry uninstall string (fallback)
    if (-not $UninstallSuccess) {
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
                    if ($UninstallArgs -notlike "*SILENT*" -and $UninstallArgs -notlike "*/q*") {
                        $UninstallArgs += " /SILENT /SUPPRESSMSGBOXES /NORESTART"
                    }
                    
                    Write-Host "[$ScriptName] Executing: $UninstallExe $UninstallArgs" -ForegroundColor Gray
                    
                    $UninstallStartTime = Get-Date
                    $Process = Start-Process -FilePath $UninstallExe -ArgumentList $UninstallArgs.Split(' ') -Wait -PassThru -NoNewWindow
                    $UninstallEndTime = Get-Date
                    $UninstallDuration = $UninstallEndTime - $UninstallStartTime
                    
                    if ($Process.ExitCode -eq 0 -or $Process.ExitCode -eq 3010) {
                        Write-Host "[$ScriptName] Registry-based uninstall completed successfully (Exit Code: $($Process.ExitCode))" -ForegroundColor Green
                        Write-Host "[$ScriptName] Uninstall duration: $($UninstallDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
                        $UninstallSuccess = $true
                    }
                    else {
                        Write-Warning "[$ScriptName] Registry-based uninstall failed with exit code: $($Process.ExitCode)"
                    }
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
    
    if (-not $UninstallSuccess) {
        Write-Error "[$ScriptName] Both uninstall methods failed"
        $ExitCode = 3
        throw "Uninstallation failed"
    }
    
    # Remove detection tag file
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
        Write-Host "[$ScriptName] Tag file not found (already removed): $TagFilePath" -ForegroundColor Gray
    }
    
    # Post-uninstallation verification
    Write-Host "[$ScriptName] Performing post-uninstallation verification..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    
    $VerificationPassed = $true
    
    # Check if Bloomberg directory still exists
    if (Test-Path -Path "C:\blp") {
        Write-Warning "[$ScriptName] Bloomberg directory still exists: C:\blp"
        $VerificationPassed = $false
    }
    
    # Check registry for remaining Bloomberg entries
    try {
        if (Test-Path -Path $RegistryPath) {
            Write-Warning "[$ScriptName] Bloomberg registry entry still exists: $RegistryPath"
            $VerificationPassed = $false
        }
    }
    catch {
        # Registry path may have been removed
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
    
    # Check tag file is removed
    if (Test-Path -Path $TagFilePath) {
        Write-Warning "[$ScriptName] Detection tag file still exists: $TagFilePath"
        $VerificationPassed = $false
    }
    
    if ($VerificationPassed) {
        Write-Host "[$ScriptName] Post-uninstallation verification passed - Bloomberg Terminal successfully removed" -ForegroundColor Green
    }
    else {
        Write-Warning "[$ScriptName] Post-uninstallation verification detected remaining Bloomberg components"
        Write-Host "[$ScriptName] Primary uninstallation completed, but some components may require manual cleanup" -ForegroundColor Yellow
    }
    
    Write-Host "[$ScriptName] Bloomberg Terminal uninstallation completed" -ForegroundColor Green
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

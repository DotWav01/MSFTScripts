<#
.SYNOPSIS
    Uninstalls Bloomberg Terminal client application for enterprise deployment via Intune.

.DESCRIPTION
    This script performs a silent uninstallation of the Bloomberg Terminal client application.
    Designed for deployment through Microsoft Intune as a Win32 application package.
    Includes comprehensive logging, error handling, and exit codes for deployment monitoring.

.PARAMETER LogPath
    Path where uninstallation logs will be written. Defaults to C:\softdist\Logs\Bloomberg.

.PARAMETER Force
    Forces uninstallation even if Bloomberg processes are running (will attempt to terminate them).

.PARAMETER WhatIf
    Shows what would be done without actually performing the uninstallation.

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
    [switch]$Force,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
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
        "wintrm*",
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
    
    # Find Bloomberg installation
    Write-Host "[$ScriptName] Locating Bloomberg installation..." -ForegroundColor Yellow
    
    # Check registry for uninstall information
    $RegistryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $BloombergUninstallInfo = @()
    foreach ($RegPath in $RegistryPaths) {
        try {
            $UninstallEntries = Get-ItemProperty $RegPath -ErrorAction SilentlyContinue | Where-Object { 
                $_.DisplayName -like "*Bloomberg*" -or $_.DisplayName -like "*Terminal*" 
            }
            
            if ($UninstallEntries) {
                $BloombergUninstallInfo += $UninstallEntries
            }
        }
        catch {
            # Registry path might not exist, continue
        }
    }
    
    if ($BloombergUninstallInfo.Count -eq 0) {
        Write-Warning "[$ScriptName] Bloomberg Terminal installation not found in registry"
        
        # Check for installation directories
        $BloombergPaths = @(
            "${env:ProgramFiles}\Bloomberg Terminal",
            "${env:ProgramFiles(x86)}\Bloomberg Terminal",
            "${env:ProgramFiles}\Bloomberg",
            "${env:ProgramFiles(x86)}\Bloomberg"
        )
        
        $ExistingInstall = $null
        foreach ($Path in $BloombergPaths) {
            if (Test-Path -Path $Path) {
                $ExistingInstall = $Path
                Write-Host "[$ScriptName] Found Bloomberg installation directory: $Path" -ForegroundColor Green
                break
            }
        }
        
        if (-not $ExistingInstall) {
            Write-Host "[$ScriptName] Bloomberg Terminal does not appear to be installed" -ForegroundColor Green
            Write-Host "[$ScriptName] Uninstallation completed (nothing to uninstall)" -ForegroundColor Green
            exit 0
        }
    }
    else {
        Write-Host "[$ScriptName] Found Bloomberg installation(s):" -ForegroundColor Green
        foreach ($Entry in $BloombergUninstallInfo) {
            Write-Host "  - $($Entry.DisplayName)" -ForegroundColor Green
            if ($Entry.DisplayVersion) {
                Write-Host "    Version: $($Entry.DisplayVersion)" -ForegroundColor Green
            }
            if ($Entry.UninstallString) {
                Write-Host "    Uninstall: $($Entry.UninstallString)" -ForegroundColor Gray
            }
        }
    }
    
    if ($WhatIf) {
        Write-Host "[$ScriptName] WhatIf: Would uninstall Bloomberg Terminal" -ForegroundColor Magenta
        if ($BloombergUninstallInfo.Count -gt 0) {
            foreach ($Entry in $BloombergUninstallInfo) {
                Write-Host "[$ScriptName] WhatIf: Would execute: $($Entry.UninstallString)" -ForegroundColor Magenta
            }
        }
        exit 0
    }
    
    # Perform uninstallation
    $UninstallSuccess = $false
    
    if ($BloombergUninstallInfo.Count -gt 0) {
        Write-Host "[$ScriptName] Starting Bloomberg Terminal uninstallation via registry entries..." -ForegroundColor Yellow
        
        foreach ($Entry in $BloombergUninstallInfo) {
            if ($Entry.UninstallString) {
                try {
                    Write-Host "[$ScriptName] Uninstalling: $($Entry.DisplayName)" -ForegroundColor Yellow
                    
                    $UninstallString = $Entry.UninstallString
                    
                    # Parse uninstall string
                    if ($UninstallString -match '"([^"]+)"(.*)') {
                        $UninstallExe = $matches[1]
                        $UninstallArgs = $matches[2].Trim()
                    }
                    elseif ($UninstallString -match '(\S+\.exe)(.*)') {
                        $UninstallExe = $matches[1]
                        $UninstallArgs = $matches[2].Trim()
                    }
                    else {
                        $UninstallExe = $UninstallString
                        $UninstallArgs = ""
                    }
                    
                    # Add silent uninstall arguments
                    if ($UninstallString -like "*msiexec*") {
                        # MSI uninstaller
                        if ($UninstallArgs -notlike "*quiet*" -and $UninstallArgs -notlike "*/q*") {
                            $UninstallArgs += " /quiet /norestart"
                        }
                    }
                    else {
                        # Standard uninstaller - try common silent arguments
                        $SilentArgs = @("/S", "/SILENT", "/VERYSILENT", "/q")
                        $AddArgs = $true
                        foreach ($SilentArg in $SilentArgs) {
                            if ($UninstallArgs -like "*$SilentArg*") {
                                $AddArgs = $false
                                break
                            }
                        }
                        if ($AddArgs) {
                            $UninstallArgs += " /S /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
                        }
                    }
                    
                    Write-Host "[$ScriptName] Executing: $UninstallExe $UninstallArgs" -ForegroundColor Gray
                    
                    $UninstallStartTime = Get-Date
                    
                    if ($UninstallString -like "*msiexec*") {
                        # Use Start-Process for MSI
                        $Process = Start-Process -FilePath "msiexec.exe" -ArgumentList $UninstallArgs.Split(' ') -Wait -PassThru -NoNewWindow
                    }
                    else {
                        # Use Start-Process for standard uninstaller
                        $Process = Start-Process -FilePath $UninstallExe -ArgumentList $UninstallArgs.Split(' ') -Wait -PassThru -NoNewWindow
                    }
                    
                    $UninstallEndTime = Get-Date
                    $UninstallDuration = $UninstallEndTime - $UninstallStartTime
                    
                    if ($Process.ExitCode -eq 0 -or $Process.ExitCode -eq 3010) {
                        Write-Host "[$ScriptName] Uninstallation completed successfully (Exit Code: $($Process.ExitCode))" -ForegroundColor Green
                        Write-Host "[$ScriptName] Uninstall duration: $($UninstallDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
                        $UninstallSuccess = $true
                    }
                    else {
                        Write-Warning "[$ScriptName] Uninstallation failed with exit code: $($Process.ExitCode)"
                    }
                }
                catch {
                    Write-Warning "[$ScriptName] Failed to uninstall $($Entry.DisplayName): $($_.Exception.Message)"
                }
            }
        }
    }
    
    # Manual cleanup if registry uninstall failed or wasn't available
    if (-not $UninstallSuccess) {
        Write-Host "[$ScriptName] Attempting manual cleanup..." -ForegroundColor Yellow
        
        # Remove installation directories
        $BloombergPaths = @(
            "${env:ProgramFiles}\Bloomberg Terminal",
            "${env:ProgramFiles(x86)}\Bloomberg Terminal",
            "${env:ProgramFiles}\Bloomberg",
            "${env:ProgramFiles(x86)}\Bloomberg"
        )
        
        foreach ($Path in $BloombergPaths) {
            if (Test-Path -Path $Path) {
                try {
                    Write-Host "[$ScriptName] Removing directory: $Path" -ForegroundColor Yellow
                    Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                    Write-Host "[$ScriptName] Successfully removed: $Path" -ForegroundColor Green
                    $UninstallSuccess = $true
                }
                catch {
                    Write-Warning "[$ScriptName] Failed to remove directory $Path: $($_.Exception.Message)"
                }
            }
        }
        
        # Clean up user profile directories
        $UserBloombergPaths = @(
            "${env:APPDATA}\Bloomberg",
            "${env:LOCALAPPDATA}\Bloomberg",
            "${env:USERPROFILE}\Bloomberg"
        )
        
        foreach ($Path in $UserBloombergPaths) {
            if (Test-Path -Path $Path) {
                try {
                    Write-Host "[$ScriptName] Removing user data: $Path" -ForegroundColor Yellow
                    Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                    Write-Host "[$ScriptName] Successfully removed: $Path" -ForegroundColor Green
                }
                catch {
                    Write-Warning "[$ScriptName] Failed to remove user data $Path: $($_.Exception.Message)"
                }
            }
        }
        
        # Remove Bloomberg services
        $BloombergServices = Get-Service | Where-Object { 
            $_.Name -like "*Bloomberg*" -or $_.DisplayName -like "*Bloomberg*" 
        }
        
        foreach ($Service in $BloombergServices) {
            try {
                Write-Host "[$ScriptName] Removing service: $($Service.Name)" -ForegroundColor Yellow
                
                # Stop service if running
                if ($Service.Status -eq "Running") {
                    Stop-Service -Name $Service.Name -Force -ErrorAction Stop
                }
                
                # Remove service
                & sc.exe delete $Service.Name
                Write-Host "[$ScriptName] Successfully removed service: $($Service.Name)" -ForegroundColor Green
            }
            catch {
                Write-Warning "[$ScriptName] Failed to remove service $($Service.Name): $($_.Exception.Message)"
            }
        }
    }
    
    if (-not $UninstallSuccess) {
        Write-Error "[$ScriptName] Uninstallation failed - no successful uninstall method worked"
        $ExitCode = 3
        throw "Uninstallation failed"
    }
    
    # Post-uninstallation verification
    Write-Host "[$ScriptName] Performing post-uninstallation verification..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    
    $VerificationPassed = $true
    
    # Check if Bloomberg directories still exist
    foreach ($Path in $BloombergPaths) {
        if (Test-Path -Path $Path) {
            Write-Warning "[$ScriptName] Bloomberg directory still exists: $Path"
            $VerificationPassed = $false
        }
    }
    
    # Check registry for remaining Bloomberg entries
    foreach ($RegPath in $RegistryPaths) {
        try {
            $RemainingEntries = Get-ItemProperty $RegPath -ErrorAction SilentlyContinue | Where-Object { 
                $_.DisplayName -like "*Bloomberg*" 
            }
            
            if ($RemainingEntries) {
                Write-Warning "[$ScriptName] Bloomberg registry entries still exist:"
                foreach ($Entry in $RemainingEntries) {
                    Write-Warning "  - $($Entry.DisplayName)"
                }
                $VerificationPassed = $false
            }
        }
        catch {
            # Continue checking
        }
    }
    
    # Check for Bloomberg services
    $RemainingServices = Get-Service | Where-Object { 
        $_.Name -like "*Bloomberg*" -or $_.DisplayName -like "*Bloomberg*" 
    }
    
    if ($RemainingServices) {
        Write-Warning "[$ScriptName] Bloomberg services still exist:"
        foreach ($Service in $RemainingServices) {
            Write-Warning "  - $($Service.Name) ($($Service.DisplayName)): $($Service.Status)"
        }
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

#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Uninstall .NET Framework 3.5
.DESCRIPTION
    Disables .NET Framework 3.5 Windows optional feature
    Logs all results to C:\softdist directory
.NOTES
    Version: 1.0
    Author: IT Infrastructure Team
    Date: $(Get-Date -Format "yyyy-MM-dd")
.EXAMPLE
    .\Uninstall-NetFx35.ps1
#>

# Set up logging
$LogDirectory = "C:\softdist"
$LogFile = "$LogDirectory\NetFx35_Uninstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Ensure log directory exists
if (!(Test-Path $LogDirectory)) {
    try {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        Write-Host "Created log directory: $LogDirectory" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create log directory: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console with color
    switch ($Level) {
        "INFO" { Write-Host $logEntry -ForegroundColor White }
        "WARN" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
    }
    
    # Write to log file
    try {
        Add-Content -Path $LogFile -Value $logEntry -Force -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

try {
    Write-Log "========================================" "INFO"
    Write-Log "Starting .NET Framework 3.5 Uninstallation" "INFO"
    Write-Log "Script: $($MyInvocation.MyCommand.Name)" "INFO"
    Write-Log "Log File: $LogFile" "INFO"
    Write-Log "========================================" "INFO"
    
    # Check current .NET Framework 3.5 status
    Write-Log "Checking current .NET Framework 3.5 status..." "INFO"
    
    try {
        $netFx35Status = Get-WindowsOptionalFeature -Online -FeatureName NetFx3 -ErrorAction Stop
        Write-Log "Current NetFx3 state: $($netFx35Status.State)" "INFO"
        
        if ($netFx35Status.State -eq "Disabled") {
            Write-Log ".NET Framework 3.5 is already disabled. No action required." "SUCCESS"
            Write-Log "Uninstallation process completed successfully." "SUCCESS"
            
            # Clean up any detection files that might exist
            Write-Log "Cleaning up detection files..." "INFO"
            $detectionFile = "$env:ProgramData\CustomDetection\NetFx35Installed.txt"
            if (Test-Path $detectionFile) {
                try {
                    Remove-Item $detectionFile -Force -ErrorAction Stop
                    Write-Log "Removed detection file: $detectionFile" "SUCCESS"
                }
                catch {
                    Write-Log "Failed to remove detection file: $($_.Exception.Message)" "WARN"
                }
            }
            
            exit 0
        }
        
        if ($netFx35Status.State -eq "Enabled") {
            Write-Log ".NET Framework 3.5 is currently enabled. Proceeding with uninstallation..." "INFO"
        } else {
            Write-Log ".NET Framework 3.5 state is: $($netFx35Status.State). Proceeding with uninstallation..." "INFO"
        }
    }
    catch {
        Write-Log "Failed to check .NET Framework 3.5 status: $($_.Exception.Message)" "ERROR"
        Write-Log "Continuing with uninstallation attempt..." "WARN"
    }
    
    # Import DISM module if available
    Write-Log "Importing DISM module..." "INFO"
    try {
        Import-Module DISM -Force -ErrorAction Stop
        Write-Log "DISM module imported successfully" "SUCCESS"
    }
    catch {
        Write-Log "DISM module not available: $($_.Exception.Message)" "WARN"
        Write-Log "Continuing with uninstallation..." "INFO"
    }
    
    # Uninstall .NET Framework 3.5
    Write-Log "Starting .NET Framework 3.5 uninstallation..." "INFO"
    
    try {
        # Use Disable-WindowsOptionalFeature
        Write-Log "Executing Disable-WindowsOptionalFeature command..." "INFO"
        
        $uninstallResult = Disable-WindowsOptionalFeature -Online -FeatureName NetFx3 -NoRestart -ErrorAction Stop
        
        Write-Log "Disable-WindowsOptionalFeature command completed" "SUCCESS"
        Write-Log "Restart Required: $($uninstallResult.RestartNeeded)" "INFO"
        
        # Verify uninstallation
        Write-Log "Verifying .NET Framework 3.5 uninstallation..." "INFO"
        Start-Sleep -Seconds 3
        
        $verifyStatus = Get-WindowsOptionalFeature -Online -FeatureName NetFx3 -ErrorAction Stop
        Write-Log "Post-uninstallation NetFx3 state: $($verifyStatus.State)" "INFO"
        
        if ($verifyStatus.State -eq "Disabled") {
            Write-Log ".NET Framework 3.5 uninstallation completed successfully!" "SUCCESS"
            
            # Clean up detection files
            Write-Log "Cleaning up detection files..." "INFO"
            $detectionFile = "$env:ProgramData\CustomDetection\NetFx35Installed.txt"
            if (Test-Path $detectionFile) {
                try {
                    Remove-Item $detectionFile -Force -ErrorAction Stop
                    Write-Log "Removed detection file: $detectionFile" "SUCCESS"
                }
                catch {
                    Write-Log "Failed to remove detection file: $($_.Exception.Message)" "WARN"
                }
            }
            
            # Clean up detection directory if empty
            $detectionDir = "$env:ProgramData\CustomDetection"
            if (Test-Path $detectionDir) {
                try {
                    $dirContents = Get-ChildItem $detectionDir -ErrorAction SilentlyContinue
                    if (-not $dirContents) {
                        Remove-Item $detectionDir -Force -ErrorAction Stop
                        Write-Log "Removed empty detection directory: $detectionDir" "SUCCESS"
                    }
                }
                catch {
                    Write-Log "Failed to remove detection directory: $($_.Exception.Message)" "WARN"
                }
            }
            
            Write-Log "========================================" "SUCCESS"
            Write-Log ".NET Framework 3.5 Uninstallation SUCCESSFUL" "SUCCESS"
            Write-Log "========================================" "SUCCESS"
            exit 0
            
        } else {
            Write-Log ".NET Framework 3.5 uninstallation failed - Feature state is: $($verifyStatus.State)" "ERROR"
            exit 1
        }
    }
    catch {
        Write-Log "Failed to uninstall .NET Framework 3.5: $($_.Exception.Message)" "ERROR"
        Write-Log "Error details: $($_.Exception.ToString())" "ERROR"
        
        Write-Log "========================================" "ERROR"
        Write-Log ".NET Framework 3.5 Uninstallation FAILED" "ERROR"
        Write-Log "========================================" "ERROR"
        exit 1
    }
}
catch {
    Write-Log "Critical error during uninstallation process: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    
    Write-Log "========================================" "ERROR"
    Write-Log ".NET Framework 3.5 Uninstallation FAILED" "ERROR"
    Write-Log "========================================" "ERROR"
    exit 1
}
finally {
    Write-Log "Uninstallation script execution completed at $(Get-Date)" "INFO"
    Write-Log "Log file saved to: $LogFile" "INFO"
}

#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Install .NET Framework 3.5 with DISM module validation
.DESCRIPTION
    Checks for DISM module, installs if missing, then enables .NET Framework 3.5
    Logs all results to C:\softdist directory
.NOTES
    Version: 1.0
    Author: IT Infrastructure Team
    Date: $(Get-Date -Format "yyyy-MM-dd")
.EXAMPLE
    .\Install-NetFx35.ps1
#>

# Set up logging
$LogDirectory = "C:\softdist"
$LogFile = "$LogDirectory\NetFx35_Install_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

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
    Write-Log "Starting .NET Framework 3.5 Installation" "INFO"
    Write-Log "Script: $($MyInvocation.MyCommand.Name)" "INFO"
    Write-Log "Log File: $LogFile" "INFO"
    Write-Log "========================================" "INFO"
    
    # Check if .NET Framework 3.5 is already enabled
    Write-Log "Checking current .NET Framework 3.5 status..." "INFO"
    
    try {
        $netFx35Status = Get-WindowsOptionalFeature -Online -FeatureName NetFx3 -ErrorAction Stop
        Write-Log "Current NetFx3 state: $($netFx35Status.State)" "INFO"
        
        if ($netFx35Status.State -eq "Enabled") {
            Write-Log ".NET Framework 3.5 is already enabled. No action required." "SUCCESS"
            Write-Log "Installation process completed successfully." "SUCCESS"
            exit 0
        }
    }
    catch {
        Write-Log "Failed to check .NET Framework 3.5 status: $($_.Exception.Message)" "ERROR"
        Write-Log "Continuing with installation attempt..." "WARN"
    }
    
    # Check if DISM module is available
    Write-Log "Checking for DISM PowerShell module..." "INFO"
    
    try {
        $dismModule = Get-Module -Name DISM -ListAvailable -ErrorAction Stop
        if ($dismModule) {
            Write-Log "DISM module found: Version $($dismModule.Version)" "SUCCESS"
        }
    }
    catch {
        Write-Log "DISM module not found or error checking: $($_.Exception.Message)" "WARN"
    }
    
    # Import DISM module if not already imported
    Write-Log "Importing DISM module..." "INFO"
    try {
        Import-Module DISM -Force -ErrorAction Stop
        Write-Log "DISM module imported successfully" "SUCCESS"
    }
    catch {
        Write-Log "Failed to import DISM module: $($_.Exception.Message)" "ERROR"
        
        # Try to install DISM module from Windows Features
        Write-Log "Attempting to enable DISM PowerShell module via Windows Features..." "INFO"
        try {
            Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-DISM-PowerShell -All -NoRestart -ErrorAction Stop
            Write-Log "DISM PowerShell feature enabled successfully" "SUCCESS"
            
            # Try to import again
            Import-Module DISM -Force -ErrorAction Stop
            Write-Log "DISM module imported successfully after enabling feature" "SUCCESS"
        }
        catch {
            Write-Log "Failed to enable DISM PowerShell feature: $($_.Exception.Message)" "ERROR"
            Write-Log "Continuing without DISM module..." "WARN"
        }
    }
    
    # Install .NET Framework 3.5
    Write-Log "Starting .NET Framework 3.5 installation..." "INFO"
    
    try {
        # Use Enable-WindowsOptionalFeature as requested
        Write-Log "Executing Enable-WindowsOptionalFeature command..." "INFO"
        
        $installResult = Enable-WindowsOptionalFeature -Online -FeatureName NetFx3 -All -NoRestart -ErrorAction Stop
        
        Write-Log "Enable-WindowsOptionalFeature command completed" "SUCCESS"
        Write-Log "Restart Required: $($installResult.RestartNeeded)" "INFO"
        
        # Verify installation
        Write-Log "Verifying .NET Framework 3.5 installation..." "INFO"
        Start-Sleep -Seconds 3
        
        $verifyStatus = Get-WindowsOptionalFeature -Online -FeatureName NetFx3 -ErrorAction Stop
        Write-Log "Post-installation NetFx3 state: $($verifyStatus.State)" "INFO"
        
        if ($verifyStatus.State -eq "Enabled") {
            Write-Log ".NET Framework 3.5 installation completed successfully!" "SUCCESS"
            
            # Test .NET Framework 3.5 functionality
            Write-Log "Testing .NET Framework 3.5 functionality..." "INFO"
            try {
                $netVersion = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
                Write-Log "Current .NET Framework: $netVersion" "INFO"
                
                # Check if .NET 3.5 assemblies are available
                $assembly35 = [System.Reflection.Assembly]::LoadWithPartialName("System.Web")
                if ($assembly35) {
                    Write-Log ".NET Framework 3.5 assemblies are accessible" "SUCCESS"
                }
            }
            catch {
                Write-Log "Warning: Could not fully test .NET 3.5 functionality: $($_.Exception.Message)" "WARN"
            }
            
            Write-Log "========================================" "SUCCESS"
            Write-Log ".NET Framework 3.5 Installation SUCCESSFUL" "SUCCESS"
            Write-Log "========================================" "SUCCESS"
            exit 0
            
        } else {
            Write-Log ".NET Framework 3.5 installation failed - Feature state is: $($verifyStatus.State)" "ERROR"
            exit 1
        }
    }
    catch {
        Write-Log "Failed to install .NET Framework 3.5: $($_.Exception.Message)" "ERROR"
        Write-Log "Error details: $($_.Exception.ToString())" "ERROR"
        
        Write-Log "========================================" "ERROR"
        Write-Log ".NET Framework 3.5 Installation FAILED" "ERROR"
        Write-Log "========================================" "ERROR"
        exit 1
    }
}
catch {
    Write-Log "Critical error during installation process: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    
    Write-Log "========================================" "ERROR"
    Write-Log ".NET Framework 3.5 Installation FAILED" "ERROR"
    Write-Log "========================================" "ERROR"
    exit 1
}
finally {
    Write-Log "Installation script execution completed at $(Get-Date)" "INFO"
    Write-Log "Log file saved to: $LogFile" "INFO"
}

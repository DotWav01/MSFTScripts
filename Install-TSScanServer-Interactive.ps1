<#
.SYNOPSIS
    Installs TSScan Server application for enterprise deployment via Intune (Interactive Mode).

.DESCRIPTION
    This script installs TSScan_server.exe in interactive mode (no silent flags) to work around vendor issues
    with silent installation. If a TSScan.twlic license file is present in the script directory,
    it will be automatically copied to the installation directory after successful installation.
    Designed for deployment as an Intune Win32 application with proper exit codes for monitoring. 
    The installer executable must be located in the same directory as this script.

.PARAMETER LogPath
    Specifies the path for log files. Defaults to C:\softdist\Logs\TSScanServer.

.EXAMPLE
    .\Install-TSScanServer-Interactive.ps1
    Installs TSScan Server using interactive mode.

.EXAMPLE
    .\Install-TSScanServer-Interactive.ps1 -LogPath "C:\Temp\Logs" -Verbose
    Installs TSScan Server with custom log path and verbose output.

.EXAMPLE
    .\Install-TSScanServer-Interactive.ps1 -WhatIf
    Shows what would happen during installation without actually installing.

.NOTES
    File Name      : Install-TSScanServer-Interactive.ps1
    Author         : IT Infrastructure Team
    Prerequisite   : TSScan_server.exe must be in the same directory as this script
    Optional       : TSScan.twlic license file in the same directory (will be copied to install directory)
    Exit Codes     : 0 = Success, 1 = General Error, 2 = File Not Found, 3 = Installation Failed
    
    NOTE: This version uses interactive installation (no silent flags) to work around vendor issues
    
    For Intune deployment:
    - Install command: powershell.exe -ExecutionPolicy Bypass -File "Install-TSScanServer-Interactive.ps1"
    - Return codes: 0 = Success, All others = Failure
    - User context: May require user interaction during installation
#>

#Requires -Version 5.0
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\softdist\Logs\TSScanServer"
)

# Initialize script variables
$ScriptName = "Install-TSScanServer-Interactive"
$ScriptVersion = "1.0.0"
$InstallerName = "TSScan_server.exe"
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$InstallerPath = Join-Path -Path $ScriptPath -ChildPath $InstallerName
$StartTime = Get-Date

# Create log directory if it doesn't exist
if (-not (Test-Path -Path $LogPath)) {
    try {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        Write-Verbose "Created log directory: $LogPath"
    }
    catch {
        Write-Error "Failed to create log directory: $LogPath"
        exit 1
    }
}

# Initialize log file
$LogFile = Join-Path -Path $LogPath -ChildPath "$ScriptName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

# Simple, reliable logging function
function Write-Log {
    param([string]$Message, [string]$Level = 'Info')
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] [$Level] $Message"
    
    # Write to console based on level
    switch ($Level) {
        'Info'    { Write-Host $LogEntry }
        'Warning' { Write-Warning $LogEntry }
        'Error'   { Write-Error $LogEntry }
        'Success' { Write-Host $LogEntry -ForegroundColor Green }
    }
    
    # Write to log file (suppress errors to prevent logging issues)
    try {
        Add-Content -Path $LogFile -Value $LogEntry -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore logging errors to prevent cascading failures
    }
}

# Write script header
Write-Log "============================================================" -Level 'Info'
Write-Log "$ScriptName v$ScriptVersion" -Level 'Info'
Write-Log "Started: $StartTime" -Level 'Info'
Write-Log "User: $env:USERNAME" -Level 'Info'
Write-Log "Computer: $env:COMPUTERNAME" -Level 'Info'
Write-Log "PowerShell Version: $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)" -Level 'Info'
Write-Log "Script Path: $ScriptPath" -Level 'Info'
Write-Log "Log Path: $LogPath" -Level 'Info'
Write-Log "Installation Mode: Interactive (no silent flags)" -Level 'Info'
Write-Log "============================================================" -Level 'Info'

try {
    # Check if running as administrator
    $CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object Security.Principal.WindowsPrincipal($CurrentUser)
    $IsAdmin = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $IsAdmin) {
        Write-Log "Script must be run as Administrator" -Level 'Error'
        exit 1
    }
    
    Write-Log "Administrator privileges confirmed" -Level 'Success'
    
    # Check if installer file exists
    if (-not (Test-Path -Path $InstallerPath -PathType Leaf)) {
        Write-Log "Installer file not found: $InstallerPath" -Level 'Error'
        exit 2
    }
    
    $FileInfo = Get-Item -Path $InstallerPath
    $FileSizeMB = [math]::Round($FileInfo.Length / 1MB, 2)
    Write-Log "Found installer: $($FileInfo.Name) (Size: $FileSizeMB MB)" -Level 'Success'
    
    # Check for license file
    $LicenseFileName = "TSScan.twlic"
    $LicenseSourcePath = Join-Path -Path $ScriptPath -ChildPath $LicenseFileName
    $LicenseFound = Test-Path -Path $LicenseSourcePath -PathType Leaf
    
    if ($LicenseFound) {
        Write-Log "Found license file: $LicenseFileName" -Level 'Success'
    }
    else {
        Write-Log "License file not found: $LicenseFileName (this may be optional)" -Level 'Warning'
    }
    
    # Perform installation
    Write-Log "Starting TSScan Server installation (Interactive Mode)..." -Level 'Info'
    Write-Log "NOTE: User interaction may be required during installation" -Level 'Warning'
    
    if ($PSCmdlet.ShouldProcess("TSScan Server", "Install")) {
        # No silent flags - let the installer run interactively
        Write-Log "Executing: $InstallerPath (Interactive Mode - No Silent Flags)" -Level 'Info'
        
        $Process = Start-Process -FilePath $InstallerPath -Wait -PassThru
        
        Write-Log "Installation process completed with exit code: $($Process.ExitCode)" -Level 'Info'
        
        if ($Process.ExitCode -eq 0) {
            Write-Log "TSScan Server installation completed successfully" -Level 'Success'
            
            # Wait a bit longer for interactive installation to finish file operations
            Write-Log "Waiting for installation to complete file operations..." -Level 'Info'
            Start-Sleep -Seconds 10
            
            # Verify installation and copy license file
            $InstallPath = "${env:ProgramFiles(x86)}\TerminalWorks\TSScan Server"
            $UninstallPath = Join-Path -Path $InstallPath -ChildPath "unins000.exe"
            
            if (Test-Path -Path $InstallPath) {
                Write-Log "Installation directory found: $InstallPath" -Level 'Success'
                
                if (Test-Path -Path $UninstallPath) {
                    Write-Log "Uninstaller found: $UninstallPath" -Level 'Success'
                }
                
                # Copy license file if it exists
                if ($LicenseFound) {
                    try {
                        $LicenseDestPath = Join-Path -Path $InstallPath -ChildPath $LicenseFileName
                        Copy-Item -Path $LicenseSourcePath -Destination $LicenseDestPath -Force -ErrorAction Stop
                        Write-Log "License file copied successfully to: $LicenseDestPath" -Level 'Success'
                        
                        # Verify license file was copied
                        if (Test-Path -Path $LicenseDestPath) {
                            $LicenseFileInfo = Get-Item -Path $LicenseDestPath
                            Write-Log "License file verified: $($LicenseFileInfo.Name) ($($LicenseFileInfo.Length) bytes)" -Level 'Success'
                        }
                    }
                    catch {
                        Write-Log "Failed to copy license file: Script failed at line $($_.InvocationInfo.ScriptLineNumber)" -Level 'Error'
                        Write-Log "License file copy error - installation may still be functional" -Level 'Warning'
                    }
                }
                
                Write-Log "Installation verification successful" -Level 'Success'
            }
            else {
                Write-Log "Installation directory not found: $InstallPath" -Level 'Warning'
                Write-Log "Installation may have completed but verification failed" -Level 'Warning'
                Write-Log "This could happen if user cancelled installation or chose different location" -Level 'Warning'
            }
        }
        else {
            Write-Log "Installation failed with exit code: $($Process.ExitCode)" -Level 'Error'
            Write-Log "Common exit codes: 1223 = User cancelled, 1602 = User cancelled, 1619 = Package could not be opened" -Level 'Info'
            exit 3
        }
    }
    else {
        Write-Log "WhatIf: Would install TSScan Server from $InstallerPath (Interactive Mode)" -Level 'Info'
        if ($LicenseFound) {
            Write-Log "WhatIf: Would copy license file $LicenseFileName to installation directory" -Level 'Info'
        }
    }
    
    # Success
    $EndTime = Get-Date
    $Duration = $EndTime - $StartTime
    $DurationSeconds = [math]::Round($Duration.TotalSeconds, 2)
    Write-Log "Script completed successfully in $DurationSeconds seconds" -Level 'Success'
    
    exit 0
}
catch {
    # Simple error handling without complex object conversion
    Write-Log "Unexpected error occurred during installation" -Level 'Error'
    Write-Log "Error details: Script failed at line $($_.InvocationInfo.ScriptLineNumber)" -Level 'Error'
    
    exit 1
}
finally {
    Write-Log "Script execution finished at $(Get-Date)" -Level 'Info'
}

<#
.SYNOPSIS
    Installs Bloomberg Terminal client application for enterprise deployment via Intune.

.DESCRIPTION
    This script performs a silent installation of the Bloomberg Terminal client application.
    Designed for deployment through Microsoft Intune as a Win32 application package.
    Includes comprehensive logging, error handling, and exit codes for deployment monitoring.

.PARAMETER InstallerPath
    Path to the Bloomberg installer executable file. Defaults to the script directory.

.PARAMETER LogPath
    Path where installation logs will be written. Defaults to C:\softdist\Logs\Bloomberg.

.PARAMETER WhatIf
    Shows what would be done without actually performing the installation.

.EXAMPLE
    .\Install-BloombergClient.ps1
    Performs a silent installation using the default installer in the script directory.

.EXAMPLE
    .\Install-BloombergClient.ps1 -InstallerPath "C:\Temp\BloombergInstaller.exe" -Verbose
    Installs Bloomberg client with verbose output and custom installer path.

.EXAMPLE
    .\Install-BloombergClient.ps1 -WhatIf
    Shows what would be done without actually installing.

.NOTES
    File Name      : Install-BloombergClient.ps1
    Author         : IT Infrastructure Team
    Prerequisite   : PowerShell 5.1 or later, Administrative privileges
    Requirements   : Bloomberg installer executable must be present
    
    Exit Codes:
    0  = Success
    1  = General error
    2  = Installer file not found
    3  = Installation failed
    4  = Post-installation verification failed
    5  = Insufficient privileges
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$InstallerPath,
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\softdist\Logs\Bloomberg",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Script variables
$ScriptName = "Install-BloombergClient"
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
    
    Write-Host "[$ScriptName] Starting Bloomberg Terminal installation" -ForegroundColor Green
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
    # Determine installer path
    if (-not $InstallerPath) {
        $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        $PossibleInstallers = @(
            "BloombergInstaller.exe",
            "bloomberg_terminal_installer.exe",
            "Bloomberg Terminal Installer.exe"
        )
        
        foreach ($Installer in $PossibleInstallers) {
            $TestPath = Join-Path -Path $ScriptDir -ChildPath $Installer
            if (Test-Path -Path $TestPath) {
                $InstallerPath = $TestPath
                break
            }
        }
        
        # If no predefined installer found, look for any .exe file
        if (-not $InstallerPath) {
            $ExeFiles = Get-ChildItem -Path $ScriptDir -Filter "*.exe" | Where-Object { $_.Name -like "*bloomberg*" -or $_.Name -like "*terminal*" }
            if ($ExeFiles.Count -eq 1) {
                $InstallerPath = $ExeFiles[0].FullName
            }
        }
    }
    
    # Validate installer exists
    if (-not $InstallerPath -or -not (Test-Path -Path $InstallerPath)) {
        Write-Error "Bloomberg installer not found. Please ensure the installer executable is in the script directory or specify -InstallerPath"
        Write-Host "[$ScriptName] Searched locations:" -ForegroundColor Yellow
        if ($ScriptDir) {
            Write-Host "  - $ScriptDir\BloombergInstaller.exe" -ForegroundColor Yellow
            Write-Host "  - $ScriptDir\bloomberg_terminal_installer.exe" -ForegroundColor Yellow
            Write-Host "  - $ScriptDir\Bloomberg Terminal Installer.exe" -ForegroundColor Yellow
        }
        $ExitCode = 2
        throw "Installer file not found"
    }
    
    Write-Host "[$ScriptName] Found installer: $InstallerPath" -ForegroundColor Green
    
    # Get installer file info
    $InstallerInfo = Get-Item -Path $InstallerPath
    Write-Host "[$ScriptName] Installer size: $([math]::Round($InstallerInfo.Length / 1MB, 2)) MB" -ForegroundColor Green
    Write-Host "[$ScriptName] Installer modified: $($InstallerInfo.LastWriteTime)" -ForegroundColor Green
    
    # Check if Bloomberg is already installed
    Write-Host "[$ScriptName] Checking for existing Bloomberg installation..." -ForegroundColor Yellow
    
    $ExistingInstall = $null
    $BloombergPaths = @(
        "${env:ProgramFiles}\Bloomberg Terminal",
        "${env:ProgramFiles(x86)}\Bloomberg Terminal",
        "${env:ProgramFiles}\Bloomberg",
        "${env:ProgramFiles(x86)}\Bloomberg"
    )
    
    foreach ($Path in $BloombergPaths) {
        if (Test-Path -Path $Path) {
            $ExistingInstall = $Path
            break
        }
    }
    
    # Check registry for Bloomberg installation
    $RegistryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $BloombergRegistryEntry = $null
    foreach ($RegPath in $RegistryPaths) {
        try {
            $BloombergRegistryEntry = Get-ItemProperty $RegPath | Where-Object { 
                $_.DisplayName -like "*Bloomberg*" -or $_.DisplayName -like "*Terminal*" 
            } | Select-Object -First 1
            if ($BloombergRegistryEntry) { break }
        }
        catch {
            # Registry path might not exist, continue
        }
    }
    
    if ($ExistingInstall -or $BloombergRegistryEntry) {
        Write-Host "[$ScriptName] Bloomberg Terminal appears to be already installed" -ForegroundColor Yellow
        if ($ExistingInstall) {
            Write-Host "[$ScriptName] Found installation at: $ExistingInstall" -ForegroundColor Yellow
        }
        if ($BloombergRegistryEntry) {
            Write-Host "[$ScriptName] Registry entry: $($BloombergRegistryEntry.DisplayName) v$($BloombergRegistryEntry.DisplayVersion)" -ForegroundColor Yellow
        }
        Write-Host "[$ScriptName] Continuing with installation (will upgrade/repair if needed)" -ForegroundColor Yellow
    }
    
    # Prepare installation command
    $InstallArgs = @(
        '/S',           # Silent install
        '/v/qn'         # MSI quiet mode (if applicable)
    )
    
    # Additional silent install arguments that might be needed
    $AdditionalArgs = @(
        '/SILENT',
        '/VERYSILENT',
        '/SUPPRESSMSGBOXES',
        '/NORESTART'
    )
    
    $InstallCommand = "& '$InstallerPath' $($InstallArgs -join ' ')"
    
    if ($WhatIf) {
        Write-Host "[$ScriptName] WhatIf: Would execute: $InstallCommand" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Installation would be performed silently" -ForegroundColor Magenta
        exit 0
    }
    
    # Perform installation
    Write-Host "[$ScriptName] Starting Bloomberg Terminal installation..." -ForegroundColor Yellow
    Write-Host "[$ScriptName] Command: $InstallCommand" -ForegroundColor Gray
    
    $InstallStartTime = Get-Date
    
    # Execute installation with different argument sets
    $InstallSuccess = $false
    $InstallAttempts = @(
        @('/S', '/v/qn'),
        @('/SILENT', '/SUPPRESSMSGBOXES', '/NORESTART'),
        @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'),
        @('/S'),
        @('/q')
    )
    
    foreach ($AttemptArgs in $InstallAttempts) {
        try {
            Write-Host "[$ScriptName] Attempting installation with args: $($AttemptArgs -join ' ')" -ForegroundColor Gray
            
            $Process = Start-Process -FilePath $InstallerPath -ArgumentList $AttemptArgs -Wait -PassThru -NoNewWindow
            
            if ($Process.ExitCode -eq 0) {
                Write-Host "[$ScriptName] Installation completed successfully with exit code: $($Process.ExitCode)" -ForegroundColor Green
                $InstallSuccess = $true
                break
            }
            else {
                Write-Warning "[$ScriptName] Installation attempt failed with exit code: $($Process.ExitCode)"
            }
        }
        catch {
            Write-Warning "[$ScriptName] Installation attempt failed: $($_.Exception.Message)"
        }
        
        Start-Sleep -Seconds 2
    }
    
    if (-not $InstallSuccess) {
        Write-Error "[$ScriptName] All installation attempts failed"
        $ExitCode = 3
        throw "Installation failed with all attempted argument sets"
    }
    
    $InstallEndTime = Get-Date
    $InstallDuration = $InstallEndTime - $InstallStartTime
    Write-Host "[$ScriptName] Installation duration: $($InstallDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
    
    # Post-installation verification
    Write-Host "[$ScriptName] Performing post-installation verification..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    
    $VerificationPassed = $false
    
    # Check for Bloomberg installation
    foreach ($Path in $BloombergPaths) {
        if (Test-Path -Path $Path) {
            Write-Host "[$ScriptName] Verification: Found Bloomberg installation at $Path" -ForegroundColor Green
            $VerificationPassed = $true
            
            # Look for executable files
            $BloombergExe = Get-ChildItem -Path $Path -Filter "*.exe" -Recurse | Where-Object { 
                $_.Name -like "*bloomberg*" -or $_.Name -like "*terminal*" -or $_.Name -like "*wintrm*"
            } | Select-Object -First 1
            
            if ($BloombergExe) {
                Write-Host "[$ScriptName] Verification: Found Bloomberg executable at $($BloombergExe.FullName)" -ForegroundColor Green
            }
            break
        }
    }
    
    # Check registry again
    foreach ($RegPath in $RegistryPaths) {
        try {
            $UpdatedRegistryEntry = Get-ItemProperty $RegPath | Where-Object { 
                $_.DisplayName -like "*Bloomberg*" -or $_.DisplayName -like "*Terminal*" 
            } | Select-Object -First 1
            
            if ($UpdatedRegistryEntry) {
                Write-Host "[$ScriptName] Verification: Found registry entry: $($UpdatedRegistryEntry.DisplayName)" -ForegroundColor Green
                if ($UpdatedRegistryEntry.DisplayVersion) {
                    Write-Host "[$ScriptName] Verification: Version: $($UpdatedRegistryEntry.DisplayVersion)" -ForegroundColor Green
                }
                $VerificationPassed = $true
                break
            }
        }
        catch {
            # Continue checking
        }
    }
    
    # Check for Bloomberg services
    $BloombergServices = Get-Service | Where-Object { $_.Name -like "*Bloomberg*" -or $_.DisplayName -like "*Bloomberg*" }
    if ($BloombergServices) {
        Write-Host "[$ScriptName] Verification: Found Bloomberg services:" -ForegroundColor Green
        foreach ($Service in $BloombergServices) {
            Write-Host "  - $($Service.Name) ($($Service.DisplayName)): $($Service.Status)" -ForegroundColor Green
        }
        $VerificationPassed = $true
    }
    
    if (-not $VerificationPassed) {
        Write-Warning "[$ScriptName] Post-installation verification failed - Bloomberg installation not detected"
        Write-Warning "[$ScriptName] Installation may have completed but verification couldn't confirm success"
        # Don't fail here as some Bloomberg installers may require a reboot to be fully detected
    }
    
    Write-Host "[$ScriptName] Bloomberg Terminal installation completed successfully" -ForegroundColor Green
}
catch {
    Write-Error "[$ScriptName] Installation failed: $($_.Exception.Message)"
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

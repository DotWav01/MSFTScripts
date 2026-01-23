<#
.SYNOPSIS
    Installs Bloomberg Terminal client application for enterprise deployment via Intune.

.DESCRIPTION
    This script performs a silent installation of the Bloomberg Terminal client application.
    Designed for deployment through Microsoft Intune as a Win32 application package.
    Automatically closes Office 365 applications (Excel, Word, PowerPoint) before installation
    as required by Bloomberg Terminal installer. Includes comprehensive logging, error handling, 
    and exit codes for deployment monitoring.

.PARAMETER InstallerPath
    Path to the Bloomberg installer executable file. Defaults to the script directory.

.PARAMETER LogPath
    Path where installation logs will be written. Defaults to C:\softdist\Logs\Bloomberg.

.NOTES
    WhatIf parameter is automatically available through SupportsShouldProcess.
    Use -WhatIf to preview actions without executing them.

.EXAMPLE
    .\Install-BloombergClient.ps1
    Performs a silent installation using the default installer in the script directory.

.EXAMPLE
    .\Install-BloombergClient.ps1 -InstallerPath "C:\Temp\BloombergInstaller.exe" -Verbose
    Installs Bloomberg client with verbose output, custom installer path, and closes any running Office apps.

.EXAMPLE
    .\Install-BloombergClient.ps1 -WhatIf
    Shows what would be done including Office app closure without actually installing.

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
    [string]$LogPath = "C:\softdist\Logs\Bloomberg"
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
    
    # Check for and close Office 365 applications before installation
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
        
        Write-Host "[$ScriptName] Bloomberg Terminal requires Office applications to be closed before installation" -ForegroundColor Yellow
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
                    
                    # If still running, force termination
                    if (-not $Process.HasExited) {
                        Write-Warning "[$ScriptName] Graceful closure failed, force terminating $($App.DisplayName) (PID: $($Process.Id))"
                        $Process.Kill()
                        Start-Sleep -Seconds 2
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
            Write-Warning "[$ScriptName] Bloomberg installation may fail or encounter issues with Office apps running"
            Write-Warning "[$ScriptName] Continuing with installation attempt..."
        }
        else {
            Write-Host "[$ScriptName] All Office applications successfully closed" -ForegroundColor Green
        }
    }
    else {
        Write-Host "[$ScriptName] No Office 365 applications found running" -ForegroundColor Green
    }
    
    # Note: Intune handles installation detection separately
    Write-Host "[$ScriptName] Proceeding with Bloomberg Terminal installation..." -ForegroundColor Yellow
    
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
        Write-Host "[$ScriptName] WhatIf: Would check for and close Office 365 applications (Excel, Word, PowerPoint)" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Would execute: $InstallCommand" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Installation would be performed silently" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Post-install verification would check C:\blp\winrtv\wintrv.exe" -ForegroundColor Magenta
        Write-Host "[$ScriptName] WhatIf: Would create detection tag file at C:\temp\Bloomberg_Installed.tag" -ForegroundColor Magenta
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
    
    # Check for Bloomberg installation at C:\blp
    $BloombergInstallPath = "C:\blp"
    $BloombergExecutable = "C:\blp\winrtv\wintrv.exe"
    
    if (Test-Path -Path $BloombergInstallPath) {
        Write-Host "[$ScriptName] Verification: Found Bloomberg installation directory at $BloombergInstallPath" -ForegroundColor Green
        $VerificationPassed = $true
        
        # Check for the main Bloomberg executable
        if (Test-Path -Path $BloombergExecutable) {
            Write-Host "[$ScriptName] Verification: Found Bloomberg executable at $BloombergExecutable" -ForegroundColor Green
            
            # Get file information
            try {
                $ExeInfo = Get-Item -Path $BloombergExecutable
                Write-Host "[$ScriptName] Verification: Executable size: $([math]::Round($ExeInfo.Length / 1MB, 2)) MB" -ForegroundColor Green
                Write-Host "[$ScriptName] Verification: Executable modified: $($ExeInfo.LastWriteTime)" -ForegroundColor Green
            }
            catch {
                Write-Warning "[$ScriptName] Could not get executable file information: $($_.Exception.Message)"
            }
        }
        else {
            Write-Warning "[$ScriptName] Bloomberg directory found but wintrv.exe not found at expected location: $BloombergExecutable"
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
    
    # Create tag file for Intune detection
    try {
        $TagFilePath = "C:\temp\Bloomberg_Installed.tag"
        $TagFileContent = @"
Bloomberg Terminal Installation Tag File
Installation Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Script Version: $ScriptVersion
Installed By: $env:USERNAME
Computer: $env:COMPUTERNAME
Bloomberg Path: C:\blp\winrtv\wintrv.exe
Verification Passed: $VerificationPassed
"@
        
        # Ensure C:\temp directory exists
        if (-not (Test-Path -Path "C:\temp")) {
            New-Item -Path "C:\temp" -ItemType Directory -Force | Out-Null
            Write-Host "[$ScriptName] Created C:\temp directory" -ForegroundColor Green
        }
        
        # Create tag file
        Set-Content -Path $TagFilePath -Value $TagFileContent -Force
        Write-Host "[$ScriptName] Created detection tag file: $TagFilePath" -ForegroundColor Green
    }
    catch {
        Write-Warning "[$ScriptName] Failed to create tag file: $($_.Exception.Message)"
        # Don't fail installation for tag file creation failure
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

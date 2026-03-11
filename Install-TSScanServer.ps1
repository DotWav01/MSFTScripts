#Requires -Version 5.0
<#
.SYNOPSIS
    Installs TSScan Server with interactive GUI support for Intune SYSTEM context deployment.

.DESCRIPTION
    Installs TSScan_server.exe in interactive mode for use with ServiceUI.exe.
    This script is called by Launch-TSScanInstall.bat, which is itself launched by
    ServiceUI.exe. This two-stage approach ensures that both PowerShell AND the
    TSScan GUI wizard are projected into the active user session.

    After successful installation, the TSScan license file (tsscan-site.twlic) is
    automatically copied to the appropriate location.

    Session flow:
      Intune (Session 0 SYSTEM)
        -> ServiceUI.exe -process:explorer.exe
          -> Launch-TSScanInstall.bat  (now in user Session 1+)
            -> powershell.exe          (inherits user session)
              -> TSScan_server.exe     (GUI visible to user)

.PARAMETER LogPath
    Directory for log files. Defaults to C:\softdist\Logs\TSScanServer.

.PARAMETER LicenseFileName
    Name of the license file to deploy. Defaults to tsscan-site.twlic.

.EXAMPLE
    .\Install-TSScanServer.ps1
    Standard installation called via Launch-TSScanInstall.bat through ServiceUI.

.EXAMPLE
    .\Install-TSScanServer.ps1 -LogPath "C:\Temp\Logs" -Verbose
    Installation with custom log path and verbose output.

.EXAMPLE
    .\Install-TSScanServer.ps1 -WhatIf
    Test mode - shows what would happen without making changes.

.NOTES
    Version:        2.0
    Author:         IT Infrastructure
    Purpose:        Intune Win32 App - Available (SYSTEM install, interactive via ServiceUI)
    Package files:  TSScan_server.exe, TSScan-site.twlic, ServiceUIx64.exe,
                    Launch-TSScanInstall.bat, Install-TSScanServer.ps1

    Intune Install Command:
        ServiceUIx64.exe -process:explorer.exe Launch-TSScanInstall.bat

    Intune Uninstall Command:
        powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Uninstall-TSScanServer.ps1"

    Install Behavior: System
    Deployment type:  Available (Company Portal)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$LogPath = 'C:\softdist\Logs\TSScanServer',

    [Parameter()]
    [string]$LicenseFileName = 'TSScan-site.twlic'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Logging -----------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry     = "[$timestamp] [$Level] $Message"

    # Console output
    switch ($Level) {
        'WARN'    { Write-Warning $Message }
        'ERROR'   { Write-Error   $Message -ErrorAction Continue }
        'SUCCESS' { Write-Verbose $Message }
        default   { Write-Verbose $Message }
    }

    # File output
    try {
        Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8
    } catch {
        # Suppress logging errors to avoid masking the real error
    }
}
#endregion

#region --- Initialise --------------------------------------------------------
# Resolve the package root directory (same folder as this script)
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Ensure log directory exists
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$script:LogFile = Join-Path $LogPath "Install_TSScanServer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

Write-Log "===== TSScan Server Installation Started ====="
Write-Log "Script root  : $ScriptRoot"
Write-Log "Log file     : $script:LogFile"
Write-Log "Running as   : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

# Detect session context (informational)
try {
    $sessionId = (Get-Process -Id $PID).SessionId
    Write-Log "Session ID   : $sessionId $(if ($sessionId -eq 0) { '(WARNING: Session 0 - GUI may not be visible)' } else { '(User session - GUI will be visible)' })"
} catch {
    Write-Log "Session ID   : Unable to determine" -Level WARN
}
#endregion

#region --- Validate package files -------------------------------------------
$InstallerPath = Join-Path $ScriptRoot 'TSScan_server.exe'
$LicensePath   = Join-Path $ScriptRoot $LicenseFileName

if (-not (Test-Path $InstallerPath)) {
    Write-Log "Installer not found: $InstallerPath" -Level ERROR
    exit 2
}
Write-Log "Installer    : $InstallerPath [FOUND]"

$LicenseFound = Test-Path $LicensePath
if ($LicenseFound) {
    Write-Log "License file : $LicensePath [FOUND]"
} else {
    Write-Log "License file : $LicensePath [NOT FOUND - will skip license deployment]" -Level WARN
}
#endregion

#region --- Run installer (interactive, no silent flags) ----------------------
Write-Log "Launching TSScan installer in interactive mode..."
Write-Log "The user should see the installation wizard on their screen."

if ($PSCmdlet.ShouldProcess($InstallerPath, 'Run interactive installer')) {
    try {
        # Start the installer with NO silent flags so the wizard appears to the user.
        # Because ServiceUI has already projected this PowerShell process into the user
        # session, child processes (including the TSScan wizard) inherit that session
        # and are therefore visible on the user's desktop.
        $proc = Start-Process -FilePath $InstallerPath `
                              -Wait `
                              -PassThru

        $exitCode = $proc.ExitCode
        Write-Log "Installer finished with exit code: $exitCode"

        switch ($exitCode) {
            0       { Write-Log "Installation completed successfully." -Level SUCCESS }
            1602    { Write-Log "User cancelled the installation (exit 1602)." -Level WARN ; exit 1602 }
            1223    { Write-Log "User cancelled the installation (exit 1223)." -Level WARN ; exit 1223 }
            3010    { Write-Log "Installation successful - reboot required (exit 3010)." -Level SUCCESS }
            default {
                Write-Log "Installer returned unexpected exit code: $exitCode" -Level WARN
                # Do not immediately fail - verify installation below
            }
        }
    } catch {
        Write-Log "Failed to launch installer: $_" -Level ERROR
        exit 1
    }
}
#endregion

#region --- Verify installation -----------------------------------------------
Write-Log "Verifying installation..."

$installPaths = @(
    'C:\Program Files (x86)\TerminalWorks\TSScan Server',
    'C:\Program Files\TerminalWorks\TSScan Server'
)

$installedPath = $installPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $installedPath) {
    Write-Log "Installation verification FAILED - install directory not found." -Level ERROR
    Write-Log "Checked paths: $($installPaths -join ', ')"
    exit 1
}

Write-Log "Installation verified at: $installedPath" -Level SUCCESS
#endregion

#region --- Deploy license file -----------------------------------------------
if ($LicenseFound) {
    Write-Log "Deploying license file to: $installedPath"

    if ($PSCmdlet.ShouldProcess($installedPath, 'Copy license file')) {
        try {
            Copy-Item -Path $LicensePath -Destination $installedPath -Force
            Write-Log "License file deployed successfully." -Level SUCCESS
        } catch {
            # Non-fatal: log the error but do not fail the install
            Write-Log "Failed to copy license file: $_" -Level WARN
            Write-Log "Manual license deployment may be required." -Level WARN
        }
    }
} else {
    Write-Log "Skipping license deployment (file not found in package)." -Level WARN
}
#endregion

Write-Log "===== TSScan Server Installation Completed Successfully ====="
exit 0

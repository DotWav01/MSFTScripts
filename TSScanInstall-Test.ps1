#Requires -Version 5.1

<#
.SYNOPSIS
    Installs TSScan Server and deploys the site license file.

.DESCRIPTION
    Launched via ServiceUI_x64.exe to bridge the SYSTEM context to the active
    user session, allowing the TSScan Server interactive installer wizard to be
    displayed to the user. After the wizard completes, the site license file is
    copied to the installation directory.

.NOTES
    Author  : IT Infrastructure
    Version : 1.0.0
    Date    : 2026-03-12

    Deployment:
        Intune Win32 - System context
        Install : ServiceUI_x64.exe -Process:explorer.exe powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Install-TSScanServer.ps1"
        Uninstall: powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Uninstall-TSScanServer.ps1"
        Detection: File exists - C:\Program Files (x86)\TerminalWorks\TSScan Server\unins000.exe

.EXAMPLE
    # Run manually as SYSTEM via PsExec for testing:
    psexec.exe -s powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Install-TSScanServer.ps1"
#>

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
$LogPath = 'C:\softdist\Logs'
$LogFile = Join-Path $LogPath 'TSScanServer_Install.log'

if (-not (Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry
    Write-Output $entry
}

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Definition
$InstallerPath = Join-Path $ScriptDir 'TSScan_server.exe'
$LicenseSrc    = Join-Path $ScriptDir 'TSScan-site.twlic'

Write-Log "=== TSScan Server Install Started ==="
Write-Log "Script directory : $ScriptDir"
Write-Log "Running as       : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------
foreach ($file in @($InstallerPath, $LicenseSrc)) {
    if (-not (Test-Path $file)) {
        Write-Log "Required file not found: $file" 'ERROR'
        exit 1
    }
}
Write-Log "Pre-flight checks passed. Installer and license file verified."

# ---------------------------------------------------------------------------
# Launch interactive installer
# ---------------------------------------------------------------------------
Write-Log "Launching TSScan Server interactive installer wizard..."
try {
    $process = Start-Process -FilePath $InstallerPath `
        -Wait `
        -PassThru

    Write-Log "Installer exited with code: $($process.ExitCode)"
} catch {
    Write-Log "Failed to launch installer: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Verify installation directory exists
# ---------------------------------------------------------------------------
$installDirX86 = "${env:ProgramFiles(x86)}\TerminalWorks\TSScan Server"
$installDirX64 = "$env:ProgramFiles\TerminalWorks\TSScan Server"

if (Test-Path $installDirX86) {
    $installDir = $installDirX86
} elseif (Test-Path $installDirX64) {
    $installDir = $installDirX64
} else {
    Write-Log "Installation directory not found after installer completed. Installer may have been cancelled or failed." 'ERROR'
    exit 1
}

Write-Log "Installation directory found: $installDir"

# ---------------------------------------------------------------------------
# Copy license file
# ---------------------------------------------------------------------------
$licenseDst = Join-Path $installDir 'TSScan-site.twlic'
try {
    Copy-Item -Path $LicenseSrc -Destination $licenseDst -Force
    Write-Log "License file copied to: $licenseDst"
} catch {
    Write-Log "Failed to copy license file: $_" 'ERROR'
    exit 1
}

Write-Log "=== TSScan Server Install Completed Successfully ==="
exit 0

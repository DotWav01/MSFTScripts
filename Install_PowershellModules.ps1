<#
.SYNOPSIS
    Installs the Microsoft.Graph module and dependencies from a local offline package.

.DESCRIPTION
    Installs Microsoft.Graph and all its dependencies from a folder previously populated
    by Save-GraphModuleOffline.ps1. No internet or PSGallery access is required.

    Supports two installation scopes:
    - AllUsers  : Installs to $env:ProgramFiles\WindowsPowerShell\Modules (requires elevation)
    - CurrentUser: Installs to $HOME\Documents\WindowsPowerShell\Modules (no elevation needed)

.PARAMETER SourcePath
    Path to the folder containing the downloaded module files (output of Save-GraphModuleOffline.ps1).

.PARAMETER Scope
    Installation scope. 'AllUsers' (default) or 'CurrentUser'.
    AllUsers requires the script to be run as Administrator.

.PARAMETER Force
    Overwrite existing module versions if already installed.

.PARAMETER ModuleName
    Target module to install. Defaults to 'Microsoft.Graph'.
    The script installs this module and all sub-modules found in SourcePath.

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline'
    Installs all modules from the staging folder for all users (requires elevation).

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline' -Scope CurrentUser
    Installs for the current user only. No elevation required.

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline' -Force -Verbose
    Reinstalls/upgrades modules, with verbose output.

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - Administrator rights if using -Scope AllUsers
    - SourcePath must be the output folder from Save-GraphModuleOffline.ps1

    Log file is written to C:\softdist\Logs\ (consistent with enterprise logging convention).
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$SourcePath,

    [Parameter()]
    [ValidateSet('AllUsers', 'CurrentUser')]
    [string]$Scope = 'AllUsers',

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [string]$ModuleName = 'Microsoft.Graph'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Logging ──────────────────────────────────────────────────────────────────
$LogDir  = 'C:\softdist\Logs'
$LogFile = Join-Path $LogDir "Install-GraphModule_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Verbose $entry
    if ($Level -eq 'ERROR') { Write-Error $Message }
    elseif ($Level -eq 'WARN') { Write-Warning $Message }
    else { Write-Host $entry }
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

# ── Pre-flight ────────────────────────────────────────────────────────────────
# Ensure log directory exists
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

Write-Host "`n=== Microsoft.Graph Offline Installer ===" -ForegroundColor Cyan
Write-Log "Starting offline installation from: $SourcePath"
Write-Log "Scope: $Scope | Force: $Force"

# Elevation check for AllUsers scope
if ($Scope -eq 'AllUsers') {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "AllUsers scope requires Administrator privileges. Re-run as Administrator or use -Scope CurrentUser." -Level ERROR
        exit 1
    }
}

# Determine install destination
$installBase = if ($Scope -eq 'AllUsers') {
    "$env:ProgramFiles\WindowsPowerShell\Modules"
} else {
    "$HOME\Documents\WindowsPowerShell\Modules"
}
Write-Log "Install destination: $installBase"

# ── Register local repository ─────────────────────────────────────────────────
$repoName = 'GraphOfflineRepo'
$resolvedSource = Resolve-Path $SourcePath

# Remove existing temp repo registration if present
if (Get-PSRepository -Name $repoName -ErrorAction SilentlyContinue) {
    Write-Log "Removing existing repository registration: $repoName"
    Unregister-PSRepository -Name $repoName
}

Write-Log "Registering local repository: $repoName -> $resolvedSource"
try {
    Register-PSRepository -Name $repoName `
        -SourceLocation $resolvedSource `
        -InstallationPolicy Trusted `
        -ErrorAction Stop
    Write-Log "Repository registered successfully."
} catch {
    Write-Log "Failed to register local repository: $_" -Level ERROR
    exit 1
}

# ── Discover modules in source ────────────────────────────────────────────────
$modulefolders = Get-ChildItem -Path $SourcePath -Directory | Sort-Object Name
Write-Log "Found $($modulefolders.Count) module folder(s) in source path."

if ($modulefolders.Count -eq 0) {
    Write-Log "No module folders found in '$SourcePath'. Verify the source path is correct." -Level ERROR
    Unregister-PSRepository -Name $repoName -ErrorAction SilentlyContinue
    exit 1
}

# ── Install modules ───────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$installed = 0
$skipped   = 0
$failed    = 0

# Install dependencies first (everything except the meta-module), then the meta-module last
$dependencyModules = $modulefolders | Where-Object { $_.Name -ne $ModuleName }
$metaModule        = $modulefolders | Where-Object { $_.Name -eq $ModuleName }
$orderedModules    = @($dependencyModules) + @($metaModule) | Where-Object { $_ }

foreach ($folder in $orderedModules) {
    $name = $folder.Name

    # Detect version from subfolder (Save-Module creates Name\Version\ structure)
    $versionFolder = Get-ChildItem -Path $folder.FullName -Directory | Sort-Object Name -Descending | Select-Object -First 1
    $version = if ($versionFolder) { $versionFolder.Name } else { 'unknown' }

    # Check if already installed
    $existingModule = Get-Module -Name $name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    if ($existingModule -and -not $Force) {
        Write-Log "SKIP: $name (already installed: v$($existingModule.Version))" -Level WARN
        $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Status = 'Skipped' })
        $skipped++
        continue
    }

    $installParams = @{
        Name        = $name
        Repository  = $repoName
        Scope       = $Scope
        Force       = $Force.IsPresent
        ErrorAction = 'Stop'
    }

    if ($PSCmdlet.ShouldProcess("$name v$version", "Install module to $Scope")) {
        try {
            Write-Log "Installing: $name v$version"
            Install-Module @installParams
            Write-Log "  OK: $name v$version installed."
            $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Status = 'Installed' })
            $installed++
        } catch {
            Write-Log "  FAILED: $name - $_" -Level WARN
            $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Status = "Failed: $_" })
            $failed++
        }
    }
}

# ── Cleanup ───────────────────────────────────────────────────────────────────
Write-Log "Unregistering temporary repository: $repoName"
Unregister-PSRepository -Name $repoName -ErrorAction SilentlyContinue

# ── Verify installation ───────────────────────────────────────────────────────
Write-Log "Verifying Microsoft.Graph installation..."
try {
    $graphModule = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    if ($graphModule) {
        Write-Log "Verification OK: $ModuleName v$($graphModule.Version) is available."
    } else {
        Write-Log "Verification WARNING: '$ModuleName' not found in available modules after install." -Level WARN
    }
} catch {
    Write-Log "Verification check failed: $_" -Level WARN
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n=== Installation Summary ===" -ForegroundColor Cyan
Write-Host "Installed : $installed"
Write-Host "Skipped   : $skipped (already present; use -Force to overwrite)"
Write-Host "Failed    : $failed"
Write-Host "`nResults:" -ForegroundColor Cyan
$results | Format-Table -AutoSize

Write-Host "`nLog file: $LogFile" -ForegroundColor Gray

if ($failed -gt 0) {
    Write-Host "`nSome modules failed to install. Check the log for details." -ForegroundColor Yellow
    exit 1
}

Write-Host "`nInstallation complete. Test with: Import-Module Microsoft.Graph" -ForegroundColor Green

<#
.SYNOPSIS
    Installs the Microsoft.Graph module and dependencies from a local offline package.

.DESCRIPTION
    Installs Microsoft.Graph and all its dependencies from a folder previously populated
    by Save-GraphModuleOffline.ps1. No internet or PSGallery access is required.

    By default installs to BOTH PowerShell 5.1 and PowerShell 7+ module paths so the
    module is available regardless of which PS version is used.

    Uses direct folder copy into PSModulePath — no Register-PSRepository needed.

.PARAMETER SourcePath
    Path to the folder containing the downloaded module files (output of Save-GraphModuleOffline.ps1).
    Expected structure: SourcePath\ModuleName\Version\<module files>

.PARAMETER Scope
    Installation scope:
    - AllUsers    (default) : Installs to ProgramFiles paths. Requires elevation.
    - CurrentUser           : Installs to user Documents paths. No elevation needed.

.PARAMETER Force
    Overwrite existing module versions if already installed.

.PARAMETER ModuleName
    Top-level module name used for final verification. Defaults to 'Microsoft.Graph'.

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline'
    Installs to both PS 5.1 and PS 7 AllUsers paths (requires elevation).

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline' -Scope CurrentUser
    Installs to both PS 5.1 and PS 7 CurrentUser paths. No elevation needed.

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline' -Force
    Reinstalls/upgrades, overwriting existing versions in both paths.

.NOTES
    Requires PowerShell 5.1 or later. Administrator rights needed for AllUsers scope.
    Log written to C:\softdist\Logs\
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$SourcePath,

    [Parameter()]
    [ValidateSet('AllUsers','CurrentUser')]
    [string]$Scope = 'AllUsers',

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [string]$ModuleName = 'Microsoft.Graph'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Logging ───────────────────────────────────────────────────────────────────
$LogDir  = 'C:\softdist\Logs'
$LogFile = Join-Path $LogDir "Install-GraphModule_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if     ($Level -eq 'ERROR') { Write-Error   $Message }
    elseif ($Level -eq 'WARN')  { Write-Warning $Message }
    else                        { Write-Host    $entry   }
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

# ── Pre-flight ────────────────────────────────────────────────────────────────
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

Write-Host "`n=== Microsoft.Graph Offline Installer ===" -ForegroundColor Cyan
Write-Log "Source path : $SourcePath"
Write-Log "Scope       : $Scope"
Write-Log "Force       : $($Force.IsPresent)"

# Elevation check for AllUsers
if ($Scope -eq 'AllUsers') {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "AllUsers scope requires Administrator. Re-run elevated or use -Scope CurrentUser." -Level ERROR
        exit 1
    }
}

# ── Target paths — always install to both PS 5.1 and PS 7 ────────────────────
$installTargets = if ($Scope -eq 'AllUsers') {
    @(
        "$env:ProgramFiles\WindowsPowerShell\Modules",   # PS 5.1
        "$env:ProgramFiles\PowerShell\Modules"           # PS 7+
    )
} else {
    @(
        "$HOME\Documents\WindowsPowerShell\Modules",     # PS 5.1
        "$HOME\Documents\PowerShell\Modules"             # PS 7+
    )
}

foreach ($target in $installTargets) {
    if (-not (Test-Path $target)) {
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        Write-Log "Created: $target"
    }
}

Write-Log "Install targets:"
foreach ($target in $installTargets) { Write-Log "  -> $target" }

# ── Discover modules in source ────────────────────────────────────────────────
$modulefolders = @(Get-ChildItem -Path $SourcePath -Directory | Sort-Object Name)

if ($modulefolders.Count -eq 0) {
    Write-Log "No module folders found in '$SourcePath'. Verify the path is correct." -Level ERROR
    exit 1
}

Write-Log "Found $($modulefolders.Count) module folder(s) to install."

# ── Copy modules — sub-modules first, meta-module last ───────────────────────
$results   = [System.Collections.ArrayList]::new()
$installed = 0
$skipped   = 0
$failed    = 0

$subModules = @($modulefolders | Where-Object { $_.Name -ne $ModuleName })
$metaModule = @($modulefolders | Where-Object { $_.Name -eq $ModuleName })
$ordered    = @($subModules) + @($metaModule) | Where-Object { $_ -ne $null }

foreach ($moduleFolder in $ordered) {
    $name = $moduleFolder.Name

    $versionFolders = @(Get-ChildItem -Path $moduleFolder.FullName -Directory | Sort-Object Name)

    if ($versionFolders.Count -eq 0) {
        Write-Log "  SKIP: $name - no version subfolders found." -Level WARN
        $null = $results.Add([PSCustomObject]@{ Module = $name; Version = 'unknown'; Target = 'N/A'; Status = 'Skipped - no version folder' })
        $skipped++
        continue
    }

    foreach ($versionFolder in $versionFolders) {
        $version = $versionFolder.Name

        foreach ($installBase in $installTargets) {
            $destModule = Join-Path $installBase $name
            $destVer    = Join-Path $destModule $version

            if ((Test-Path $destVer) -and -not $Force) {
                Write-Log "  SKIP: $name v$version in $installBase (already present)"
                $null = $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Target = $installBase; Status = 'Skipped' })
                $skipped++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$name v$version -> $installBase", 'Copy module')) {
                try {
                    if (-not (Test-Path $destModule)) {
                        New-Item -ItemType Directory -Path $destModule -Force | Out-Null
                    }
                    Copy-Item -Path $versionFolder.FullName -Destination $destVer -Recurse -Force -ErrorAction Stop
                    Write-Log "  OK : $name v$version -> $installBase"
                    $null = $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Target = $installBase; Status = 'Installed' })
                    $installed++
                } catch {
                    Write-Log "  FAIL: $name v$version -> $installBase | $_" -Level WARN
                    $null = $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Target = $installBase; Status = "Failed: $_" })
                    $failed++
                }
            }
        }
    }
}

# ── Verify ────────────────────────────────────────────────────────────────────
Write-Log "Verifying $ModuleName..."
$graphModule = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
if ($graphModule) {
    Write-Log "Verification OK: $ModuleName v$($graphModule.Version) at $($graphModule.ModuleBase)"
} else {
    Write-Log "Verification WARNING: '$ModuleName' not found. Try opening a new PowerShell session." -Level WARN
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n=== Installation Summary ===" -ForegroundColor Cyan
Write-Host "Installed : $installed"
Write-Host "Skipped   : $skipped  (already present; use -Force to overwrite)"
Write-Host "Failed    : $failed"
Write-Host ""
$results | Format-Table Module, Version, Status -AutoSize -GroupBy Target
Write-Host "Log: $LogFile" -ForegroundColor Gray

if ($failed -gt 0) {
    Write-Host "`nSome modules failed. Check the log for details." -ForegroundColor Yellow
    exit 1
}

Write-Host "`nDone. Test with:" -ForegroundColor Green
Write-Host "  Import-Module Microsoft.Graph" -ForegroundColor Green
Write-Host "  Get-Module Microsoft.Graph -ListAvailable" -ForegroundColor Green

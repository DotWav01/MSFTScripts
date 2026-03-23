<#
.SYNOPSIS
    Installs a PowerShell module and dependencies from a local offline package.

.DESCRIPTION
    Installs any module and its dependencies from a folder previously populated
    by Save-PSModuleOffline.ps1. No internet or PSGallery access is required.

    Copies module folders directly into the PSModulePath for both PowerShell 5.1
    and PowerShell 7+ locations so the module is available in either version.

    Expected source structure (produced by Save-PSModuleOffline.ps1):
        SourcePath\
            <ModuleName>\
                <Version>\
                    <module files>
            <DependencyModule>\
                <Version>\
                    <module files>

.PARAMETER ModuleName
    The top-level module name being installed. Used for ordering (installed last,
    after all dependencies) and for post-install verification.
    Must match the folder name in SourcePath.

.PARAMETER SourcePath
    Path to the folder containing the downloaded module files.
    This is the OutputPath from Save-PSModuleOffline.ps1.

.PARAMETER Scope
    Installation scope:
    - AllUsers    (default): ProgramFiles paths. Requires elevation.
    - CurrentUser          : User Documents paths. No elevation needed.

.PARAMETER Force
    Overwrite existing module versions if already installed.

.EXAMPLE
    .\Install-PSModuleOffline.ps1 -ModuleName 'Microsoft.Graph' -SourcePath 'C:\Staging\Microsoft.Graph_Offline'
    Installs Microsoft.Graph to both PS 5.1 and PS 7 AllUsers paths.

.EXAMPLE
    .\Install-PSModuleOffline.ps1 -ModuleName 'ExchangeOnlineManagement' -SourcePath 'D:\Modules\EXO' -Scope CurrentUser
    Installs ExchangeOnlineManagement for current user only. No elevation needed.

.EXAMPLE
    .\Install-PSModuleOffline.ps1 -ModuleName 'Az' -SourcePath 'D:\Modules\Az' -Force
    Reinstalls/upgrades Az, overwriting any existing versions.

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - Administrator rights for AllUsers scope
    - SourcePath must be the output of Save-PSModuleOffline.ps1

    Installs to both PS 5.1 and PS 7+ paths simultaneously so no re-run is needed.
    Log written to C:\softdist\Logs\
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$ModuleName,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$SourcePath,

    [Parameter()]
    [ValidateSet('AllUsers','CurrentUser')]
    [string]$Scope = 'AllUsers',

    [Parameter()]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Logging ───────────────────────────────────────────────────────────────────
$LogDir  = 'C:\softdist\Logs'
$LogFile = Join-Path $LogDir "Install-${ModuleName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

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

Write-Host "`n=== PSModule Offline Installer ===" -ForegroundColor Cyan
Write-Log "Module      : $ModuleName"
Write-Log "Source path : $SourcePath"
Write-Log "Scope       : $Scope"
Write-Log "Force       : $($Force.IsPresent)"

# Elevation check
if ($Scope -eq 'AllUsers') {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "AllUsers scope requires Administrator. Re-run elevated or use -Scope CurrentUser." -Level ERROR
        exit 1
    }
}

# ── Install targets — both PS 5.1 and PS 7+ ──────────────────────────────────
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
        Write-Log "Created directory: $target"
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

Write-Log "Found $($modulefolders.Count) module folder(s) in source."

# ── Order: dependencies first, top-level module last ─────────────────────────
$subModules = @($modulefolders | Where-Object { $_.Name -ne $ModuleName })
$metaModule = @($modulefolders | Where-Object { $_.Name -eq $ModuleName })
$ordered    = @($subModules) + @($metaModule) | Where-Object { $_ -ne $null }

# ── Copy ──────────────────────────────────────────────────────────────────────
$results   = [System.Collections.ArrayList]::new()
$installed = 0
$skipped   = 0
$failed    = 0

foreach ($moduleFolder in $ordered) {
    $name           = $moduleFolder.Name
    $versionFolders = @(Get-ChildItem -Path $moduleFolder.FullName -Directory | Sort-Object Name)

    if ($versionFolders.Count -eq 0) {
        Write-Log "  SKIP: $name - no version subfolder found (unexpected structure)." -Level WARN
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
                Write-Log "  SKIP: $name v$version -> $(Split-Path $installBase -Leaf) (already present)"
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
Write-Log "Verifying '$ModuleName' is available..."
$verifyModule = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
if ($verifyModule) {
    Write-Log "Verification OK: $ModuleName v$($verifyModule.Version) at $($verifyModule.ModuleBase)"
} else {
    Write-Log "Verification WARNING: '$ModuleName' not found. Try opening a new PowerShell session." -Level WARN
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n=== Installation Summary ===" -ForegroundColor Cyan
Write-Host "Module    : $ModuleName"
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
Write-Host "  Import-Module $ModuleName" -ForegroundColor Green
Write-Host "  Get-Module $ModuleName -ListAvailable" -ForegroundColor Green

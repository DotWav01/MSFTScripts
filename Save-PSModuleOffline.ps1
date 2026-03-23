<#
.SYNOPSIS
    Downloads a PowerShell module and all dependencies for offline installation.

.DESCRIPTION
    Saves any PSGallery module and its full dependency tree to a local directory
    for transfer to an air-gapped or PSGallery-restricted server.
    Must be run on an internet-connected machine with access to PSGallery.

    Output folder structure (created by Save-Module):
        OutputPath\
            <ModuleName>\
                <Version>\
                    <module files>
            <DependencyModule>\
                <Version>\
                    <module files>

.PARAMETER ModuleName
    The module to download. Accepts any valid PSGallery module name.
    Examples: 'Microsoft.Graph', 'ExchangeOnlineManagement', 'Az', 'Microsoft.Graph.Users'

.PARAMETER OutputPath
    Directory where module files will be saved. Created if it does not exist.
    Defaults to '.\<ModuleName>_Offline' in the current directory.

.PARAMETER RequiredVersion
    Specific version to download. If omitted, downloads the latest stable version.

.PARAMETER Repository
    PSGallery repository to download from. Defaults to 'PSGallery'.

.EXAMPLE
    .\Save-PSModuleOffline.ps1 -ModuleName 'Microsoft.Graph'
    Downloads the latest Microsoft.Graph and all dependencies to .\Microsoft.Graph_Offline\

.EXAMPLE
    .\Save-PSModuleOffline.ps1 -ModuleName 'ExchangeOnlineManagement' -OutputPath 'D:\Staging\EXO'
    Downloads ExchangeOnlineManagement to a custom path.

.EXAMPLE
    .\Save-PSModuleOffline.ps1 -ModuleName 'Microsoft.Graph' -RequiredVersion '2.20.0'
    Downloads a specific version of Microsoft.Graph.

.EXAMPLE
    .\Save-PSModuleOffline.ps1 -ModuleName 'Az' -OutputPath 'D:\Staging\Az'
    Downloads the full Az module suite and dependencies.

.NOTES
    Requirements:
    - Internet access to PSGallery
    - PowerShellGet v2+ recommended (v1.x may miss transitive dependencies)
    - No elevation required for Save-Module

    To update PowerShellGet before running:
        Install-Module PowerShellGet -Force -AllowClobber -Scope CurrentUser

    After download, copy the entire OutputPath folder to the target server
    and run Install-PSModuleOffline.ps1.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ModuleName,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$RequiredVersion,

    [Parameter()]
    [string]$Repository = 'PSGallery'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Default output path to .\<ModuleName>_Offline if not specified
if (-not $OutputPath) {
    $OutputPath = Join-Path '.' "${ModuleName}_Offline"
}

# ── Pre-flight ────────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# ── Logging ───────────────────────────────────────────────────────────────────
$LogFile = Join-Path $OutputPath "Download_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if     ($Level -eq 'ERROR') { Write-Error   $Message }
    elseif ($Level -eq 'WARN')  { Write-Warning $Message }
    else                        { Write-Host    $entry   }
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

Write-Host "`n=== PSModule Offline Downloader ===" -ForegroundColor Cyan
Write-Log "Module      : $ModuleName"
Write-Log "Repository  : $Repository"
Write-Log "Output path : $(Resolve-Path $OutputPath)"

# ── Verify PowerShellGet ──────────────────────────────────────────────────────
$psGet = Get-Module PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
if (-not $psGet) {
    Write-Log "PowerShellGet module not found. Install it first." -Level ERROR
    exit 1
}
Write-Log "PowerShellGet version: $($psGet.Version)"
if ($psGet.Version -lt [version]'2.0') {
    Write-Log "PowerShellGet v2.0+ recommended (current: $($psGet.Version)). Run: Install-Module PowerShellGet -Force -AllowClobber -Scope CurrentUser" -Level WARN
}

# ── Verify repository is reachable ────────────────────────────────────────────
Write-Log "Checking connectivity to '$Repository'..."
try {
    $null = Find-Module -Name 'PowerShellGet' -Repository $Repository -ErrorAction Stop
    Write-Log "Repository '$Repository' is reachable."
} catch {
    Write-Log "Cannot reach repository '$Repository'. Check internet connectivity. $_" -Level ERROR
    exit 1
}

# ── Resolve version ───────────────────────────────────────────────────────────
if ($RequiredVersion) {
    Write-Log "Targeting version : $RequiredVersion"
} else {
    try {
        $found = Find-Module -Name $ModuleName -Repository $Repository -ErrorAction Stop
        $RequiredVersion = $found.Version
        Write-Log "Latest version    : $RequiredVersion"
    } catch {
        Write-Log "Could not find module '$ModuleName' in '$Repository'. Verify the module name is correct. $_" -Level ERROR
        exit 1
    }
}

# ── Download ──────────────────────────────────────────────────────────────────
Write-Log "Saving '$ModuleName' v$RequiredVersion and all dependencies..."
Write-Host "`nDownloading — this may take several minutes for large modules.`n" -ForegroundColor Yellow

try {
    Save-Module -Name $ModuleName `
                -Path $OutputPath `
                -Repository $Repository `
                -RequiredVersion $RequiredVersion `
                -Force `
                -ErrorAction Stop
    Write-Log "Download completed successfully."
} catch {
    Write-Log "Save-Module failed: $_" -Level ERROR
    exit 1
}

# ── Summary ───────────────────────────────────────────────────────────────────
$savedModules = @(Get-ChildItem -Path $OutputPath -Directory)
$totalSize    = (Get-ChildItem -Path $OutputPath -Recurse -File | Measure-Object -Property Length -Sum).Sum
$totalSizeMB  = [Math]::Round($totalSize / 1MB, 1)

Write-Host "`n=== Download Summary ===" -ForegroundColor Cyan
Write-Host "Module             : $ModuleName v$RequiredVersion"
Write-Host "Packages saved     : $($savedModules.Count)"
Write-Host "Total size         : ${totalSizeMB} MB"
Write-Host "Output path        : $(Resolve-Path $OutputPath)"
Write-Host "`nPackages included:" -ForegroundColor Cyan
$savedModules | Sort-Object Name | ForEach-Object { Write-Host "  - $($_.Name)" }

Write-Host "`nNext steps:" -ForegroundColor Green
Write-Host "  1. Copy the folder '$((Resolve-Path $OutputPath).Path)' to the target server."
Write-Host "  2. On the target server, run:"
Write-Host "       .\Install-PSModuleOffline.ps1 -ModuleName '$ModuleName' -SourcePath '<copied folder path>'"
Write-Host "`nLog: $LogFile"

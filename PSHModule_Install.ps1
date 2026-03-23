<#
.SYNOPSIS
    Installs the Microsoft.Graph module and dependencies from a local offline package.

.DESCRIPTION
    Installs Microsoft.Graph and all its dependencies from a folder previously populated
    by Save-GraphModuleOffline.ps1. No internet or PSGallery access is required.

    Uses direct folder copy into the PSModulePath — does NOT use Register-PSRepository,
    which requires a NuGet v2 feed structure and causes "no match found" errors with
    raw Save-Module output folders.

    Supports two installation scopes:
    - AllUsers   : Installs to $env:ProgramFiles\WindowsPowerShell\Modules (requires elevation)
    - CurrentUser: Installs to $HOME\Documents\WindowsPowerShell\Modules (no elevation needed)

.PARAMETER SourcePath
    Path to the folder containing the downloaded module files (output of Save-GraphModuleOffline.ps1).
    Expected structure: SourcePath\ModuleName\Version\<module files>

.PARAMETER Scope
    Installation scope. 'AllUsers' (default) or 'CurrentUser'.
    AllUsers requires the script to be run as Administrator.

.PARAMETER Force
    Overwrite existing module versions if already installed.

.PARAMETER ModuleName
    The top-level module name. Used only for final verification. Defaults to 'Microsoft.Graph'.

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline'
    Installs all modules from the staging folder for all users (requires elevation).

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline' -Scope CurrentUser
    Installs for the current user only. No elevation required.

.EXAMPLE
    .\Install-GraphModuleOffline.ps1 -SourcePath 'C:\Staging\GraphModuleOffline' -Force -Verbose
    Reinstalls/upgrades all modules, overwriting existing versions.

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - Administrator rights if using -Scope AllUsers
    - SourcePath must be the output folder from Save-GraphModuleOffline.ps1

    Save-Module output structure:
        SourcePath\
            Microsoft.Graph\
                2.36.1\
                    Microsoft.Graph.psd1
                    ...
            Microsoft.Graph.Authentication\
                2.36.1\
                    ...

    Log file is written to C:\softdist\Logs\
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

# Logging
$LogDir  = 'C:\softdist\Logs'
$LogFile = Join-Path $LogDir "Install-GraphModule_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Verbose $entry
    if     ($Level -eq 'ERROR') { Write-Error   $Message }
    elseif ($Level -eq 'WARN')  { Write-Warning $Message }
    else                        { Write-Host    $entry   }
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

# Pre-flight
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

Write-Host "`n=== Microsoft.Graph Offline Installer ===" -ForegroundColor Cyan
Write-Log "Source path  : $SourcePath"
Write-Log "Scope        : $Scope"
Write-Log "Force        : $Force"

# Elevation check for AllUsers scope
if ($Scope -eq 'AllUsers') {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "AllUsers scope requires Administrator. Re-run elevated or use -Scope CurrentUser." -Level ERROR
        exit 1
    }
}

# Resolve install destination — handles both PS 5.1 and PS 7+ paths
$psEdition   = $PSVersionTable.PSVersion.Major
$installBase = if ($Scope -eq 'AllUsers') {
    if ($psEdition -ge 6) { "$env:ProgramFiles\PowerShell\Modules" }
    else                  { "$env:ProgramFiles\WindowsPowerShell\Modules" }
} else {
    if ($psEdition -ge 6) { "$HOME\Documents\PowerShell\Modules" }
    else                  { "$HOME\Documents\WindowsPowerShell\Modules" }
}

Write-Log "Install destination: $installBase (PS $psEdition)"

if (-not (Test-Path $installBase)) {
    if ($PSCmdlet.ShouldProcess($installBase, 'Create module directory')) {
        New-Item -ItemType Directory -Path $installBase -Force | Out-Null
        Write-Log "Created module directory: $installBase"
    }
}

# Discover modules in source
# Save-Module creates: SourcePath\<ModuleName>\<Version>\<files>
$modulefolders = Get-ChildItem -Path $SourcePath -Directory | Sort-Object Name

if ($modulefolders.Count -eq 0) {
    Write-Log "No module folders found in '$SourcePath'. Verify the path is correct." -Level ERROR
    exit 1
}

Write-Log "Found $($modulefolders.Count) module folder(s) to install."

# Copy modules — sub-modules first, meta-module last
$results   = [System.Collections.Generic.List[PSCustomObject]]::new()
$installed = 0
$skipped   = 0
$failed    = 0

$subModules = $modulefolders | Where-Object { $_.Name -ne $ModuleName }
$metaModule = $modulefolders | Where-Object { $_.Name -eq $ModuleName }
$ordered    = @($subModules) + @($metaModule) | Where-Object { $_ }

foreach ($moduleFolder in $ordered) {
    $name = $moduleFolder.Name

    # Each module folder contains version subfolders (e.g. 2.36.1\)
    $versionFolders = Get-ChildItem -Path $moduleFolder.FullName -Directory | Sort-Object Name

    if ($versionFolders.Count -eq 0) {
        Write-Log "  SKIP: $name - no version subfolders found (unexpected structure)." -Level WARN
        $results.Add([PSCustomObject]@{ Module = $name; Version = 'unknown'; Status = 'Skipped - no version folder' })
        $skipped++
        continue
    }

    foreach ($versionFolder in $versionFolders) {
        $version    = $versionFolder.Name
        $destModule = Join-Path $installBase $name
        $destVer    = Join-Path $destModule $version

        # Skip if already installed at this version (unless -Force)
        if ((Test-Path $destVer) -and -not $Force) {
            Write-Log "  SKIP: $name v$version (already present)"
            $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Status = 'Skipped' })
            $skipped++
            continue
        }

        if ($PSCmdlet.ShouldProcess("$name v$version", "Copy to $destVer")) {
            try {
                if (-not (Test-Path $destModule)) {
                    New-Item -ItemType Directory -Path $destModule -Force | Out-Null
                }

                Copy-Item -Path $versionFolder.FullName -Destination $destVer -Recurse -Force -ErrorAction Stop

                Write-Log "  OK : $name v$version"
                $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Status = 'Installed' })
                $installed++
            } catch {
                Write-Log "  FAIL: $name v$version - $_" -Level WARN
                $results.Add([PSCustomObject]@{ Module = $name; Version = $version; Status = "Failed: $_" })
                $failed++
            }
        }
    }
}

# Verify
Write-Log "Verifying $ModuleName availability..."
$graphModule = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
if ($graphModule) {
    Write-Log "Verification OK: $ModuleName v$($graphModule.Version) found at $($graphModule.ModuleBase)"
} else {
    Write-Log "Verification WARNING: '$ModuleName' not found. You may need to open a new PowerShell session." -Level WARN
}

# Summary
Write-Host "`n=== Installation Summary ===" -ForegroundColor Cyan
Write-Host "Installed : $installed"
Write-Host "Skipped   : $skipped  (already present; use -Force to overwrite)"
Write-Host "Failed    : $failed"
Write-Host ""
$results | Format-Table Module, Version, Status -AutoSize
Write-Host "Log: $LogFile" -ForegroundColor Gray

if ($failed -gt 0) {
    Write-Host "`nSome modules failed. Check the log for details." -ForegroundColor Yellow
    exit 1
}

Write-Host "`nDone. Test with:" -ForegroundColor Green
Write-Host "  Import-Module Microsoft.Graph" -ForegroundColor Green
Write-Host "  Get-Module Microsoft.Graph -ListAvailable" -ForegroundColor Green

#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Resets specific UWP apps, ensures Windows Search service is running,
    and resets the Windows Search Box.

.DESCRIPTION
    - Re-registers: MicrosoftWindows.Client.WebExperience,
                    Microsoft.BingSearch,
                    MicrosoftWindows.Client.CBS
    - Ensures the WSearch (Windows Search) service is running
    - Resets the Windows Search Box (clears cache, kills searchui/searchapp)
    - Logs all actions to a timestamped log file

.NOTES
    Must be run as Administrator.
    Execution Policy: Run with -ExecutionPolicy Bypass if needed.
    Example: PowerShell.exe -ExecutionPolicy Bypass -File "Reset-SearchAndUWP.ps1"
#>

# ============================================================
#  EXECUTION POLICY CHECK
# ============================================================
$currentPolicy = Get-ExecutionPolicy -Scope Process
if ($currentPolicy -notin @('Bypass', 'Unrestricted', 'RemoteSigned')) {
    Write-Warning @"
=========================================================
  WARNING: Execution Policy may block this script.
  Current policy: $currentPolicy

  To run this script, use:
  PowerShell.exe -ExecutionPolicy Bypass -File "$($MyInvocation.MyCommand.Path)"

  Or set for current user:
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
=========================================================
"@
    $continue = Read-Host "Attempt to continue anyway? (Y/N)"
    if ($continue -ne 'Y') { Exit 1 }
}

# ============================================================
#  ELEVATION CHECK
# ============================================================
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $argList = "-ExecutionPolicy Bypass -NoExit -File `"$($MyInvocation.MyCommand.Path)`""
        Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList $argList
        Exit
    }
}

# ============================================================
#  LOGGING SETUP
# ============================================================
$LogDir  = "$env:SystemDrive\Logs\SearchReset"
$LogFile = "$LogDir\SearchReset_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

Write-Log "======================================================"
Write-Log " Reset-SearchAndUWP.ps1 started"
Write-Log " Log file: $LogFile"
Write-Log "======================================================"

# ============================================================
#  SECTION 1 - RE-REGISTER SPECIFIC UWP APPS
# ============================================================
Write-Log "--- SECTION 1: Re-registering UWP Apps ---"

$appsToReset = @(
    "MicrosoftWindows.Client.WebExperience",
    "Microsoft.BingSearch",
    "MicrosoftWindows.Client.CBS"
)

foreach ($appName in $appsToReset) {
    Write-Log "Looking up package: $appName"
    $packages = Get-AppXPackage -AllUsers -Name $appName -ErrorAction SilentlyContinue

    if (-not $packages) {
        Write-Log "  Package not found: $appName - skipping." "WARN"
        continue
    }

    foreach ($pkg in $packages) {
        $manifest = "$($pkg.InstallLocation)\AppXManifest.xml"
        Write-Log "  Found: $($pkg.PackageFullName)"

        if (Test-Path $manifest) {
            try {
                Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction Stop
                Write-Log "  Successfully re-registered: $($pkg.PackageFullName)"
            }
            catch {
                Write-Log "  Failed to re-register: $($pkg.PackageFullName) - $($_.Exception.Message)" "ERROR"
            }
        }
        else {
            Write-Log "  AppXManifest.xml not found at: $manifest - skipping." "WARN"
        }
    }
}

# ============================================================
#  SECTION 2 - ENSURE WINDOWS SEARCH SERVICE IS RUNNING
# ============================================================
Write-Log "--- SECTION 2: Windows Search Service ---"

$searchService = Get-Service -Name "WSearch" -ErrorAction SilentlyContinue

if (-not $searchService) {
    Write-Log "WSearch service not found on this system." "WARN"
}
else {
    Write-Log "WSearch service status: $($searchService.Status)"

    if ($searchService.StartType -eq 'Disabled') {
        Write-Log "WSearch service is Disabled - re-enabling..." "WARN"
        Set-Service -Name "WSearch" -StartupType Automatic
        Write-Log "WSearch startup type set to Automatic."
    }

    if ($searchService.Status -ne 'Running') {
        try {
            Start-Service -Name "WSearch" -ErrorAction Stop
            Write-Log "WSearch service started successfully."
        }
        catch {
            Write-Log "Failed to start WSearch service: $($_.Exception.Message)" "ERROR"
        }
    }
    else {
        Write-Log "WSearch service is already running - no action needed."
    }
}

# ============================================================
#  SECTION 3 - RESET WINDOWS SEARCH BOX
#  (Original Microsoft script - Copyright 2019, Microsoft Corporation)
# ============================================================
Write-Log "--- SECTION 3: Resetting Windows Search Box ---"

function T-R {
    [CmdletBinding()]
    Param([String] $n)
    $o = Get-Item -LiteralPath $n -ErrorAction SilentlyContinue
    return ($o -ne $null)
}

function R-R {
    [CmdletBinding()]
    Param([String] $l)
    $m = T-R $l
    if ($m) {
        Remove-Item -Path $l -Recurse -ErrorAction SilentlyContinue
    }
}

function S-D {
    Write-Log "Removing Cortana/Search testability registry keys..."
    R-R "HKLM:\SOFTWARE\Microsoft\Cortana\Testability"
    R-R "HKLM:\SOFTWARE\Microsoft\Search\Testability"
}

function K-P {
    [CmdletBinding()]
    Param([String] $g)
    $h = Get-Process $g -ErrorAction SilentlyContinue
    $i = (Get-Date).AddSeconds(2)
    $k = Get-Date
    while ((($i - $k) -gt 0) -and $h) {
        $k = Get-Date
        $h = Get-Process $g -ErrorAction SilentlyContinue
        if ($h) {
            $h.CloseMainWindow() | Out-Null
            Stop-Process -Id $h.Id -Force
        }
        $h = Get-Process $g -ErrorAction SilentlyContinue
    }
}

function D-FF {
    [CmdletBinding()]
    Param([string[]] $e)
    foreach ($f in $e) {
        if (Test-Path -Path $f) {
            Remove-Item -Recurse -Force $f -ErrorAction SilentlyContinue
        }
    }
}

function D-W {
    Write-Log "Clearing Search/Cortana web cache folders..."
    $d = @(
        "$Env:localappdata\Packages\Microsoft.Cortana_8wekyb3d8bbwe\AC\AppCache",
        "$Env:localappdata\Packages\Microsoft.Cortana_8wekyb3d8bbwe\AC\INetCache",
        "$Env:localappdata\Packages\Microsoft.Cortana_8wekyb3d8bbwe\AC\INetCookies",
        "$Env:localappdata\Packages\Microsoft.Cortana_8wekyb3d8bbwe\AC\INetHistory",
        "$Env:localappdata\Packages\Microsoft.Windows.Cortana_cw5n1h2txyewy\AC\AppCache",
        "$Env:localappdata\Packages\Microsoft.Windows.Cortana_cw5n1h2txyewy\AC\INetCache",
        "$Env:localappdata\Packages\Microsoft.Windows.Cortana_cw5n1h2txyewy\AC\INetCookies",
        "$Env:localappdata\Packages\Microsoft.Windows.Cortana_cw5n1h2txyewy\AC\INetHistory",
        "$Env:localappdata\Packages\Microsoft.Search_8wekyb3d8bbwe\AC\AppCache",
        "$Env:localappdata\Packages\Microsoft.Search_8wekyb3d8bbwe\AC\INetCache",
        "$Env:localappdata\Packages\Microsoft.Search_8wekyb3d8bbwe\AC\INetCookies",
        "$Env:localappdata\Packages\Microsoft.Search_8wekyb3d8bbwe\AC\INetHistory",
        "$Env:localappdata\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\AC\AppCache",
        "$Env:localappdata\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\AC\INetCache",
        "$Env:localappdata\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\AC\INetCookies",
        "$Env:localappdata\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\AC\INetHistory"
    )
    D-FF $d
}

function R-L {
    [CmdletBinding()]
    Param([String] $c)
    Write-Log "Stopping process: $c"
    K-P $c 2>&1 | Out-Null
    D-W
    K-P $c 2>&1 | Out-Null
    Start-Sleep -Seconds 5
}

# Determine correct search process name
$searchProcess = "searchui"
$searchPkgPath = "$Env:localappdata\Packages\Microsoft.Windows.Search_cw5n1h2txyewy"
if (Test-Path -Path $searchPkgPath) {
    $searchProcess = "searchapp"
}

Write-Log "Search process identified as: $searchProcess"

S-D 2>&1 | Out-Null
R-L $searchProcess

Write-Log "Windows Search Box reset complete."

# ============================================================
#  DONE
# ============================================================
Write-Log "======================================================"
Write-Log " All sections complete."
Write-Log " Log saved to: $LogFile"
Write-Log "======================================================"

Write-Host ""
Write-Host "Press any key to exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null

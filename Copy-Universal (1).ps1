#Requires -Version 5.1

<#
.SYNOPSIS
    Universal file copy script with Robocopy - handles any path with spaces and special characters.

.DESCRIPTION
    This PowerShell script provides a robust wrapper around Robocopy that can handle any file path
    including those with spaces, special characters (&, %, !, ^, etc.). It preserves file integrity
    (timestamps, attributes, versioning) without attempting to copy permissions.

.PARAMETER SourcePath
    The source directory path to copy from. Supports any characters including spaces and special symbols.

.PARAMETER DestinationPath
    The destination directory path to copy to. Supports any characters including spaces and special symbols.

.PARAMETER ExcludeDirectories
    Array of directory names to exclude from the copy operation.

.PARAMETER ExcludeFiles
    Array of file patterns to exclude from the copy operation.

.PARAMETER LogDirectory
    Directory where log files will be stored. Defaults to user's Documents folder.

.PARAMETER TestMode
    If specified, shows what would be copied without actually copying (Robocopy /L flag).

.PARAMETER Verbose
    Provides detailed output during the operation.

.EXAMPLE
    .\Copy-Universal.ps1 -SourcePath "Z:\Share Documents\Strategy - FP&A" -DestinationPath "Y:\Share Documents\Strategy - FP&A"
    
    Copies the FP&A strategy files from Z: to Y: drive.

.EXAMPLE
    .\Copy-Universal.ps1 -SourcePath "\\server\share\Complex Folder! (2024) & More" -DestinationPath "D:\Backup\Complex Folder! (2024) & More" -TestMode
    
    Tests copying a folder with complex special characters without actually copying.

.EXAMPLE
    .\Copy-Universal.ps1 -SourcePath "C:\Source" -DestinationPath "D:\Backup" -ExcludeDirectories @("temp", "cache") -ExcludeFiles @("*.tmp", "*.log")
    
    Copies files while excluding specified directories and file patterns.

.NOTES
    Author: Systems Administrator
    Version: 1.0
    
    This script automatically handles:
    - Paths with spaces
    - Special characters (&, %, !, ^, |, <, >, etc.)
    - UNC paths
    - Long file paths
    - Network drives
    
    Robocopy options used:
    /E - Copy subdirectories including empty ones
    /COPY:DAT - Copy Data, Attributes, Timestamps (no permissions)
    /DCOPY:DAT - Copy directory Data, Attributes, Timestamps
    /R:3 - Retry 3 times on failure
    /W:5 - Wait 5 seconds between retries
    /MT:4 - Multi-threaded operation
    /NP - No progress percentage
    /TEE - Output to console and log
    /UNILOG - Unicode log file
    /XJD - Exclude junction points (directories)
    /XJF - Exclude junction points (files)
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Source directory path")]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,
    
    [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Destination directory path")]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath,
    
    [Parameter(HelpMessage = "Directories to exclude")]
    [string[]]$ExcludeDirectories = @(),
    
    [Parameter(HelpMessage = "File patterns to exclude")]
    [string[]]$ExcludeFiles = @(),
    
    [Parameter(HelpMessage = "Directory for log files")]
    [string]$LogDirectory = $null,
    
    [Parameter(HelpMessage = "Test mode - show what would be copied without copying")]
    [switch]$TestMode
)

# Initialize logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = if ($LogDirectory) { $LogDirectory } else { [Environment]::GetFolderPath("MyDocuments") }
$logFile = Join-Path -Path $logDir -ChildPath "FileCopy_$timestamp.log"

# Ensure log directory exists
if (-not (Test-Path $logDir)) {
    try {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Warning "Could not create log directory '$logDir'. Using temp directory."
        $logDir = $env:TEMP
        $logFile = Join-Path -Path $logDir -ChildPath "FileCopy_$timestamp.log"
    }
}

function Write-LogMessage {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console with color
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "INFO" { Write-Host $logEntry -ForegroundColor Cyan }
        default { Write-Host $logEntry }
    }
    
    # Write to log file
    try {
        $logEntry | Add-Content -Path $logFile -Encoding UTF8
    }
    catch {
        # Silently continue if logging fails
    }
}

function Test-PathAccess {
    param(
        [string]$Path,
        [string]$AccessType = "Read"
    )
    
    try {
        if ($AccessType -eq "Read") {
            $testResult = Test-Path -Path $Path -PathType Container
            return $testResult
        }
        elseif ($AccessType -eq "Write") {
            if (-not (Test-Path -Path $Path)) {
                # Try to create the directory
                New-Item -Path $Path -ItemType Directory -Force | Out-Null
                return $true
            }
            else {
                # Test write access by creating a temp file
                $testFile = Join-Path -Path $Path -ChildPath "writetest_$(Get-Random).tmp"
                try {
                    "test" | Out-File -FilePath $testFile -Encoding ASCII
                    Remove-Item -Path $testFile -Force
                    return $true
                }
                catch {
                    return $false
                }
            }
        }
    }
    catch {
        return $false
    }
}

function Get-PathInfo {
    param([string]$Path)
    
    if (Test-Path -Path $Path -PathType Container) {
        try {
            $items = Get-ChildItem -Path $Path -Force
            $folders = ($items | Where-Object { $_.PSIsContainer }).Count
            $files = ($items | Where-Object { -not $_.PSIsContainer }).Count
            return @{
                Exists = $true
                Folders = $folders
                Files = $files
                TotalItems = $folders + $files
            }
        }
        catch {
            return @{
                Exists = $true
                Folders = "Unknown"
                Files = "Unknown"
                TotalItems = "Access Denied"
            }
        }
    }
    else {
        return @{
            Exists = $false
            Folders = 0
            Files = 0
            TotalItems = 0
        }
    }
}

try {
    Write-LogMessage "=== Universal File Copy Operation Started ===" "INFO"
    Write-LogMessage "PowerShell Version: $($PSVersionTable.PSVersion)" "INFO"
    Write-LogMessage "Source: '$SourcePath'" "INFO"
    Write-LogMessage "Destination: '$DestinationPath'" "INFO"
    Write-LogMessage "Log File: '$logFile'" "INFO"
    
    if ($TestMode) {
        Write-LogMessage "RUNNING IN TEST MODE - No files will actually be copied" "WARNING"
    }
    
    Write-Host ""
    
    # Validate source path
    Write-LogMessage "Validating source path..." "INFO"
    $sourceInfo = Get-PathInfo -Path $SourcePath
    
    if (-not $sourceInfo.Exists) {
        Write-LogMessage "Source path does not exist: '$SourcePath'" "ERROR"
        Write-LogMessage "Please verify:" "ERROR"
        Write-LogMessage "  1. Network drives are connected" "ERROR"
        Write-LogMessage "  2. Path spelling is correct" "ERROR"
        Write-LogMessage "  3. You have read access to the path" "ERROR"
        return
    }
    
    Write-LogMessage "Source path validated successfully" "SUCCESS"
    Write-LogMessage "  Contents: $($sourceInfo.Folders) folders, $($sourceInfo.Files) files" "INFO"
    
    # Validate/prepare destination
    Write-LogMessage "Validating destination path..." "INFO"
    $destParent = Split-Path -Path $DestinationPath -Parent
    
    if (-not (Test-PathAccess -Path $destParent -AccessType "Write")) {
        Write-LogMessage "Cannot write to destination parent directory: '$destParent'" "ERROR"
        return
    }
    
    Write-LogMessage "Destination path validated successfully" "SUCCESS"
    
    # Build Robocopy arguments using PowerShell's argument handling
    $robocopyArgs = @()
    $robocopyArgs += $SourcePath
    $robocopyArgs += $DestinationPath
    $robocopyArgs += "/E"                    # Copy subdirectories including empty
    $robocopyArgs += "/COPY:DAT"            # Copy Data, Attributes, Timestamps
    $robocopyArgs += "/DCOPY:DAT"           # Copy directory metadata  
    $robocopyArgs += "/R:3"                 # 3 retries
    $robocopyArgs += "/W:5"                 # Wait 5 seconds
    $robocopyArgs += "/MT:4"                # Multi-threaded
    $robocopyArgs += "/NP"                  # No progress percentage
    $robocopyArgs += "/TEE"                 # Console and log output
    $robocopyArgs += "/UNILOG:$logFile"     # Unicode log file
    $robocopyArgs += "/XJD"                 # Exclude junction points (dirs)
    $robocopyArgs += "/XJF"                 # Exclude junction points (files)
    
    # Add exclusions if specified
    if ($ExcludeDirectories.Count -gt 0) {
        $robocopyArgs += "/XD"
        $robocopyArgs += $ExcludeDirectories
        Write-LogMessage "Excluding directories: $($ExcludeDirectories -join ', ')" "INFO"
    }
    
    if ($ExcludeFiles.Count -gt 0) {
        $robocopyArgs += "/XF"
        $robocopyArgs += $ExcludeFiles
        Write-LogMessage "Excluding files: $($ExcludeFiles -join ', ')" "INFO"
    }
    
    # Add test mode flag if specified
    if ($TestMode) {
        $robocopyArgs += "/L"
    }
    
    Write-LogMessage "Robocopy arguments prepared: $($robocopyArgs.Count) arguments" "INFO"
    Write-Verbose "Full command: robocopy $($robocopyArgs -join ' ')"
    
    Write-Host ""
    Write-LogMessage "Starting Robocopy operation..." "INFO"
    
    if ($PSCmdlet.ShouldProcess($DestinationPath, "Copy files from '$SourcePath'")) {
        $startTime = Get-Date
        
        # Execute Robocopy with proper argument passing
        $process = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -Wait -PassThru -NoNewWindow
        $exitCode = $process.ExitCode
        
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Host ""
        Write-LogMessage "=== Operation Results ===" "INFO"
        Write-LogMessage "Duration: $($duration.ToString('mm\:ss'))" "INFO"
        Write-LogMessage "Exit Code: $exitCode" "INFO"
        
        # Interpret Robocopy exit codes
        $success = $false
        $message = ""
        
        switch ($exitCode) {
            0 { 
                $message = "No files copied - all files are up to date"
                $success = $true
            }
            1 { 
                $message = "Files copied successfully"
                $success = $true
            }
            2 { 
                $message = "Some additional files/directories were detected"
                $success = $true
            }
            3 { 
                $message = "Files copied successfully and additional files detected"
                $success = $true
            }
            4 { 
                $message = "Some mismatched files or directories were detected"
                $success = $false
            }
            5 { 
                $message = "Some files were copied and some mismatches detected"
                $success = $false
            }
            6 { 
                $message = "Additional and mismatched files/directories exist"
                $success = $false
            }
            7 { 
                $message = "Files copied, but additional and mismatched items exist"
                $success = $false
            }
            8 { 
                $message = "Some files or directories could not be copied"
                $success = $false
            }
            16 { 
                $message = "Fatal error - Robocopy did not copy any files"
                $success = $false
            }
            default { 
                $message = "Operation completed with exit code $exitCode"
                $success = ($exitCode -band 16) -eq 0  # Success if bit 4 (fatal error) is not set
            }
        }
        
        if ($success) {
            Write-LogMessage $message "SUCCESS"
        }
        else {
            Write-LogMessage $message "WARNING"
            Write-LogMessage "Check the detailed log for more information: '$logFile'" "WARNING"
        }
        
        # Return summary object
        $result = [PSCustomObject]@{
            SourcePath = $SourcePath
            DestinationPath = $DestinationPath
            ExitCode = $exitCode
            Success = $success
            Message = $message
            Duration = $duration
            LogFile = $logFile
            StartTime = $startTime
            EndTime = $endTime
            TestMode = $TestMode.IsPresent
        }
        
        Write-Host ""
        Write-LogMessage "Detailed log available at: '$logFile'" "INFO"
        Write-LogMessage "Operation completed successfully" "SUCCESS"
        
        return $result
    }
}
catch {
    Write-LogMessage "Critical error during copy operation: $($_.Exception.Message)" "ERROR"
    Write-LogMessage "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    throw
}

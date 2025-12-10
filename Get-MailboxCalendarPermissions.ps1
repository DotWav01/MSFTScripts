<#
.SYNOPSIS
    Retrieves Exchange Online mailboxes with delegated calendar permissions and exports to CSV.

.DESCRIPTION
    This script connects to Exchange Online and retrieves all mailboxes with their calendar permissions.
    It identifies which users have delegated access to each mailbox's calendar and their permission levels.
    Results are exported to a CSV file for analysis.
    
    The script will:
    - Connect to Exchange Online using modern authentication
    - Get all user mailboxes in the tenant
    - Check calendar permissions for each mailbox
    - Identify delegated permissions (excluding default/anonymous)
    - Export results to CSV with detailed permission information

.PARAMETER OutputPath
    Specifies the path and filename for the CSV export. 
    Default: "MailboxCalendarPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

.PARAMETER MailboxFilter
    Optional filter to limit which mailboxes to check. Supports wildcards.
    Example: "*@contoso.com" or "john*"

.PARAMETER IncludeSystemPermissions
    Switch to include system/default permissions in the output (Default, Anonymous, etc.)
    By default, these are filtered out to focus on delegated permissions.

.PARAMETER MaxConcurrentJobs
    Maximum number of concurrent background jobs for processing mailboxes.
    Default: 10. Adjust based on your system capabilities and Exchange Online throttling.

.PARAMETER ProgressInterval
    How often to display progress updates (every N mailboxes processed).
    Default: 25

.EXAMPLE
    .\Get-MailboxCalendarPermissions.ps1
    
    Connects to Exchange Online and exports all mailbox calendar permissions to a timestamped CSV file.

.EXAMPLE
    .\Get-MailboxCalendarPermissions.ps1 -OutputPath "C:\Reports\CalendarPermissions.csv" -MailboxFilter "*@contoso.com"
    
    Exports calendar permissions for mailboxes ending with @contoso.com to a specific file path.

.EXAMPLE
    .\Get-MailboxCalendarPermissions.ps1 -IncludeSystemPermissions -MaxConcurrentJobs 5
    
    Includes system permissions in output and limits concurrent processing to 5 jobs.

.NOTES
    Author: Alexander
    Version: 1.0
    Created: December 2025
    
    Requirements:
    - ExchangeOnlineManagement PowerShell module
    - Exchange Online administrator permissions
    - Permission to read mailbox configurations
    
    Limitations:
    - Large tenants may take significant time to process
    - Exchange Online throttling may affect performance
    - Requires appropriate Exchange Online permissions

.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/
#>

#Requires -Module ExchangeOnlineManagement

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "MailboxCalendarPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory = $false)]
    [string]$MailboxFilter,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemPermissions,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 20)]
    [int]$MaxConcurrentJobs = 10,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$ProgressInterval = 25
)

# Initialize variables
$Results = @()
$ProcessedCount = 0
$ErrorCount = 0
$StartTime = Get-Date

# System/default permissions to exclude (unless IncludeSystemPermissions is specified)
$SystemPermissions = @('Default', 'Anonymous', 'NT AUTHORITY\SYSTEM', 'SELF')

Write-Host "Starting Exchange Online Mailbox Calendar Permissions Report" -ForegroundColor Green
Write-Host "Output Path: $OutputPath" -ForegroundColor Cyan
Write-Host "Start Time: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

try {
    # Check if Exchange Online module is available
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed. Please install it using: Install-Module -Name ExchangeOnlineManagement"
    }

    # Connect to Exchange Online
    Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Yellow
    try {
        # Try to use existing connection first
        $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $existingConnection) {
            Connect-ExchangeOnline -ShowBanner:$false
        }
        Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        throw "Failed to connect to Exchange Online: $($_.Exception.Message)"
    }

    # Get all mailboxes
    Write-Host "`nRetrieving mailboxes..." -ForegroundColor Yellow
    
    $GetMailboxParams = @{
        ResultSize = 'Unlimited'
        RecipientTypeDetails = 'UserMailbox'
    }
    
    if ($MailboxFilter) {
        $GetMailboxParams.Filter = "EmailAddresses -like '*$MailboxFilter*'"
        Write-Host "Applying filter: $MailboxFilter" -ForegroundColor Cyan
    }
    
    $Mailboxes = Get-Mailbox @GetMailboxParams | Sort-Object DisplayName
    
    if ($Mailboxes.Count -eq 0) {
        Write-Warning "No mailboxes found matching the criteria"
        return
    }
    
    Write-Host "Found $($Mailboxes.Count) mailboxes to process" -ForegroundColor Green

    # Function to process individual mailbox permissions
    $ProcessMailboxPermissions = {
        param($Mailbox, $IncludeSystem, $SystemPerms)
        
        $MailboxResults = @()
        $CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
        
        try {
            # Get calendar folder permissions
            $Permissions = Get-MailboxFolderPermission -Identity $CalendarPath -ErrorAction Stop
            
            foreach ($Permission in $Permissions) {
                # Skip system permissions unless specifically included
                $IsSystemPermission = $SystemPerms -contains $Permission.User.DisplayName
                
                if ($IncludeSystem -or -not $IsSystemPermission) {
                    $PermissionResult = [PSCustomObject]@{
                        MailboxDisplayName = $Mailbox.DisplayName
                        MailboxPrimaryEmail = $Mailbox.PrimarySmtpAddress
                        MailboxAlias = $Mailbox.Alias
                        MailboxSamAccountName = $Mailbox.SamAccountName
                        DelegateUser = $Permission.User.DisplayName
                        DelegateUserType = if ($Permission.User.RecipientType) { $Permission.User.RecipientType } else { "Unknown" }
                        PermissionLevel = $Permission.AccessRights -join '; '
                        IsSystemPermission = $IsSystemPermission
                        CalendarPath = $CalendarPath
                        ProcessedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    }
                    $MailboxResults += $PermissionResult
                }
            }
        }
        catch {
            # Handle errors for individual mailboxes
            $ErrorResult = [PSCustomObject]@{
                MailboxDisplayName = $Mailbox.DisplayName
                MailboxPrimaryEmail = $Mailbox.PrimarySmtpAddress
                MailboxAlias = $Mailbox.Alias
                MailboxSamAccountName = $Mailbox.SamAccountName
                DelegateUser = "ERROR"
                DelegateUserType = "ERROR"
                PermissionLevel = $_.Exception.Message
                IsSystemPermission = $false
                CalendarPath = $CalendarPath
                ProcessedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            $MailboxResults += $ErrorResult
        }
        
        return $MailboxResults
    }

    # Process mailboxes with background jobs for better performance
    Write-Host "`nProcessing mailbox permissions..." -ForegroundColor Yellow
    $Jobs = @()
    $JobBatch = @()
    
    foreach ($Mailbox in $Mailboxes) {
        # Create background job for each mailbox
        $Job = Start-Job -ScriptBlock $ProcessMailboxPermissions -ArgumentList $Mailbox, $IncludeSystemPermissions.IsPresent, $SystemPermissions
        $JobBatch += @{
            Job = $Job
            Mailbox = $Mailbox
        }
        
        # Process in batches to avoid too many concurrent jobs
        if ($JobBatch.Count -ge $MaxConcurrentJobs) {
            # Wait for this batch to complete
            $CompletedJobs = $JobBatch | ForEach-Object { $_.Job }
            Wait-Job -Job $CompletedJobs | Out-Null
            
            # Collect results and clean up
            foreach ($JobInfo in $JobBatch) {
                try {
                    $JobResults = Receive-Job -Job $JobInfo.Job -ErrorAction Stop
                    $Results += $JobResults
                    $ProcessedCount++
                }
                catch {
                    Write-Warning "Error processing mailbox $($JobInfo.Mailbox.DisplayName): $($_.Exception.Message)"
                    $ErrorCount++
                }
                finally {
                    Remove-Job -Job $JobInfo.Job -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Display progress
            if ($ProcessedCount -gt 0 -and ($ProcessedCount % $ProgressInterval -eq 0)) {
                $PercentComplete = [math]::Round(($ProcessedCount / $Mailboxes.Count) * 100, 1)
                Write-Host "Progress: $ProcessedCount/$($Mailboxes.Count) mailboxes processed ($PercentComplete%)" -ForegroundColor Cyan
            }
            
            $JobBatch = @()
        }
    }
    
    # Process remaining jobs in the last batch
    if ($JobBatch.Count -gt 0) {
        $CompletedJobs = $JobBatch | ForEach-Object { $_.Job }
        Wait-Job -Job $CompletedJobs | Out-Null
        
        foreach ($JobInfo in $JobBatch) {
            try {
                $JobResults = Receive-Job -Job $JobInfo.Job -ErrorAction Stop
                $Results += $JobResults
                $ProcessedCount++
            }
            catch {
                Write-Warning "Error processing mailbox $($JobInfo.Mailbox.DisplayName): $($_.Exception.Message)"
                $ErrorCount++
            }
            finally {
                Remove-Job -Job $JobInfo.Job -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # Clean up any remaining jobs
    Get-Job | Where-Object { $_.State -eq 'Completed' -or $_.State -eq 'Failed' } | Remove-Job -Force -ErrorAction SilentlyContinue

    # Export results to CSV
    Write-Host "`nExporting results to CSV..." -ForegroundColor Yellow
    
    if ($Results.Count -gt 0) {
        # Ensure output directory exists
        $OutputDirectory = Split-Path $OutputPath -Parent
        if ($OutputDirectory -and -not (Test-Path $OutputDirectory)) {
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        }
        
        $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
        
        # Display summary statistics
        Write-Host "`n" + "="*80 -ForegroundColor Magenta
        Write-Host "SUMMARY REPORT" -ForegroundColor Magenta
        Write-Host "="*80 -ForegroundColor Magenta
        
        $EndTime = Get-Date
        $Duration = $EndTime - $StartTime
        
        Write-Host "Processing completed: $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
        Write-Host "Total duration: $($Duration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
        Write-Host "Total mailboxes processed: $ProcessedCount" -ForegroundColor Green
        Write-Host "Total permission entries found: $($Results.Count)" -ForegroundColor Green
        Write-Host "Errors encountered: $ErrorCount" -ForegroundColor $(if ($ErrorCount -gt 0) { 'Red' } else { 'Green' })
        
        # Permission statistics
        $DelegatedPermissions = $Results | Where-Object { -not $_.IsSystemPermission }
        $SystemPermissionsFound = $Results | Where-Object { $_.IsSystemPermission }
        $UniqueMailboxesWithDelegation = ($DelegatedPermissions | Group-Object MailboxPrimaryEmail).Count
        $UniqueDelegates = ($DelegatedPermissions | Group-Object DelegateUser).Count
        
        Write-Host "`nPermission Statistics:" -ForegroundColor Yellow
        Write-Host "  Mailboxes with delegated permissions: $UniqueMailboxesWithDelegation" -ForegroundColor Cyan
        Write-Host "  Total delegated permission entries: $($DelegatedPermissions.Count)" -ForegroundColor Cyan
        Write-Host "  Unique delegates: $UniqueDelegates" -ForegroundColor Cyan
        Write-Host "  System permission entries: $($SystemPermissionsFound.Count)" -ForegroundColor Cyan
        
        # Top permission levels
        if ($DelegatedPermissions.Count -gt 0) {
            Write-Host "`nTop Permission Levels:" -ForegroundColor Yellow
            $DelegatedPermissions | Group-Object PermissionLevel | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
                Write-Host "  $($_.Name): $($_.Count) entries" -ForegroundColor Cyan
            }
        }
        
        Write-Host "`nOutput file: $OutputPath" -ForegroundColor Green
    }
    else {
        Write-Warning "No permission entries found to export"
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host "Stack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}
finally {
    # Clean up any remaining background jobs
    Get-Job | Where-Object { $_.Name -like "*" } | Remove-Job -Force -ErrorAction SilentlyContinue
    
    Write-Host "`nScript execution completed." -ForegroundColor Green
}

<#
.SYNOPSIS
    Retrieves Exchange Online mailboxes with delegated calendar permissions - Throttling Resistant Version

.DESCRIPTION
    Enhanced version with throttling protection, retry logic, and batch processing
    for large tenant environments. Includes automatic session refresh and error recovery.

.PARAMETER OutputPath
    Specifies the path and filename for the CSV export.

.PARAMETER InputCsvPath
    Path to CSV file containing list of users to check.

.PARAMETER CsvEmailColumn
    Name of the column in CSV containing email addresses. Default: "Email"

.PARAMETER IncludeSystemPermissions
    Switch to include system/default permissions in output.

.PARAMETER BatchSize
    Number of mailboxes to process before pausing. Default: 50

.PARAMETER DelayBetweenBatches
    Seconds to wait between batches to avoid throttling. Default: 30

.PARAMETER RetryAttempts
    Number of retry attempts for failed mailboxes. Default: 3

.PARAMETER SessionRefreshInterval
    Number of mailboxes to process before refreshing session. Default: 200

.EXAMPLE
    .\Get-MailboxCalendarPermissions-Enhanced.ps1 -BatchSize 25 -DelayBetweenBatches 45

.NOTES
    Author: Alexander
    Version: 2.0 - Throttling Resistant
    Created: December 2025
#>

#Requires -Module ExchangeOnlineManagement

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "MailboxCalendarPermissions_Enhanced_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if ($_ -and -not (Test-Path $_ -PathType Leaf)) {
            throw "CSV file not found: $_"
        }
        return $true
    })]
    [string]$InputCsvPath,
    
    [Parameter(Mandatory = $false)]
    [string]$CsvEmailColumn = "Email",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemPermissions,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 100)]
    [int]$BatchSize = 50,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 300)]
    [int]$DelayBetweenBatches = 30,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 5)]
    [int]$RetryAttempts = 3,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(50, 500)]
    [int]$SessionRefreshInterval = 200,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$ProgressInterval = 25
)

# Initialize variables
$Results = @()
$ProcessedCount = 0
$ErrorCount = 0
$RetryCount = 0
$StartTime = Get-Date
$LastSessionRefresh = Get-Date

# System/default permissions to exclude
$SystemPermissions = @('Default', 'Anonymous', 'NT AUTHORITY\SYSTEM', 'SELF')

Write-Host "Exchange Online Calendar Permissions - Enhanced Version" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "Output Path: $OutputPath" -ForegroundColor Cyan
Write-Host "Batch Size: $BatchSize mailboxes" -ForegroundColor Cyan
Write-Host "Delay Between Batches: $DelayBetweenBatches seconds" -ForegroundColor Cyan
Write-Host "Session Refresh Interval: $SessionRefreshInterval mailboxes" -ForegroundColor Cyan
if ($InputCsvPath) {
    Write-Host "Input CSV: $InputCsvPath" -ForegroundColor Cyan
}
Write-Host "Start Time: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

# Function to refresh Exchange Online session
function Refresh-ExchangeSession {
    Write-Host "`n🔄 Refreshing Exchange Online session..." -ForegroundColor Yellow
    try {
        # Test current session
        $null = Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1
        Write-Host "✓ Session is still valid" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "⚠ Session expired, reconnecting..." -ForegroundColor Yellow
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Start-Sleep -Seconds 5
            Write-Host "✓ Session refreshed successfully" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to refresh session: $($_.Exception.Message)"
            return $false
        }
    }
}

# Function to process a single mailbox with retry logic
function Process-MailboxWithRetry {
    param($Mailbox, $AttemptNumber = 1)
    
    try {
        $CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
        
        # Add small delay to avoid rapid-fire requests
        Start-Sleep -Milliseconds 500
        
        # Get calendar folder permissions with timeout
        $Permissions = Get-MailboxFolderPermission -Identity $CalendarPath -ErrorAction Stop
        
        $MailboxResults = @()
        foreach ($Permission in $Permissions) {
            # Get user name safely
            $UserName = ""
            $UserEmail = ""
            $UserUPN = ""
            
            if ($Permission.User) {
                if ($Permission.User.DisplayName) {
                    $UserName = $Permission.User.DisplayName
                } elseif ($Permission.User.ToString()) {
                    $UserName = $Permission.User.ToString()
                } else {
                    $UserName = $Permission.User
                }
                
                # Try to get email/UPN
                if ($Permission.User.RecipientPrincipal) {
                    $UserEmail = $Permission.User.RecipientPrincipal
                    $UserUPN = $Permission.User.RecipientPrincipal
                } elseif ($Permission.User.PrimarySmtpAddress) {
                    $UserEmail = $Permission.User.PrimarySmtpAddress
                } elseif ($Permission.User.EmailAddress) {
                    $UserEmail = $Permission.User.EmailAddress
                } elseif ($Permission.User.UserPrincipalName) {
                    $UserUPN = $Permission.User.UserPrincipalName
                    $UserEmail = $Permission.User.UserPrincipalName
                }
                
                if ([string]::IsNullOrEmpty($UserEmail) -and $UserName -match '@') {
                    $UserEmail = $UserName
                    $UserUPN = $UserName
                }
            } else {
                $UserName = "Unknown User"
            }
            
            # Get user type safely
            $UserType = "Unknown"
            if ($Permission.User -and $Permission.User.RecipientType) {
                $UserType = $Permission.User.RecipientType
            } elseif ($Permission.User -and $Permission.User.GetType) {
                $UserType = $Permission.User.GetType().Name
            }
            
            # Get access rights safely
            $AccessRights = "Unknown"
            if ($Permission.AccessRights) {
                if ($Permission.AccessRights -is [System.Collections.ArrayList] -or $Permission.AccessRights -is [array]) {
                    $AccessRights = ($Permission.AccessRights | ForEach-Object { $_.ToString() }) -join '; '
                } else {
                    $AccessRights = $Permission.AccessRights.ToString()
                }
            }
            
            if ($AccessRights -like "*System.Collections.ArrayList*" -or $AccessRights -eq "System.Collections.ArrayList") {
                try {
                    $AccessRights = [string]::Join("; ", $Permission.AccessRights)
                }
                catch {
                    $AccessRights = "AccessRights conversion error"
                }
            }
            
            # Determine if system permission
            $IsSystemPermission = $SystemPermissions -contains $UserName
            
            if ($IncludeSystemPermissions.IsPresent -or -not $IsSystemPermission) {
                $PermissionResult = [PSCustomObject]@{
                    MailboxDisplayName = $Mailbox.DisplayName
                    MailboxPrimaryEmail = $Mailbox.PrimarySmtpAddress
                    MailboxAlias = $Mailbox.Alias
                    MailboxSamAccountName = $Mailbox.SamAccountName
                    DelegateUser = $UserName
                    DelegateEmail = $UserEmail
                    DelegateUPN = $UserUPN
                    DelegateUserType = $UserType
                    PermissionLevel = $AccessRights
                    IsSystemPermission = $IsSystemPermission
                    CalendarPath = $CalendarPath
                    ProcessedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    AttemptNumber = $AttemptNumber
                }
                $MailboxResults += $PermissionResult
            }
        }
        
        return @{
            Success = $true
            Results = $MailboxResults
            Error = $null
        }
    }
    catch {
        return @{
            Success = $false
            Results = @()
            Error = $_.Exception.Message
        }
    }
}

try {
    # Check Exchange module
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module not installed"
    }

    # Connect to Exchange Online with verification
    Write-Host "`n🔌 Connecting to Exchange Online..." -ForegroundColor Yellow
    if (-not (Refresh-ExchangeSession)) {
        throw "Failed to establish Exchange Online connection"
    }

    # Get mailboxes (same logic as original script)
    Write-Host "`n📋 Retrieving mailboxes..." -ForegroundColor Yellow
    
    if ($InputCsvPath) {
        # CSV processing logic (same as original)
        $CsvData = Import-Csv -Path $InputCsvPath
        if (-not $CsvData) { throw "CSV file is empty" }
        
        $CsvColumns = $CsvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        if ($CsvEmailColumn -notin $CsvColumns) {
            throw "CSV column '$CsvEmailColumn' not found. Available: $($CsvColumns -join ', ')"
        }
        
        $Mailboxes = @()
        foreach ($Row in $CsvData) {
            $Email = $Row.$CsvEmailColumn
            if (-not [string]::IsNullOrWhiteSpace($Email)) {
                try {
                    $Mailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
                    $Mailboxes += $Mailbox
                }
                catch {
                    Write-Warning "Mailbox not found: $Email"
                }
            }
        }
    }
    else {
        # Get all mailboxes
        $TestMailbox = Get-Mailbox -ResultSize 1 -ErrorAction Stop
        $Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
    }
    
    if (-not $Mailboxes -or $Mailboxes.Count -eq 0) {
        throw "No mailboxes found"
    }
    
    $Mailboxes = $Mailboxes | Sort-Object DisplayName
    Write-Host "✓ Found $($Mailboxes.Count) mailboxes to process" -ForegroundColor Green

    # Process mailboxes in batches
    Write-Host "`n⚡ Processing mailboxes in batches..." -ForegroundColor Yellow
    
    for ($i = 0; $i -lt $Mailboxes.Count; $i += $BatchSize) {
        $BatchEnd = [Math]::Min($i + $BatchSize - 1, $Mailboxes.Count - 1)
        $BatchNumber = [Math]::Floor($i / $BatchSize) + 1
        $TotalBatches = [Math]::Ceiling($Mailboxes.Count / $BatchSize)
        
        Write-Host "`n📦 Processing Batch $BatchNumber of $TotalBatches (mailboxes $($i+1)-$($BatchEnd+1))" -ForegroundColor Cyan
        
        # Check if session refresh is needed
        if ($ProcessedCount -gt 0 -and ($ProcessedCount % $SessionRefreshInterval) -eq 0) {
            if (-not (Refresh-ExchangeSession)) {
                throw "Session refresh failed"
            }
            $LastSessionRefresh = Get-Date
        }
        
        # Process current batch
        for ($j = $i; $j -le $BatchEnd; $j++) {
            $Mailbox = $Mailboxes[$j]
            $Success = $false
            
            # Retry logic for individual mailboxes
            for ($attempt = 1; $attempt -le $RetryAttempts; $attempt++) {
                $Result = Process-MailboxWithRetry -Mailbox $Mailbox -AttemptNumber $attempt
                
                if ($Result.Success) {
                    $Results += $Result.Results
                    $Success = $true
                    break
                }
                else {
                    if ($attempt -lt $RetryAttempts) {
                        Write-Host "⚠ Retry $attempt for $($Mailbox.DisplayName): $($Result.Error)" -ForegroundColor Yellow
                        Start-Sleep -Seconds (5 * $attempt)  # Exponential backoff
                        $RetryCount++
                    }
                    else {
                        Write-Warning "✗ Failed $($Mailbox.DisplayName) after $RetryAttempts attempts: $($Result.Error)"
                        
                        # Add error entry
                        $ErrorResult = [PSCustomObject]@{
                            MailboxDisplayName = $Mailbox.DisplayName
                            MailboxPrimaryEmail = $Mailbox.PrimarySmtpAddress
                            MailboxAlias = $Mailbox.Alias
                            MailboxSamAccountName = $Mailbox.SamAccountName
                            DelegateUser = "ERROR"
                            DelegateEmail = "ERROR"
                            DelegateUPN = "ERROR"
                            DelegateUserType = "ERROR"
                            PermissionLevel = $Result.Error
                            IsSystemPermission = $false
                            CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
                            ProcessedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                            AttemptNumber = $RetryAttempts
                        }
                        $Results += $ErrorResult
                        $ErrorCount++
                    }
                }
            }
            
            $ProcessedCount++
            
            # Progress reporting
            if ($ProcessedCount % $ProgressInterval -eq 0) {
                $PercentComplete = [math]::Round(($ProcessedCount / $Mailboxes.Count) * 100, 1)
                $Elapsed = (Get-Date) - $StartTime
                $EstimatedTotal = $Elapsed.TotalSeconds / ($ProcessedCount / $Mailboxes.Count)
                $ETA = $StartTime.AddSeconds($EstimatedTotal)
                
                Write-Host "📊 Progress: $ProcessedCount/$($Mailboxes.Count) ($PercentComplete%) | ETA: $($ETA.ToString('HH:mm:ss')) | Errors: $ErrorCount | Retries: $RetryCount" -ForegroundColor Green
            }
        }
        
        # Delay between batches (except for last batch)
        if ($BatchEnd -lt ($Mailboxes.Count - 1)) {
            Write-Host "⏸ Pausing $DelayBetweenBatches seconds to avoid throttling..." -ForegroundColor Yellow
            Start-Sleep -Seconds $DelayBetweenBatches
        }
    }

    # Export results
    Write-Host "`n💾 Exporting results..." -ForegroundColor Yellow
    
    if ($Results.Count -gt 0) {
        $OutputDirectory = Split-Path $OutputPath -Parent
        if ($OutputDirectory -and -not (Test-Path $OutputDirectory)) {
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        }
        
        $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "✓ Results exported to: $OutputPath" -ForegroundColor Green
        
        # Summary
        $EndTime = Get-Date
        $Duration = $EndTime - $StartTime
        $DelegatedPermissions = $Results | Where-Object { -not ([bool]::Parse($_.IsSystemPermission)) -and $_.DelegateUser -ne "ERROR" }
        
        Write-Host "`n" + "="*80 -ForegroundColor Magenta
        Write-Host "📈 SUMMARY REPORT - ENHANCED VERSION" -ForegroundColor Magenta
        Write-Host "="*80 -ForegroundColor Magenta
        Write-Host "⏱ Total Duration: $($Duration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
        Write-Host "📋 Mailboxes Processed: $ProcessedCount" -ForegroundColor Green
        Write-Host "📊 Permission Entries Found: $($Results.Count)" -ForegroundColor Green
        Write-Host "👥 Delegated Permissions: $($DelegatedPermissions.Count)" -ForegroundColor Green
        Write-Host "❌ Errors: $ErrorCount" -ForegroundColor $(if ($ErrorCount -gt 0) { 'Red' } else { 'Green' })
        Write-Host "🔄 Retries: $RetryCount" -ForegroundColor Yellow
        Write-Host "📁 Output: $OutputPath" -ForegroundColor Green
        
        if ($DelegatedPermissions.Count -gt 0) {
            Write-Host "`n🏆 Top Permission Levels:" -ForegroundColor Yellow
            try {
                $PermissionGroups = $DelegatedPermissions | Group-Object PermissionLevel | Sort-Object Count -Descending | Select-Object -First 5
                foreach ($Group in $PermissionGroups) {
                    Write-Host "   $($Group.Name): $($Group.Count) entries" -ForegroundColor Cyan
                }
            }
            catch {
                Write-Host "   Summary statistics unavailable" -ForegroundColor Yellow
            }
        }
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
    Write-Host "`n🏁 Enhanced script execution completed." -ForegroundColor Green
}

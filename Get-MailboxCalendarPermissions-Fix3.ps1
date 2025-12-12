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

.PARAMETER InputCsvPath
    Path to CSV file containing list of users to check. CSV should have columns:
    - Email (required): Primary email address or UserPrincipalName
    - Name (optional): Display name for reference
    If this parameter is used, only mailboxes in the CSV will be processed.

.PARAMETER CsvEmailColumn
    Name of the column in the CSV containing email addresses.
    Default: "Email". Other common values: "UserPrincipalName", "PrimarySmtpAddress"

.PARAMETER IncludeSystemPermissions
    Switch to include system/default permissions in the output (Default, Anonymous, etc.)
    By default, these are filtered out to focus on delegated permissions.

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
    .\Get-MailboxCalendarPermissions.ps1 -InputCsvPath "C:\Lists\VIPUsers.csv" -CsvEmailColumn "UserPrincipalName"
    
    Processes only the mailboxes listed in the CSV file, using UserPrincipalName column for email addresses.

.EXAMPLE
    .\Get-MailboxCalendarPermissions.ps1 -IncludeSystemPermissions
    
    Includes system permissions in output.

.NOTES
    Author: Alexander
    Version: 1.0
    Created: December 2025
    
    Requirements:
    - ExchangeOnlineManagement PowerShell module
    - Exchange Online administrator permissions
    - Permission to read mailbox configurations
    
    Limitations:
    - Large tenants may take significant time to process (sequential processing)
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
if ($InputCsvPath) {
    Write-Host "Input CSV: $InputCsvPath" -ForegroundColor Cyan
    Write-Host "CSV Email Column: $CsvEmailColumn" -ForegroundColor Cyan
}
if ($MailboxFilter) {
    Write-Host "Mailbox Filter: $MailboxFilter" -ForegroundColor Cyan
}
Write-Host "Start Time: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

try {
    # Check if Exchange Online module is available
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed. Please install it using: Install-Module -Name ExchangeOnlineManagement"
    }

    # Connect to Exchange Online with proper verification
    Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Yellow
    $ConnectionVerified = $false
    
    try {
        # First, test if we can actually run Exchange cmdlets (not just check connection info)
        Write-Host "Testing current Exchange Online connectivity..." -ForegroundColor Yellow
        try {
            $testResult = Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1 -Property DisplayName
            if ($testResult) {
                Write-Host "✓ Existing connection is working" -ForegroundColor Green
                $ConnectionVerified = $true
            }
        }
        catch {
            Write-Host "○ No working connection found, need to authenticate" -ForegroundColor Yellow
        }
        
        # If no working connection, force a new connection
        if (-not $ConnectionVerified) {
            Write-Host "Initiating Exchange Online authentication..." -ForegroundColor Yellow
            Write-Host "⚠ You should see a sign-in prompt shortly..." -ForegroundColor Cyan
            
            # Disconnect any existing sessions to force fresh authentication
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
            catch {
                # Ignore disconnect errors
            }
            
            # Force new connection
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Write-Host "✓ Authentication completed" -ForegroundColor Green
            
            # Wait a moment for session to initialize
            Start-Sleep -Seconds 3
            
            # Verify the new connection works
            Write-Host "Verifying new connection..." -ForegroundColor Yellow
            $testResult = Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1 -Property DisplayName
            if ($testResult) {
                Write-Host "✓ Connection verified - Connected to: $($testResult.DisplayName)" -ForegroundColor Green
                $ConnectionVerified = $true
            }
        }
        
        if (-not $ConnectionVerified) {
            throw "Unable to establish working Exchange Online connection"
        }
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
        Write-Host "1. Ensure you have Exchange Online Administrator permissions" -ForegroundColor Cyan
        Write-Host "2. Check your MFA setup and authentication methods" -ForegroundColor Cyan
        Write-Host "3. Try running Connect-ExchangeOnline manually in a new PowerShell session" -ForegroundColor Cyan
        Write-Host "4. Verify your account has proper licensing" -ForegroundColor Cyan
        throw "Exchange Online connection failed"
    }

    # Get mailboxes - either from CSV or all mailboxes
    Write-Host "`nRetrieving mailboxes..." -ForegroundColor Yellow
    
    try {
        if ($InputCsvPath) {
            # Process mailboxes from CSV file
            Write-Host "Loading mailbox list from CSV: $InputCsvPath" -ForegroundColor Cyan
            
            # Import and validate CSV
            $CsvData = Import-Csv -Path $InputCsvPath
            if (-not $CsvData) {
                throw "CSV file is empty or could not be read"
            }
            
            # Validate required column exists
            $CsvColumns = $CsvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
            if ($CsvEmailColumn -notin $CsvColumns) {
                throw "CSV column '$CsvEmailColumn' not found. Available columns: $($CsvColumns -join ', ')"
            }
            
            Write-Host "Found $($CsvData.Count) entries in CSV" -ForegroundColor Green
            Write-Host "Using column '$CsvEmailColumn' for email addresses" -ForegroundColor Cyan
            
            # Get mailboxes for each email in the CSV
            $Mailboxes = @()
            $NotFoundEmails = @()
            $ProcessedEmails = @()
            
            foreach ($CsvRow in $CsvData) {
                $EmailAddress = $CsvRow.$CsvEmailColumn
                if ([string]::IsNullOrWhiteSpace($EmailAddress)) {
                    Write-Warning "Empty email address found in CSV row, skipping"
                    continue
                }
                
                # Skip duplicates
                if ($EmailAddress -in $ProcessedEmails) {
                    Write-Warning "Duplicate email address found: $EmailAddress, skipping"
                    continue
                }
                $ProcessedEmails += $EmailAddress
                
                try {
                    Write-Host "Looking up mailbox: $EmailAddress" -ForegroundColor Cyan
                    $Mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
                    $Mailboxes += $Mailbox
                }
                catch {
                    Write-Warning "Mailbox not found: $EmailAddress"
                    $NotFoundEmails += $EmailAddress
                }
            }
            
            if ($NotFoundEmails.Count -gt 0) {
                Write-Host "`nMailboxes not found ($($NotFoundEmails.Count)):" -ForegroundColor Yellow
                $NotFoundEmails | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
            }
            
            if ($Mailboxes.Count -eq 0) {
                Write-Warning "No valid mailboxes found from CSV input"
                return
            }
            
            Write-Host "Successfully found $($Mailboxes.Count) mailboxes from CSV" -ForegroundColor Green
        }
        else {
            # Get all mailboxes or filtered mailboxes (original logic)
            Write-Host "Retrieving all mailboxes matching criteria..." -ForegroundColor Yellow
            
            # Build parameters for Get-Mailbox
            $GetMailboxParams = @{
                ResultSize = 'Unlimited'
                RecipientTypeDetails = 'UserMailbox'
            }
            
            if ($MailboxFilter) {
                $GetMailboxParams.Filter = "EmailAddresses -like '*$MailboxFilter*'"
                Write-Host "Applying filter: $MailboxFilter" -ForegroundColor Cyan
            }
            
            # Test the cmdlet first with a small result set
            Write-Host "Testing mailbox retrieval..." -ForegroundColor Yellow
            $TestMailbox = Get-Mailbox -ResultSize 1 -ErrorAction Stop
            if (-not $TestMailbox) {
                throw "Unable to retrieve any mailboxes - check permissions"
            }
            
            # Now get all mailboxes
            $Mailboxes = Get-Mailbox @GetMailboxParams
            
            if ($Mailboxes.Count -eq 0) {
                Write-Warning "No mailboxes found matching the criteria"
                return
            }
            
            Write-Host "Found $($Mailboxes.Count) mailboxes to process" -ForegroundColor Green
        }
        
        # Sort mailboxes by display name
        $Mailboxes = $Mailboxes | Sort-Object DisplayName
        
    }
    catch {
        Write-Error "Failed to retrieve mailboxes: $($_.Exception.Message)"
        Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
        if ($InputCsvPath) {
            Write-Host "1. Verify CSV file exists and is readable" -ForegroundColor Cyan
            Write-Host "2. Check that email addresses in CSV are valid" -ForegroundColor Cyan
            Write-Host "3. Ensure CSV has correct column name: '$CsvEmailColumn'" -ForegroundColor Cyan
        }
        Write-Host "4. Verify your account has permission to read mailboxes" -ForegroundColor Cyan
        Write-Host "5. Check if you need 'View-Only Recipients' or higher role" -ForegroundColor Cyan
        Write-Host "6. Try running 'Get-Mailbox -ResultSize 1' manually" -ForegroundColor Cyan
        throw "Mailbox retrieval failed"
    }

    # Process mailboxes sequentially (more reliable with Exchange Online)
    Write-Host "`nProcessing mailbox permissions..." -ForegroundColor Yellow
    Write-Host "Note: Using sequential processing for better Exchange Online compatibility" -ForegroundColor Cyan
    
    $ProcessedCount = 0
    $ErrorCount = 0
    
    foreach ($Mailbox in $Mailboxes) {
        try {
            $CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
            
            # Get calendar folder permissions
            $Permissions = Get-MailboxFolderPermission -Identity $CalendarPath -ErrorAction Stop
            
            foreach ($Permission in $Permissions) {
                # Get user name safely - handle different possible structures
                $UserName = ""
                if ($Permission.User) {
                    if ($Permission.User.DisplayName) {
                        $UserName = $Permission.User.DisplayName
                    } elseif ($Permission.User.ToString()) {
                        $UserName = $Permission.User.ToString()
                    } else {
                        $UserName = $Permission.User
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
                    if ($Permission.AccessRights -is [array]) {
                        $AccessRights = $Permission.AccessRights -join '; '
                    } else {
                        $AccessRights = $Permission.AccessRights.ToString()
                    }
                }
                
                # Determine if this is a system permission
                $IsSystemPermission = $SystemPermissions -contains $UserName
                
                if ($IncludeSystemPermissions.IsPresent -or -not $IsSystemPermission) {
                    $PermissionResult = [PSCustomObject]@{
                        MailboxDisplayName = $Mailbox.DisplayName
                        MailboxPrimaryEmail = $Mailbox.PrimarySmtpAddress
                        MailboxAlias = $Mailbox.Alias
                        MailboxSamAccountName = $Mailbox.SamAccountName
                        DelegateUser = $UserName
                        DelegateUserType = $UserType
                        PermissionLevel = $AccessRights
                        IsSystemPermission = $IsSystemPermission
                        CalendarPath = $CalendarPath
                        ProcessedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    }
                    $Results += $PermissionResult
                }
            }
            
            $ProcessedCount++
            
            # Display progress
            if ($ProcessedCount -gt 0 -and ($ProcessedCount % $ProgressInterval -eq 0)) {
                $PercentComplete = [math]::Round(($ProcessedCount / $Mailboxes.Count) * 100, 1)
                Write-Host "Progress: $ProcessedCount/$($Mailboxes.Count) mailboxes processed ($PercentComplete%)" -ForegroundColor Cyan
            }
        }
        catch {
            Write-Warning "Error processing mailbox $($Mailbox.DisplayName): $($_.Exception.Message)"
            
            # Add error entry
            $ErrorResult = [PSCustomObject]@{
                MailboxDisplayName = $Mailbox.DisplayName
                MailboxPrimaryEmail = $Mailbox.PrimarySmtpAddress
                MailboxAlias = $Mailbox.Alias
                MailboxSamAccountName = $Mailbox.SamAccountName
                DelegateUser = "ERROR"
                DelegateUserType = "ERROR"
                PermissionLevel = $_.Exception.Message
                IsSystemPermission = $false
                CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
                ProcessedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            $Results += $ErrorResult
            $ErrorCount++
        }
    }

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
            try {
                $PermissionGroups = $DelegatedPermissions | Group-Object PermissionLevel -ErrorAction Stop | Sort-Object Count -Descending | Select-Object -First 5
                foreach ($Group in $PermissionGroups) {
                    Write-Host "  $($Group.Name): $($Group.Count) entries" -ForegroundColor Cyan
                }
            }
            catch {
                Write-Host "  Unable to generate permission level statistics: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "  Total delegated permissions: $($DelegatedPermissions.Count)" -ForegroundColor Cyan
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
    Write-Host "`nScript execution completed." -ForegroundColor Green
}

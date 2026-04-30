<#
.SYNOPSIS
    Audits M365 Groups for external user sharing and Entra ID Security Group assignments to SharePoint sites.

.DESCRIPTION
    This script performs two audit operations:

    Part 1 - M365 Groups External Sharing Audit:
        Retrieves all Microsoft 365 Groups that have an associated SharePoint site or Teams site,
        then identifies which groups have external (guest) users and lists those users per group.

    Part 2 - Entra ID Security Group SharePoint Assignments:
        Identifies Entra ID Security Groups that have been granted access to SharePoint sites
        by checking SharePoint site permissions for security group principals.

    Output is written to the console and optionally exported to CSV files.

.PARAMETER TenantId
    The Azure AD Tenant ID. Required for certificate-based authentication.

.PARAMETER ClientId
    The App Registration Client ID for non-interactive authentication.

.PARAMETER CertificateThumbprint
    The certificate thumbprint for app-only authentication.

.PARAMETER OutputPath
    Directory path for CSV export files. Defaults to C:\softdist\Logs\GroupAudit.

.PARAMETER SkipExternalUsers
    Skip Part 1 (external user enumeration). Useful if you only need security group assignments.

.PARAMETER SkipSecurityGroups
    Skip Part 2 (security group SharePoint assignments).

.PARAMETER ExportCsv
    Export results to CSV files in the OutputPath directory.

.PARAMETER MaxSites
    Limit the number of SharePoint sites processed in Part 2. Useful for large tenants during testing.
    Set to 0 for no limit (default).

.EXAMPLE
    .\Get-M365GroupExternalSharingAudit.ps1
    Runs interactively using delegated permissions (requires SharePoint Admin role).

.EXAMPLE
    .\Get-M365GroupExternalSharingAudit.ps1 -ExportCsv -OutputPath "C:\Reports\GroupAudit"
    Runs interactively and exports results to CSV files.

.EXAMPLE
    .\Get-M365GroupExternalSharingAudit.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "ABCDEF1234567890" -ExportCsv
    Runs using app-only authentication with a certificate.

.EXAMPLE
    .\Get-M365GroupExternalSharingAudit.ps1 -SkipSecurityGroups -ExportCsv
    Only audits M365 Groups for external users, skips the security group SharePoint check.

.NOTES
    Requirements:
        - Microsoft.Graph PowerShell SDK (Install-Module Microsoft.Graph)
        - PnP.PowerShell module for SharePoint site permission checks (Install-Module PnP.PowerShell)
        - For delegated auth: SharePoint Administrator + User.Read.All + Group.Read.All permissions
        - For app-only auth: App registration with the following Graph API permissions:
            * Group.Read.All
            * User.Read.All
            * Sites.FullControl.All (or Sites.Read.All minimum)
            * Directory.Read.All
        - PnP app registration also required for Part 2 if using app-only

    Limitations:
        - Part 2 (Security Group SharePoint check) can be slow in large tenants.
          Use -MaxSites to limit scope during testing.
        - External users are identified by the '#EXT#' UPN suffix convention.
        - SharePoint site permission enumeration via PnP requires a separate PnP connection.

    Author:  IT Infrastructure Team
    Version: 1.1.2
#>

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.Sites

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\softdist\Logs\GroupAudit",

    [Parameter(Mandatory = $false)]
    [switch]$SkipExternalUsers,

    [Parameter(Mandatory = $false)]
    [switch]$SkipSecurityGroups,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [int]$MaxSites = 0
)

#region --- Initialization & Logging ---

$ScriptVersion = "1.1.2"
$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile        = Join-Path $OutputPath "GroupAudit_$Timestamp.log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    $Entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $Entry -ErrorAction SilentlyContinue

    switch ($Level) {
        'INFO'    { Write-Host $Entry -ForegroundColor Cyan }
        'WARN'    { Write-Host $Entry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Entry -ForegroundColor Red }
        'SUCCESS' { Write-Host $Entry -ForegroundColor Green }
    }
}

function Initialize-OutputDirectory {
    if (-not (Test-Path $OutputPath)) {
        try {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            Write-Log "Created output directory: $OutputPath"
        }
        catch {
            Write-Warning "Could not create output directory '$OutputPath'. Logging to console only."
        }
    }
}

Initialize-OutputDirectory
Write-Log "=== M365 Group External Sharing & Security Group Audit v$ScriptVersion ===" -Level INFO
Write-Log "Output path: $OutputPath"

#endregion

#region --- Authentication ---

function Connect-GraphServices {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph using either app-only or delegated authentication.
    #>
    $Scopes = @(
        "Group.Read.All",
        "User.Read.All",
        "Sites.Read.All",
        "Directory.Read.All",
        "GroupMember.Read.All"
    )

    try {
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Write-Log "Connecting to Microsoft Graph using app-only authentication..."
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
        }
        else {
            Write-Log "Connecting to Microsoft Graph using delegated authentication..."
            Connect-MgGraph -Scopes $Scopes -NoWelcome
        }
        Write-Log "Successfully connected to Microsoft Graph." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Level ERROR
        throw
    }
}

#endregion

#region --- Part 1: M365 Groups with External Users ---

function Get-M365GroupsWithExternalUsers {
    <#
    .SYNOPSIS
        Retrieves all M365 Unified Groups with SharePoint/Teams sites and identifies external members.
    .OUTPUTS
        Array of PSCustomObjects with group and external user details.
    #>

    Write-Log "--- Part 1: Auditing M365 Groups for External Users ---"
    $Results = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Retrieve all M365 Unified Groups (ResourceProvisioningOptions indicates Teams)
    Write-Log "Retrieving all M365 Unified Groups..."
    try {
        $AllGroups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" `
            -Property "Id,DisplayName,Mail,Description,ResourceProvisioningOptions,Visibility,SharePointSiteUrl,CreatedDateTime" `
            -All -ConsistencyLevel eventual -ErrorAction Stop

        Write-Log "Retrieved $($AllGroups.Count) M365 Unified Groups."
    }
    catch {
        Write-Log "Failed to retrieve M365 Groups: $_" -Level ERROR
        return $Results
    }

    $GroupCounter = 0
    foreach ($Group in $AllGroups) {
        $GroupCounter++
        $HasTeams = $Group.ResourceProvisioningOptions -contains "Team"

        Write-Verbose "[$GroupCounter/$($AllGroups.Count)] Processing: $($Group.DisplayName)"

        # Use direct Graph API call instead of Get-MgGroupMember to ensure userType and
        # userPrincipalName are populated. The SDK's -Property parameter on Get-MgGroupMember
        # does not reliably populate AdditionalProperties for these fields.
        try {
            $Uri     = "https://graph.microsoft.com/v1.0/groups/$($Group.Id)/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999"
            $Members = @()

            do {
                $Response = Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop
                $Members += $Response.value
                $Uri      = $Response.'@odata.nextLink'
            } while ($Uri)
        }
        catch {
            Write-Log "Could not retrieve members for group '$($Group.DisplayName)': $_" -Level WARN
            continue
        }

        # Filter for external/guest users - identified by '#EXT#' in UPN or UserType = 'Guest'
        $ExternalMembers = $Members | Where-Object {
            $_.userPrincipalName -like "*#EXT#*" -or $_.userType -eq "Guest"
        }

        if ($ExternalMembers.Count -gt 0) {
            Write-Log "  Found $($ExternalMembers.Count) external user(s) in: $($Group.DisplayName)" -Level WARN

            foreach ($ExtUser in $ExternalMembers) {
                $Upn         = $ExtUser.userPrincipalName
                $DisplayName = $ExtUser.displayName
                $Mail        = $ExtUser.mail
                $UserType    = $ExtUser.userType

                # Derive the original external email from the B2B UPN encoding
                # Format: localpart_domain.com#EXT#@tenant.onmicrosoft.com
                # Last underscore before #EXT# is the @ separator
                $ExternalEmail = if ($Upn -like "*#EXT#*") {
                    $LocalPart = ($Upn -split "#EXT#")[0]
                    $AtIndex   = $LocalPart.LastIndexOf("_")
                    if ($AtIndex -ge 0) {
                        $LocalPart.Substring(0, $AtIndex) + "@" + $LocalPart.Substring($AtIndex + 1)
                    } else { $LocalPart }
                } else { $Mail }

                $Results.Add([PSCustomObject]@{
                    GroupId           = $Group.Id
                    GroupName         = $Group.DisplayName
                    GroupEmail        = $Group.Mail
                    GroupVisibility   = $Group.Visibility
                    HasTeamsSite      = $HasTeams
                    GroupCreated      = $Group.CreatedDateTime
                    ExternalUserId    = $ExtUser.id
                    ExternalUPN       = $Upn
                    ExternalEmail     = $ExternalEmail
                    ExternalDisplay   = $DisplayName
                    UserType          = $UserType
                })
            }
        }
    }

    Write-Log "Part 1 complete. Found $($Results.Count) external user memberships across $($AllGroups.Count) groups." -Level SUCCESS
    return $Results
}

#endregion

#region --- Part 2: Entra ID Security Groups on SharePoint Sites ---

function Get-SecurityGroupSharePointAssignments {
    <#
    .SYNOPSIS
        Identifies Entra ID Security Groups that have been granted access to SharePoint sites.
    .DESCRIPTION
        Enumerates SharePoint sites via Graph API and checks site permissions for security group principals.
        Only returns sites where security groups are explicitly listed as site members/owners/visitors.
    .OUTPUTS
        Array of PSCustomObjects with site and security group details.
    #>

    Write-Log "--- Part 2: Auditing Entra ID Security Group SharePoint Assignments ---"
    $Results = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Get all SharePoint sites via the search=* API.
    # IMPORTANT: Get-MgSite -All does NOT enumerate tenant sites without a search query
    # and will always return 0 results. The /sites?search=* endpoint is the correct approach.
    Write-Log "Retrieving SharePoint sites via Graph search API..."
    $Sites = [System.Collections.Generic.List[PSCustomObject]]::new()
    try {
        # The /sites?search= endpoint requires ConsistencyLevel: eventual header.
        # Without it, Graph returns HTTP 400 even with valid syntax and correct permissions.
        $Uri     = "https://graph.microsoft.com/v1.0/sites?search=%2A&`$select=id,displayName,webUrl,name&`$top=200"
        $Headers = @{ "ConsistencyLevel" = "eventual" }

        do {
            $Response = Invoke-MgGraphRequest -Uri $Uri -Method GET -Headers $Headers -ErrorAction Stop
            foreach ($Site in $Response.value) {
                if ($Site.webUrl -notlike "*/personal/*") {
                    $Sites.Add([PSCustomObject]@{
                        Id          = $Site.id
                        DisplayName = $Site.displayName
                        WebUrl      = $Site.webUrl
                        Name        = $Site.name
                    })
                }
            }
            $Uri = $Response.'@odata.nextLink'
        } while ($Uri)

        if ($MaxSites -gt 0) {
            Write-Log "MaxSites limit applied: processing first $MaxSites of $($Sites.Count) sites." -Level WARN
            $Sites = [System.Collections.Generic.List[PSCustomObject]]($Sites | Select-Object -First $MaxSites)
        }

        Write-Log "Processing $($Sites.Count) SharePoint sites."
    }
    catch {
        Write-Log "Failed to retrieve SharePoint sites: $_" -Level ERROR
        return $Results
    }

    # Pre-fetch all Entra ID Security Groups for lookup efficiency
    Write-Log "Pre-fetching Entra ID Security Groups for cross-reference..."
    try {
        $SecurityGroups = Get-MgGroup -Filter "securityEnabled eq true and NOT(groupTypes/any(c:c eq 'Unified'))" `
            -Property "Id,DisplayName,Mail,Description,OnPremisesSyncEnabled" `
            -All -ConsistencyLevel eventual -ErrorAction Stop

        # Build a lookup hashtable by Id for O(1) access
        $SecurityGroupLookup = @{}
        foreach ($SG in $SecurityGroups) {
            $SecurityGroupLookup[$SG.Id] = $SG
        }
        Write-Log "Found $($SecurityGroups.Count) Entra ID Security Groups."
    }
    catch {
        Write-Log "Failed to retrieve Security Groups: $_" -Level ERROR
        return $Results
    }

    $SiteCounter = 0
    foreach ($Site in $Sites) {
        $SiteCounter++
        Write-Verbose "[$SiteCounter/$($Sites.Count)] Checking permissions: $($Site.DisplayName)"

        # Retrieve site permissions via Graph
        try {
            $Permissions = Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction Stop
        }
        catch {
            Write-Log "Could not retrieve permissions for site '$($Site.DisplayName)': $_" -Level WARN
            continue
        }

        foreach ($Permission in $Permissions) {
            # Graph site permissions expose grantedToIdentitiesV2 and grantedToV2
            $GrantedIdentities = @()

            if ($Permission.GrantedToIdentitiesV2) {
                $GrantedIdentities += $Permission.GrantedToIdentitiesV2
            }
            if ($Permission.GrantedToV2) {
                $GrantedIdentities += $Permission.GrantedToV2
            }

            foreach ($Identity in $GrantedIdentities) {
                $GroupIdentity = $Identity.Group
                if (-not $GroupIdentity) { continue }

                $GroupId = $GroupIdentity.Id

                # Check if this group is a security group (not M365 group)
                if ($SecurityGroupLookup.ContainsKey($GroupId)) {
                    $SecGroup = $SecurityGroupLookup[$GroupId]

                    Write-Log "  Security group '$($SecGroup.DisplayName)' has access to: $($Site.DisplayName)" -Level WARN

                    $Results.Add([PSCustomObject]@{
                        SiteId              = $Site.Id
                        SiteName            = $Site.DisplayName
                        SiteUrl             = $Site.WebUrl
                        PermissionId        = $Permission.Id
                        PermissionRoles     = ($Permission.Roles -join ", ")
                        SecurityGroupId     = $SecGroup.Id
                        SecurityGroupName   = $SecGroup.DisplayName
                        SecurityGroupMail   = $SecGroup.Mail
                        SecurityGroupDesc   = $SecGroup.Description
                        IsSynced            = $SecGroup.OnPremisesSyncEnabled -eq $true
                    })
                }
            }
        }
    }

    Write-Log "Part 2 complete. Found $($Results.Count) security group SharePoint permission assignments." -Level SUCCESS
    return $Results
}

#endregion

#region --- Reporting ---

function Show-ExternalUserSummary {
    param([System.Collections.Generic.List[PSCustomObject]]$Data)

    if ($Data.Count -eq 0) {
        Write-Log "No external users found in any M365 Groups." -Level SUCCESS
        return
    }

    Write-Host "`n===== PART 1: M365 GROUPS WITH EXTERNAL USERS =====" -ForegroundColor Magenta

    $GroupedData = $Data | Group-Object -Property GroupName | Sort-Object Name

    foreach ($GroupEntry in $GroupedData) {
        Write-Host "`n  Group: $($GroupEntry.Name)" -ForegroundColor Yellow
        Write-Host "  ID:    $($GroupEntry.Group[0].GroupId)" -ForegroundColor DarkGray
        Write-Host "  Has Teams: $($GroupEntry.Group[0].HasTeamsSite)" -ForegroundColor DarkGray
        Write-Host "  External Users ($($GroupEntry.Count)):" -ForegroundColor Cyan

        foreach ($User in $GroupEntry.Group) {
            Write-Host "    - $($User.ExternalDisplay) | $($User.ExternalEmail) | Type: $($User.UserType)" -ForegroundColor White
        }
    }

    Write-Host "`n  SUMMARY: $($GroupedData.Count) groups have external users | $($Data.Count) total external memberships`n" -ForegroundColor Green
}

function Show-SecurityGroupSummary {
    param([System.Collections.Generic.List[PSCustomObject]]$Data)

    if ($Data.Count -eq 0) {
        Write-Log "No Entra ID Security Groups found with direct SharePoint site permissions." -Level SUCCESS
        return
    }

    Write-Host "`n===== PART 2: SECURITY GROUPS WITH SHAREPOINT ACCESS =====" -ForegroundColor Magenta

    $GroupedBySite = $Data | Group-Object -Property SiteName | Sort-Object Name

    foreach ($SiteEntry in $GroupedBySite) {
        Write-Host "`n  Site: $($SiteEntry.Name)" -ForegroundColor Yellow
        Write-Host "  URL:  $($SiteEntry.Group[0].SiteUrl)" -ForegroundColor DarkGray
        Write-Host "  Security Groups ($($SiteEntry.Count)):" -ForegroundColor Cyan

        foreach ($SGEntry in $SiteEntry.Group) {
            $SyncedLabel = if ($SGEntry.IsSynced) { " [On-Prem Synced]" } else { "" }
            Write-Host "    - $($SGEntry.SecurityGroupName)$SyncedLabel | Roles: $($SGEntry.PermissionRoles)" -ForegroundColor White
        }
    }

    Write-Host "`n  SUMMARY: $($GroupedBySite.Count) sites have security group assignments | $($Data.Count) total assignments`n" -ForegroundColor Green
}

function Export-AuditResults {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$ExternalUserData,
        [System.Collections.Generic.List[PSCustomObject]]$SecurityGroupData
    )

    if ($ExternalUserData -and $ExternalUserData.Count -gt 0) {
        $ExternalCsvPath = Join-Path $OutputPath "M365Groups_ExternalUsers_$Timestamp.csv"
        $ExternalUserData | Export-Csv -Path $ExternalCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log "External users report exported to: $ExternalCsvPath" -Level SUCCESS
    }

    if ($SecurityGroupData -and $SecurityGroupData.Count -gt 0) {
        $SecurityGroupCsvPath = Join-Path $OutputPath "SecurityGroups_SharePoint_$Timestamp.csv"
        $SecurityGroupData | Export-Csv -Path $SecurityGroupCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log "Security group SharePoint assignments exported to: $SecurityGroupCsvPath" -Level SUCCESS
    }
}

#endregion

#region --- Main Execution ---

try {
    # Connect to Graph
    Connect-GraphServices

    $ExternalUserResults     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $SecurityGroupResults    = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Part 1: External Users in M365 Groups
    if (-not $SkipExternalUsers) {
        $ExternalUserResults = Get-M365GroupsWithExternalUsers
        Show-ExternalUserSummary -Data $ExternalUserResults
    }
    else {
        Write-Log "Skipping Part 1 (SkipExternalUsers flag set)." -Level WARN
    }

    # Part 2: Security Groups on SharePoint Sites
    if (-not $SkipSecurityGroups) {
        $SecurityGroupResults = Get-SecurityGroupSharePointAssignments
        Show-SecurityGroupSummary -Data $SecurityGroupResults
    }
    else {
        Write-Log "Skipping Part 2 (SkipSecurityGroups flag set)." -Level WARN
    }

    # Export to CSV if requested
    if ($ExportCsv) {
        Export-AuditResults -ExternalUserData $ExternalUserResults -SecurityGroupData $SecurityGroupResults
    }

    Write-Log "=== Audit Complete ===" -Level SUCCESS
    Write-Log "Log file: $LogFile"
}
catch {
    Write-Log "Script encountered a fatal error: $_" -Level ERROR
    throw
}
finally {
    # Disconnect Graph session
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Disconnected from Microsoft Graph."
    }
    catch { }
}

#endregion

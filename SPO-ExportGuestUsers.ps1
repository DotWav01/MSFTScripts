<#
.SYNOPSIS
    Audits external users across all SharePoint Online sites and their permission levels.

.DESCRIPTION
    Connects to SharePoint Online and Microsoft Graph to enumerate all SPO sites,
    identifies external users (guests), and determines their permission levels,
    SharePoint group memberships, and Entra ID group memberships. Results are
    exported to a CSV file.

    Permission resolution order:
    1. Direct SharePoint group membership (Owners/Members/Visitors/custom groups)
    2. Entra ID group membership (security groups / M365 groups synced to SPO)
    3. Direct user permissions

.PARAMETER TenantName
    The SharePoint tenant name (e.g., 'contoso' for contoso.sharepoint.com).

.PARAMETER ClientId
    The Client ID (Application ID) of the Entra ID App Registration configured for PnP PowerShell.
    The app registration must have the required SharePoint and Graph delegated permissions.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to C:\softdist\Logs\SPO_ExternalUsers_<timestamp>.csv

.PARAMETER IncludeOneDrive
    If specified, includes OneDrive for Business sites in the audit.

.PARAMETER ThrottleDelayMs
    Delay in milliseconds between API calls to avoid throttling. Default: 200.
    Pass an integer value only (e.g. -ThrottleDelayMs 500), not a string like '500ms'.

.PARAMETER SiteFilter
    Optional URL filter to scope the audit to specific sites (wildcard supported).

.EXAMPLE
    .\Get-SPOExternalUserPermissions.ps1 -TenantName "contoso" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    Audits all SharePoint sites in the contoso tenant and exports results to the default log path.

.EXAMPLE
    .\Get-SPOExternalUserPermissions.ps1 -TenantName "contoso" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -OutputPath "C:\Reports\ExternalUsers.csv"
    Audits all sites and saves results to a custom path.

.EXAMPLE
    .\Get-SPOExternalUserPermissions.ps1 -TenantName "contoso" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -SiteFilter "*project*" -Verbose
    Audits only sites with 'project' in the URL, with verbose logging.

.EXAMPLE
    .\Get-SPOExternalUserPermissions.ps1 -TenantName "contoso" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -IncludeOneDrive -ThrottleDelayMs 500
    Includes OneDrive for Business sites and increases throttle delay to 500ms.

.NOTES
    Requirements:
    - PnP.PowerShell module        (Install-Module PnP.PowerShell)
    - Microsoft.Graph module       (Install-Module Microsoft.Graph)
    - SharePoint Online Administrator role
    - Entra App Registration with:
        * Delegated: SharePoint > AllSites.FullControl or Sites.Read.All
        * Delegated: Microsoft Graph > User.Read.All
        * Delegated: Microsoft Graph > GroupMember.Read.All
        * Delegated: Microsoft Graph > Directory.Read.All
        * Authentication > Mobile/desktop redirect URI:
          https://login.microsoftonline.com/common/oauth2/nativeclient
        * Authentication > Allow public client flows: Yes

    External users are identified by the presence of '#EXT#' in their UPN or
    by UserType = 'Guest' in Entra ID.

    Large tenants: Use -SiteFilter to scope the audit or expect long runtimes.
    Throttling: The script includes configurable delays and retry logic.

    Author: IT Infrastructure
    Version: 1.1.0
#>

#Requires -Modules PnP.PowerShell, Microsoft.Graph.Users, Microsoft.Graph.Groups

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$TenantName,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$IncludeOneDrive,

    [Parameter()]
    [ValidateRange(0, 60000)]
    [int]$ThrottleDelayMs = 200,

    [Parameter()]
    [string]$SiteFilter
)

#region --- Initialization ---

$ErrorActionPreference = 'Continue'
$script:Results         = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:GraphUserCache  = @{}   # UPN -> Entra user object
$script:GraphGroupCache = @{}   # UserId -> array of group display names

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

if (-not $OutputPath) {
    $LogDir = 'C:\softdist\Logs'
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    $OutputPath = Join-Path $LogDir "SPO_ExternalUsers_$Timestamp.csv"
}

$LogFile = [System.IO.Path]::ChangeExtension($OutputPath, '.log')

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $Entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogFile -Value $Entry
    switch ($Level) {
        'ERROR'   { Write-Error   $Message }
        'WARNING' { Write-Warning $Message }
        'VERBOSE' { Write-Verbose $Message }
        default   { Write-Host    $Entry }
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$BaseDelayMs = 1000
    )
    $attempt = 0
    while ($attempt -le $MaxRetries) {
        try {
            return & $ScriptBlock
        } catch {
            $attempt++
            if ($attempt -gt $MaxRetries) { throw }
            $delay = $BaseDelayMs * [Math]::Pow(2, $attempt - 1)
            Write-Log "Throttle/error on attempt $attempt. Retrying in ${delay}ms. Error: $_" -Level WARNING
            Start-Sleep -Milliseconds $delay
        }
    }
}

#endregion

#region --- Authentication ---

function Connect-Services {
    Write-Log "Connecting to SharePoint Online admin..."
    try {
        Connect-PnPOnline -Url "https://$TenantName-admin.sharepoint.com" `
                          -ClientId $ClientId `
                          -Interactive
        Write-Log "Connected to SPO admin."
    } catch {
        Write-Log "Failed to connect to SPO admin: $_" -Level ERROR
        throw
    }

    Write-Log "Connecting to Microsoft Graph..."
    try {
        Connect-MgGraph -Scopes "User.Read.All", "GroupMember.Read.All", "Directory.Read.All" `
                        -NoWelcome
        Write-Log "Connected to Microsoft Graph."
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Level ERROR
        throw
    }
}

#endregion

#region --- Graph Helpers ---

function Get-EntraUser {
    param([string]$UPN)
    if ($script:GraphUserCache.ContainsKey($UPN)) {
        return $script:GraphUserCache[$UPN]
    }
    try {
        # External users have #EXT# in UPN — Graph accepts the raw UPN for lookup
        $user = Invoke-WithRetry -ScriptBlock {
            Get-MgUser -UserId $UPN -Property Id, DisplayName, UserPrincipalName, UserType, Mail -ErrorAction Stop
        }
        $script:GraphUserCache[$UPN] = $user
        return $user
    } catch {
        Write-Log "Could not retrieve Entra user for UPN '$UPN': $_" -Level VERBOSE
        $script:GraphUserCache[$UPN] = $null
        return $null
    }
}

function Get-UserEntraGroups {
    param([string]$UserId)
    if ($script:GraphGroupCache.ContainsKey($UserId)) {
        return $script:GraphGroupCache[$UserId]
    }
    try {
        $groups = Invoke-WithRetry -ScriptBlock {
            Get-MgUserMemberOf -UserId $UserId -All -ErrorAction Stop
        }
        $groupNames = $groups |
            Where-Object { $_.AdditionalProperties['@odata.type'] -in '#microsoft.graph.group', '#microsoft.graph.directoryRole' } |
            ForEach-Object { $_.AdditionalProperties['displayName'] }
        $script:GraphGroupCache[$UserId] = $groupNames
        return $groupNames
    } catch {
        Write-Log "Could not retrieve group memberships for user '$UserId': $_" -Level WARNING
        $script:GraphGroupCache[$UserId] = @()
        return @()
    }
}

#endregion

#region --- External User Detection ---

function Test-IsExternalUser {
    param([string]$LoginName, [string]$Email)
    # SharePoint represents external users with i:0#.f|membership|..._ext_ or #EXT#
    return (
        $LoginName -match '#EXT#' -or
        $LoginName -match '_ext_' -or
        $LoginName -match 'urn:spo:guest' -or
        ($Email -and $Email -notmatch [regex]::Escape("@$TenantName") -and $Email -match '@')
    )
}

#endregion

#region --- SharePoint Permission Enumeration ---

function Get-SiteExternalUsers {
    param([string]$SiteUrl)

    Write-Log "Processing site: $SiteUrl" -Level VERBOSE

    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
    } catch {
        Write-Log "Failed to connect to site '$SiteUrl': $_" -Level ERROR
        return
    }

    # --- Enumerate SharePoint Groups and their members ---
    $spoGroups = @{}  # LoginName -> [groupName, permissionLevel]
    try {
        $groups = Invoke-WithRetry -ScriptBlock { Get-PnPGroup -ErrorAction Stop }
        foreach ($group in $groups) {
            try {
                Start-Sleep -Milliseconds $ThrottleDelayMs

                # Get the permission level(s) assigned to this group on the root web
                $roleAssignments = Invoke-WithRetry -ScriptBlock {
                    Get-PnPGroupPermissions -Identity $group.Id -ErrorAction Stop
                }
                $permLevels = ($roleAssignments | ForEach-Object { $_.Name }) -join '; '

                $members = Invoke-WithRetry -ScriptBlock {
                    Get-PnPGroupMember -Group $group.Id -ErrorAction Stop
                }
                foreach ($member in $members) {
                    if (Test-IsExternalUser -LoginName $member.LoginName -Email $member.Email) {
                        if (-not $spoGroups.ContainsKey($member.LoginName)) {
                            $spoGroups[$member.LoginName] = [System.Collections.Generic.List[PSCustomObject]]::new()
                        }
                        $spoGroups[$member.LoginName].Add([PSCustomObject]@{
                            GroupName       = $group.Title
                            PermissionLevel = if ($permLevels) { $permLevels } else { 'N/A (Entra Group)' }
                            Email           = $member.Email
                            DisplayName     = $member.Title
                        })
                    }
                }
            } catch {
                Write-Log "Error processing group '$($group.Title)' on '$SiteUrl': $_" -Level WARNING
            }
        }
    } catch {
        Write-Log "Error enumerating groups on '$SiteUrl': $_" -Level ERROR
    }

    # --- Enumerate direct role assignments (users assigned directly, not via SP group) ---
    $directUsers = @{}
    try {
        $web = Invoke-WithRetry -ScriptBlock { Get-PnPWeb -Includes RoleAssignments -ErrorAction Stop }
        foreach ($ra in $web.RoleAssignments) {
            try {
                $ra.EnsureProperties([Microsoft.SharePoint.Client.RoleAssignment], 'Member', 'RoleDefinitionBindings')
                $member = $ra.Member
                $roles  = ($ra.RoleDefinitionBindings | ForEach-Object {
                    $_.EnsureProperties([Microsoft.SharePoint.Client.RoleDefinition], 'Name')
                    $_.Name
                }) -join '; '

                # Could be a user or an Entra group assigned directly
                if ($member.PrincipalType -eq 'User') {
                    if (Test-IsExternalUser -LoginName $member.LoginName -Email $member.Email) {
                        if (-not $directUsers.ContainsKey($member.LoginName)) {
                            $directUsers[$member.LoginName] = [PSCustomObject]@{
                                LoginName       = $member.LoginName
                                Email           = $member.Email
                                DisplayName     = $member.Title
                                PermissionLevel = $roles
                                GrantedVia      = 'Direct Assignment'
                            }
                        }
                    }
                }
            } catch {
                Write-Log "Error processing role assignment on '$SiteUrl': $_" -Level VERBOSE
            }
        }
    } catch {
        Write-Log "Could not enumerate direct role assignments on '$SiteUrl': $_" -Level WARNING
    }

    # --- Combine results and resolve Entra group memberships ---
    $allExternalLogins = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($login in $spoGroups.Keys)   { [void]$allExternalLogins.Add($login) }
    foreach ($login in $directUsers.Keys) { [void]$allExternalLogins.Add($login) }

    foreach ($loginName in $allExternalLogins) {
        $email       = $null
        $displayName = $null

        # Prefer data from SPO group membership
        if ($spoGroups.ContainsKey($loginName) -and $spoGroups[$loginName].Count -gt 0) {
            $email       = $spoGroups[$loginName][0].Email
            $displayName = $spoGroups[$loginName][0].DisplayName
        } elseif ($directUsers.ContainsKey($loginName)) {
            $email       = $directUsers[$loginName].Email
            $displayName = $directUsers[$loginName].DisplayName
        }

        # Resolve SPO group memberships for this user
        $spoGroupList = if ($spoGroups.ContainsKey($loginName)) {
            ($spoGroups[$loginName] | ForEach-Object { "$($_.GroupName) [$($_.PermissionLevel)]" }) -join ' | '
        } else { '' }

        $spoPermLevel = if ($spoGroups.ContainsKey($loginName)) {
            ($spoGroups[$loginName] | ForEach-Object { $_.PermissionLevel } | Select-Object -Unique) -join '; '
        } elseif ($directUsers.ContainsKey($loginName)) {
            $directUsers[$loginName].PermissionLevel
        } else { 'Unknown' }

        $grantedVia = if ($spoGroups.ContainsKey($loginName) -and $directUsers.ContainsKey($loginName)) {
            'SharePoint Group + Direct'
        } elseif ($spoGroups.ContainsKey($loginName)) { 'SharePoint Group' }
        else { 'Direct Assignment' }

        # Entra ID group lookup
        $entraGroupList = 'N/A'
        $entraUser      = $null
        if ($email) {
            # Strip claim provider prefix for Graph lookup
            $upnForGraph = $loginName -replace '^i:0#\.f\|membership\|', ''
            $entraUser = Get-EntraUser -UPN $upnForGraph

            if (-not $entraUser -and $email) {
                # Fallback: search by email
                try {
                    $entraUser = Invoke-WithRetry -ScriptBlock {
                        Get-MgUser -Filter "mail eq '$email' or userPrincipalName eq '$email'" `
                                   -Property Id, DisplayName, UserPrincipalName, UserType, Mail `
                                   -ErrorAction Stop | Select-Object -First 1
                    }
                } catch { }
            }

            if ($entraUser) {
                $entraGroups    = Get-UserEntraGroups -UserId $entraUser.Id
                $entraGroupList = if ($entraGroups) { $entraGroups -join ' | ' } else { 'None' }
                if (-not $displayName) { $displayName = $entraUser.DisplayName }
            }
        }

        $script:Results.Add([PSCustomObject]@{
            SiteUrl         = $SiteUrl
            UserDisplayName = $displayName
            UserEmail       = $email
            LoginName       = $loginName
            UserType        = 'External/Guest'
            PermissionLevel = $spoPermLevel
            GrantedVia      = $grantedVia
            SPO_Groups      = $spoGroupList
            EntraID_Groups  = $entraGroupList
            EntraUserFound  = ($null -ne $entraUser)
        })

        Start-Sleep -Milliseconds $ThrottleDelayMs
    }

    Write-Log "Site '$SiteUrl' complete. Found $($allExternalLogins.Count) external user(s)." -Level VERBOSE
}

#endregion

#region --- Main Execution ---

try {
    Write-Log "=== SPO External User Audit Started ==="
    Write-Log "Tenant: $TenantName | Output: $OutputPath"

    Connect-Services

    # Get all sites
    Write-Log "Retrieving all SharePoint sites..."
    $allSites = Invoke-WithRetry -ScriptBlock {
        Get-PnPTenantSite -ErrorAction Stop
    }

    # Exclude OneDrive unless requested
    if (-not $IncludeOneDrive) {
        $allSites = $allSites | Where-Object { $_.Url -notlike '*-my.sharepoint.com/personal/*' }
    }

    if ($SiteFilter) {
        $allSites = $allSites | Where-Object { $_.Url -like $SiteFilter }
        Write-Log "Filtered to $($allSites.Count) site(s) matching '$SiteFilter'."
    }

    Write-Log "Found $($allSites.Count) site(s) to audit."

    $siteIndex = 0
    foreach ($site in $allSites) {
        $siteIndex++
        $pct = [Math]::Round(($siteIndex / $allSites.Count) * 100, 1)
        Write-Progress -Activity "Auditing SharePoint Sites" `
                       -Status "[$siteIndex/$($allSites.Count)] $($site.Url)" `
                       -PercentComplete $pct

        if ($PSCmdlet.ShouldProcess($site.Url, "Audit external users")) {
            Get-SiteExternalUsers -SiteUrl $site.Url
        }

        Start-Sleep -Milliseconds $ThrottleDelayMs
    }

    Write-Progress -Activity "Auditing SharePoint Sites" -Completed

    # Export results
    if ($script:Results.Count -gt 0) {
        $script:Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Log "Exported $($script:Results.Count) record(s) to '$OutputPath'."
        Write-Host "`n[SUCCESS] Audit complete. Results: $OutputPath" -ForegroundColor Green
    } else {
        Write-Log "No external users found across $($allSites.Count) site(s)." -Level WARNING
        Write-Host "[INFO] No external users found." -ForegroundColor Yellow
    }

} catch {
    Write-Log "Fatal error: $_" -Level ERROR
    throw
} finally {
    Write-Log "=== Audit Complete. Total records: $($script:Results.Count) ==="
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}

#endregion

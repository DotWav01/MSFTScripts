#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.Governance

<#
.SYNOPSIS
    Creates and configures Entra ID App Registrations for Microsoft 365 Desired State Configuration (M365DSC).

.DESCRIPTION
    This script automates the creation of the six App Registrations required for M365DSC operations,
    assigning the correct Microsoft Graph and service-specific API permissions to each.

    The six App Registrations created are:
        1. M365DSC - Entra ID
        2. M365DSC - Microsoft Teams
        3. M365DSC - Security, Compliance & Exchange
        4. M365DSC - SharePoint & OneDrive
        5. M365DSC - Intune & Defender
        6. M365DSC - O365 & Other Services

    By default, the script ONLY creates the App Registrations and assigns API permissions.
    No credentials (certificates or secrets) are configured unless explicitly requested.

    Credential options (all optional):
        -CreateSelfSignedCert   Generate a self-signed certificate and upload it to each app.
        -ImportCert             Upload an existing .cer / .pem file to each app.
        -CreateClientSecret     Generate a client secret on each app.

    Exchange role assignment (optional):
        -AssignExchangeRole     Assign an Entra directory role to the two apps that require
                                Exchange.ManageAsApp: M365DSC - Security, Compliance & Exchange
                                and M365DSC - O365 & Other Services.
        -ExchangeRole           The role to assign. Options: GlobalReader, ExchangeAdministrator.
                                Default: GlobalReader. The role is assigned as a direct active
                                assignment (not PIM-eligible).

    All permissions are Application type and require admin consent, which this script can grant
    automatically when run with sufficient privileges (Global Administrator or Privileged Role Administrator).

.PARAMETER TenantId
    The Entra ID Tenant ID (GUID). Required for connecting to Microsoft Graph.

.PARAMETER AppNamePrefix
    Prefix applied to all created App Registration display names.
    Default: "M365DSC"

.PARAMETER CreateSelfSignedCert
    Switch. When specified, a self-signed certificate is generated and uploaded to each App Registration.
    Use -CertificateExpiryYears to control the validity period (default: 1 year).
    PFX files are exported to -CertificateOutputPath.

.PARAMETER CertificateExpiryYears
    Validity period in years for generated self-signed certificates.
    Default: 1  |  Range: 1-10

.PARAMETER CertificateOutputPath
    Directory path where generated self-signed PFX files are exported.
    Default: .\M365DSC-Certs

.PARAMETER CertificatePassword
    SecureString password used to protect exported PFX files.
    If not supplied when -CreateSelfSignedCert is used, you will be prompted interactively.

.PARAMETER ImportCert
    Switch. When specified, an existing certificate file is uploaded to each App Registration.
    Requires -CertificatePath to be set.

.PARAMETER CertificatePath
    Path to an existing certificate (.cer or .pem) to upload when -ImportCert is used.

.PARAMETER CreateClientSecret
    Switch. When specified, a client secret is created on each App Registration.
    Use -SecretExpiryYears to control the validity period (default: 1 year).

.PARAMETER SecretExpiryYears
    Validity period in years for generated client secrets.
    Default: 1  |  Range: 1-2

.PARAMETER AssignExchangeRole
    Switch. When specified, assigns a directory role to the Service Principals of the two apps
    that require Exchange.ManageAsApp:
        - M365DSC - Security, Compliance & Exchange
        - M365DSC - O365 & Other Services
    The role is assigned as a direct active assignment (not PIM-eligible).
    Use -ExchangeRole to choose which role to assign.

.PARAMETER ExchangeRole
    The Entra directory role to assign when -AssignExchangeRole is used.
    Options: GlobalReader, ExchangeAdministrator
    Default: GlobalReader

.PARAMETER GrantAdminConsent
    Automatically grant tenant-wide admin consent for all configured API permissions.
    Requires Global Administrator or Privileged Role Administrator.
    Default: $true

.PARAMETER LogPath
    Full path to the log file.
    Default: .\Logs\New-M365DSCAppRegistrations_<timestamp>.log

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

    Creates all six App Registrations with API permissions only. No credentials are configured.

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CreateSelfSignedCert

    Creates all six App Registrations and generates a self-signed certificate (1-year default) for each.

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CreateSelfSignedCert -CertificateExpiryYears 3 -CertificateOutputPath "C:\Certs\M365DSC"

    Creates all six App Registrations and generates a 3-year self-signed certificate for each (overriding the 1-year default),
    exporting PFX files to C:\Certs\M365DSC.

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ImportCert -CertificatePath "C:\Certs\M365DSC.cer"

    Creates all six App Registrations and uploads an existing certificate to each.

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CreateClientSecret -SecretExpiryYears 1

    Creates all six App Registrations with a client secret (1-year expiry) on each.

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CreateSelfSignedCert -CertificateExpiryYears 2 -CreateClientSecret -SecretExpiryYears 1

    Creates all six App Registrations with both a self-signed certificate (1 year) and a client secret (1 year).

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -AssignExchangeRole

    Creates all six App Registrations and assigns the Global Reader role (default) to the two
    Exchange-dependent apps as a direct active assignment.

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -AssignExchangeRole -ExchangeRole ExchangeAdministrator

    Creates all six App Registrations and assigns the Exchange Administrator role to the two
    Exchange-dependent apps as a direct active assignment.

.EXAMPLE
    .\New-M365DSCAppRegistrations.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -WhatIf

    Simulates the full run without making any changes to Entra ID.

.NOTES
    The running user/principal requires:
        Application.ReadWrite.All
        AppRoleAssignment.ReadWrite.All        (for admin consent)
        DelegatedPermissionGrant.ReadWrite.All (for admin consent)
        RoleManagement.ReadWrite.Directory     (for -AssignExchangeRole)

    Apps using Exchange.ManageAsApp (Security/Compliance/Exchange and O365 Services) require
    the Exchange Administrator or Global Reader Entra role on their Service Principal.
    Use -AssignExchangeRole to have the script assign this automatically.

    Version:    1.2.0
    Author:     IT Infrastructure
    Updated:    2026-03-27
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, HelpMessage = 'Entra ID Tenant ID (GUID)')]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$TenantId,

    [Parameter(HelpMessage = 'Prefix for App Registration display names')]
    [ValidateNotNullOrEmpty()]
    [string]$AppNamePrefix = 'M365DSC',

    # ── Certificate: Self-Signed ─────────────────────────────────────────
    [Parameter(HelpMessage = 'Generate and upload a self-signed certificate to each app')]
    [switch]$CreateSelfSignedCert,

    [Parameter(HelpMessage = 'Validity period in years for self-signed certificates (1-10)')]
    [ValidateRange(1, 10)]
    [int]$CertificateExpiryYears = 1,

    [Parameter(HelpMessage = 'Output directory for exported self-signed PFX files')]
    [string]$CertificateOutputPath = '.\M365DSC-Certs',

    [Parameter(HelpMessage = 'Password for exported PFX files (SecureString)')]
    [System.Security.SecureString]$CertificatePassword,

    # ── Certificate: Import Existing ─────────────────────────────────────
    [Parameter(HelpMessage = 'Upload an existing certificate (.cer/.pem) to each app')]
    [switch]$ImportCert,

    [Parameter(HelpMessage = 'Path to an existing .cer or .pem certificate to upload')]
    [string]$CertificatePath,

    # ── Client Secret ─────────────────────────────────────────────────────
    [Parameter(HelpMessage = 'Create a client secret on each app')]
    [switch]$CreateClientSecret,

    [Parameter(HelpMessage = 'Validity period in years for client secrets (1-2)')]
    [ValidateRange(1, 2)]
    [int]$SecretExpiryYears = 1,

    # ── General ───────────────────────────────────────────────────────────
    [Parameter(HelpMessage = 'Automatically grant admin consent for all API permissions')]
    [bool]$GrantAdminConsent = $true,

    # ── Exchange Role Assignment ──────────────────────────────────────────
    [Parameter(HelpMessage = 'Assign an Entra directory role to Exchange-dependent app Service Principals')]
    [switch]$AssignExchangeRole,

    [Parameter(HelpMessage = 'Role to assign: GlobalReader or ExchangeAdministrator')]
    [ValidateSet('GlobalReader', 'ExchangeAdministrator')]
    [string]$ExchangeRole = 'GlobalReader',

    [Parameter(HelpMessage = 'Full path to log file')]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Parameter Validation ────────────────────────────────────────────────

if ($ImportCert -and -not $CertificatePath) {
    throw '-CertificatePath is required when -ImportCert is specified.'
}

if ($ImportCert -and $CertificatePath -and -not (Test-Path $CertificatePath)) {
    throw "Certificate file not found: $CertificatePath"
}

if ($CreateSelfSignedCert -and $ImportCert) {
    throw 'Specify either -CreateSelfSignedCert or -ImportCert, not both.'
}

#endregion

#region ── Logging ─────────────────────────────────────────────────────────────

function Initialize-Log {
    [CmdletBinding()]
    param ([string]$Path)

    if (-not $Path) {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $logDir    = Join-Path $PSScriptRoot 'Logs'
        $Path      = Join-Path $logDir "New-M365DSCAppRegistrations_$timestamp.log"
    }

    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $script:LogFile = $Path
    Write-Log -Message "Log initialised: $Path" -Level INFO
}

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"

    # Only write non-empty lines to file (prevents Add-Content binding errors)
    if ($script:LogFile -and $Message.Length -gt 0) {
        Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
    }

    # Blank messages print as empty lines for visual spacing
    if ($Message.Length -eq 0) {
        Write-Host ''
        return
    }

    $colour = switch ($Level) {
        'INFO'    { 'Cyan'   }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red'    }
        'SUCCESS' { 'Green'  }
        'DEBUG'   { 'Gray'   }
        default   { 'White'  }
    }
    Write-Host $line -ForegroundColor $colour
}

#endregion

#region ── App Registration Definitions ────────────────────────────────────────

function Get-AppDefinitions {
    <#
    .SYNOPSIS
        Returns all App Registration definitions including required API permissions.
    .NOTES
        Microsoft Graph App ID : 00000003-0000-0000-c000-000000000000
        Exchange Online App ID : 00000002-0000-0ff1-ce00-000000000000
        SharePoint App ID      : 00000003-0000-0ff1-ce00-000000000000
    #>

    $graphAppId    = '00000003-0000-0000-c000-000000000000'  # Microsoft Graph
    $exchangeAppId = '00000002-0000-0ff1-ce00-000000000000'  # Office 365 Exchange Online
    # SharePoint-scoped permissions (Sites.FullControl.All, User.Read.All SP-scoped) target
    # the SharePoint resource app (00000003-0000-0ff1-ce00-000000000000) — not Graph.
    $sharePointAppId = '00000003-0000-0ff1-ce00-000000000000'

    return @(

        # ── 1. Entra ID ─────────────────────────────────────────────────────
        @{
            ShortName   = 'EntraID'
            DisplayName = "$AppNamePrefix - Entra ID"
            Description = 'M365DSC service principal for Entra ID, Identity Governance, Conditional Access, and Directory reads.'
            Permissions = @(
                @{ ResourceAppId = $graphAppId; Scope = 'Organization.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'AccessReview.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Policy.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'AdministrativeUnit.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'RoleManagement.Read.Directory' }
                @{ ResourceAppId = $graphAppId; Scope = 'Agreement.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Application.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'CustomSecAttributeDefinition.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Policy.Read.ConditionalAccess' }
                @{ ResourceAppId = $graphAppId; Scope = 'Policy.Read.AuthenticationMethod' }
                @{ ResourceAppId = $graphAppId; Scope = 'Group.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'UserAuthenticationMethod.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'User.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Directory.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Policy.Read.DeviceConfiguration' }
                @{ ResourceAppId = $graphAppId; Scope = 'Domain.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'EntitlementManagement.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'NetworkAccess.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Device.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'ReportSettings.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'PrivilegedEligibilitySchedule.Read.AzureADGroup' }
                @{ ResourceAppId = $graphAppId; Scope = 'RoleManagementPolicy.Read.Directory' }
                @{ ResourceAppId = $graphAppId; Scope = 'GroupSettings.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'APIConnectors.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'IdentityUserFlow.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'LifecycleWorkflows.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'ProgramControl.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Policy.Read.IdentityProtection' }
                @{ ResourceAppId = $graphAppId; Scope = 'NetworkAccessPolicy.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'RoleAssignmentSchedule.Read.Directory' }
                @{ ResourceAppId = $graphAppId; Scope = 'RoleEligibilitySchedule.Read.Directory' }
                @{ ResourceAppId = $graphAppId; Scope = 'IdentityProvider.Read.All' }
            )
        }

        # ── 2. Microsoft Teams ──────────────────────────────────────────────
        @{
            ShortName   = 'Teams'
            DisplayName = "$AppNamePrefix - Microsoft Teams"
            Description = 'M365DSC service principal for Microsoft Teams topology, channel settings, and meeting policy reads.'
            Permissions = @(
                @{ ResourceAppId = $graphAppId; Scope = 'Organization.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'TeamSettings.ReadWrite.All' }  # Flagged ★ — test TeamSettings.Read.All
                @{ ResourceAppId = $graphAppId; Scope = 'TeamSettings.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'ChannelSettings.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Group.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Team.ReadBasic.All' }
            )
        }

        # ── 3. Security, Compliance & Exchange ─────────────────────────────
        @{
            ShortName            = 'ExchangeCompliance'
            DisplayName          = "$AppNamePrefix - Security, Compliance & Exchange"
            Description          = 'M365DSC service principal for Exchange Online configuration, compliance policies, retention, and DLP reads.'
            ExchangeRoleRequired = $true
            Permissions          = @(
                @{ ResourceAppId = $graphAppId;    Scope = 'Organization.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'Group.Read.All' }
                @{ ResourceAppId = $exchangeAppId; Scope = 'Exchange.ManageAsApp' }
            )
        }

        # ── 4. SharePoint & OneDrive ────────────────────────────────────────
        @{
            ShortName   = 'SharePoint'
            DisplayName = "$AppNamePrefix - SharePoint & OneDrive"
            Description = 'M365DSC service principal for SharePoint Online tenant settings, site policies, and OneDrive configuration reads.'
            Permissions = @(
                @{ ResourceAppId = $graphAppId;      Scope = 'Organization.Read.All' }
                @{ ResourceAppId = $graphAppId;      Scope = 'Domain.Read.All' }
                @{ ResourceAppId = $graphAppId;      Scope = 'Group.Read.All' }
                @{ ResourceAppId = $graphAppId;      Scope = 'SharePointTenantSettings.Read.All' }
                @{ ResourceAppId = $graphAppId;      Scope = 'User.Read.All' }
                @{ ResourceAppId = $sharePointAppId; Scope = 'Sites.FullControl.All'; IsSharePoint = $true }  # Flagged ★ — test Sites.Read.All
                @{ ResourceAppId = $sharePointAppId; Scope = 'User.Read.All';         IsSharePoint = $true }
            )
        }

        # ── 5. Intune & Defender ────────────────────────────────────────────
        @{
            ShortName   = 'IntuneDefender'
            DisplayName = "$AppNamePrefix - Intune & Defender"
            Description = 'M365DSC service principal for Intune device management, compliance policies, Defender, and Windows 365 Cloud PC reads.'
            Permissions = @(
                @{ ResourceAppId = $graphAppId; Scope = 'Organization.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'Group.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'DeviceManagementConfiguration.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'CloudPC.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'DeviceManagementApps.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'DeviceManagementManagedDevices.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'DeviceManagementServiceConfig.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'DeviceManagementScripts.Read.All' }
                @{ ResourceAppId = $graphAppId; Scope = 'DeviceManagementConfiguration.ReadWrite.All' }  # Flagged ★ — test .Read.All
                @{ ResourceAppId = $graphAppId; Scope = 'DeviceManagementRBAC.Read.All' }
            )
        }

        # ── 6. O365 & Other Services ────────────────────────────────────────
        @{
            ShortName            = 'O365Services'
            DisplayName          = "$AppNamePrefix - O365 & Other Services"
            Description          = 'M365DSC service principal for Planner, Power Platform, Forms, Fabric, Viva, and other M365 org-wide service reads.'
            ExchangeRoleRequired = $true
            Permissions          = @(
                @{ ResourceAppId = $graphAppId;    Scope = 'Organization.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'PeopleSettings.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'Application.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'ExternalConnection.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'Group.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'Application.ReadWrite.All' }  # Flagged ★ — test Application.Read.All
                @{ ResourceAppId = $graphAppId;    Scope = 'ReportSettings.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'OrgSettings-Microsoft365Install.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'OrgSettings-Forms.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'OrgSettings-Todo.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'OrgSettings-AppsAndServices.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'OrgSettings-DynamicsVoice.Read.All' }
                @{ ResourceAppId = $graphAppId;    Scope = 'Tasks.Read.All' }
                @{ ResourceAppId = $exchangeAppId; Scope = 'Exchange.ManageAsApp' }
            )
        }
    )
}

#endregion

#region ── Graph Helpers ───────────────────────────────────────────────────────

function Connect-ToGraph {
    [CmdletBinding()]
    param ([string]$TenantId)

    Write-Log -Message "Connecting to Microsoft Graph (Tenant: $TenantId)..." -Level INFO

    $scopes = @(
        'Application.ReadWrite.All'
        'AppRoleAssignment.ReadWrite.All'
        'DelegatedPermissionGrant.ReadWrite.All'
        'RoleManagement.ReadWrite.Directory'
    )

    try {
        Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        Write-Log -Message "Connected as: $($ctx.Account)" -Level SUCCESS
    }
    catch {
        Write-Log -Message "Failed to connect to Microsoft Graph: $_" -Level ERROR
        throw
    }
}

function Get-GraphServicePrincipal {
    [CmdletBinding()]
    param ([string]$AppId)

    if ($script:SPNCache.ContainsKey($AppId)) {
        return $script:SPNCache[$AppId]
    }

    $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
    if ($sp) {
        $script:SPNCache[$AppId] = $sp
    }
    return $sp
}

#endregion

#region ── Certificate Functions ───────────────────────────────────────────────

function New-SelfSignedAppCertificate {
    <#
    .SYNOPSIS
        Generates a self-signed certificate, exports a PFX, and returns the public key details.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [string]                       $Subject,
        [int]                          $ExpiryYears,
        [string]                       $OutputPath,
        [System.Security.SecureString] $Password
    )

    $notAfter  = (Get-Date).AddYears($ExpiryYears)
    $storePath = 'Cert:\CurrentUser\My'

    Write-Log -Message "  Generating self-signed certificate: CN=$Subject (expires $($notAfter.ToString('yyyy-MM-dd')))" -Level INFO

    if ($PSCmdlet.ShouldProcess('Certificate store', "Create self-signed certificate CN=$Subject")) {

        $cert = New-SelfSignedCertificate `
            -Subject           "CN=$Subject" `
            -CertStoreLocation $storePath `
            -KeyExportPolicy   Exportable `
            -KeySpec           Signature `
            -KeyLength         2048 `
            -KeyAlgorithm      RSA `
            -HashAlgorithm     SHA256 `
            -NotAfter          $notAfter `
            -ErrorAction       Stop

        $certBytes  = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        $certBase64 = [Convert]::ToBase64String($certBytes)

        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }

        $safeName = $Subject -replace '[^\w\-]', '_'
        $pfxPath  = Join-Path $OutputPath "$safeName.pfx"
        $cert | Export-PfxCertificate -FilePath $pfxPath -Password $Password -Force | Out-Null

        Write-Log -Message "  PFX exported   : $pfxPath" -Level SUCCESS
        Write-Log -Message "  Thumbprint     : $($cert.Thumbprint)" -Level INFO

        # Clean up from local cert store after export
        Remove-Item -Path "$storePath\$($cert.Thumbprint)" -ErrorAction SilentlyContinue

        return [PSCustomObject]@{
            Thumbprint = $cert.Thumbprint
            CertBase64 = $certBase64
            PfxPath    = $pfxPath
            NotAfter   = $notAfter
            Subject    = "CN=$Subject"
        }
    }
    else {
        return [PSCustomObject]@{
            Thumbprint = '0000000000000000000000000000000000000000'
            CertBase64 = 'WHATIF_BASE64'
            PfxPath    = Join-Path $OutputPath "$Subject.pfx"
            NotAfter   = $notAfter
            Subject    = "CN=$Subject"
        }
    }
}

function Get-ExistingCertificate {
    <#
    .SYNOPSIS
        Loads an existing .cer or .pem file and returns the Base64-encoded public key.
    #>
    [CmdletBinding()]
    param ([string]$Path)

    $cert       = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Path)
    $certBytes  = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $certBase64 = [Convert]::ToBase64String($certBytes)

    Write-Log -Message "  Loaded certificate : $Path" -Level INFO
    Write-Log -Message "  Thumbprint         : $($cert.Thumbprint)" -Level INFO
    Write-Log -Message "  Expires            : $($cert.NotAfter.ToString('yyyy-MM-dd'))" -Level INFO

    return [PSCustomObject]@{
        Thumbprint = $cert.Thumbprint
        CertBase64 = $certBase64
        PfxPath    = $null
        NotAfter   = $cert.NotAfter
        Subject    = $cert.Subject
    }
}

function Add-CertificateToApp {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [string]   $AppObjectId,
        [string]   $AppDisplayName,
        [string]   $CertBase64,
        [string]   $Thumbprint,
        [string]   $Subject,
        [datetime] $NotAfter
    )

    Write-Log -Message "  Uploading certificate to '$AppDisplayName' (Thumbprint: $Thumbprint)" -Level INFO

    if ($PSCmdlet.ShouldProcess($AppDisplayName, 'Upload certificate credential')) {
        $keyCredential = @{
            type        = 'AsymmetricX509Cert'
            usage       = 'Verify'
            key         = [Convert]::FromBase64String($CertBase64)
            displayName = "M365DSC Certificate - Expires $($NotAfter.ToString('yyyy-MM-dd'))"
            endDateTime = $NotAfter.ToUniversalTime().ToString('o')
        }

        Update-MgApplication -ApplicationId $AppObjectId -KeyCredentials @($keyCredential) -ErrorAction Stop
        Write-Log -Message "  Certificate uploaded successfully." -Level SUCCESS
    }
}

#endregion

#region ── Client Secret Functions ─────────────────────────────────────────────

function Add-ClientSecretToApp {
    <#
    .SYNOPSIS
        Creates a client secret on an App Registration.
    .NOTES
        The secret value is only available at creation time.
        It is displayed on-screen but intentionally NOT written to the log file.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [string] $AppObjectId,
        [string] $AppDisplayName,
        [int]    $ExpiryYears
    )

    $endDate = (Get-Date).AddYears($ExpiryYears)
    Write-Log -Message "  Creating client secret for '$AppDisplayName' (expires $($endDate.ToString('yyyy-MM-dd')))" -Level INFO

    if ($PSCmdlet.ShouldProcess($AppDisplayName, 'Create client secret')) {
        $secret = Add-MgApplicationPassword `
            -ApplicationId      $AppObjectId `
            -PasswordCredential @{
                displayName = "M365DSC Secret - Created $(Get-Date -Format 'yyyy-MM-dd')"
                endDateTime = $endDate.ToUniversalTime().ToString('o')
            } `
            -ErrorAction Stop

        Write-Log -Message "  Client secret created. ID: $($secret.KeyId)" -Level SUCCESS

        # Display to console only — secret value is never written to log
        Write-Host ''
        Write-Host '  ╔══════════════════════════════════════════════════════════════╗' -ForegroundColor Magenta
        Write-Host '  ║  SECRET VALUE — Copy this now. It will not be shown again.  ║' -ForegroundColor Magenta
        Write-Host "  ║  App    : $AppDisplayName" -ForegroundColor Magenta
        Write-Host "  ║  Secret : $($secret.SecretText)" -ForegroundColor Magenta
        Write-Host '  ╚══════════════════════════════════════════════════════════════╝' -ForegroundColor Magenta
        Write-Host ''

        return [PSCustomObject]@{
            SecretId  = $secret.KeyId
            ExpiresOn = $endDate
        }
    }
    else {
        return [PSCustomObject]@{
            SecretId  = 'WHATIF-GUID'
            ExpiresOn = $endDate
        }
    }
}

#endregion

#region ── App Registration Creation ───────────────────────────────────────────

function New-AppRegistration {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [hashtable] $AppDef
    )

    $displayName = $AppDef.DisplayName
    Write-Log -Message '' -Level INFO
    Write-Log -Message "━━━ Processing: $displayName" -Level INFO

    # ── Build RequiredResourceAccess manifest ────────────────────────────
    $resourceGroups         = $AppDef.Permissions | Group-Object -Property ResourceAppId
    $requiredResourceAccess = [System.Collections.Generic.List[object]]::new()

    foreach ($group in $resourceGroups) {
        $resourceAppId = $group.Name
        $resourceSP    = Get-GraphServicePrincipal -AppId $resourceAppId

        if (-not $resourceSP) {
            Write-Log -Message "  Resource SP not found for AppId '$resourceAppId' — skipping permission group." -Level WARN
            continue
        }

        $roleAccess = [System.Collections.Generic.List[object]]::new()

        foreach ($perm in $group.Group) {
            $role = $resourceSP.AppRoles | Where-Object {
                $_.Value -eq $perm.Scope -and $_.AllowedMemberTypes -contains 'Application'
            }

            if ($role) {
                $roleAccess.Add(@{ id = $role.Id; type = 'Role' })
                Write-Log -Message "    Resolved: $($perm.Scope)" -Level DEBUG
            }
            else {
                Write-Log -Message "    Not found: '$($perm.Scope)' on resource '$resourceAppId'" -Level WARN
            }
        }

        if ($roleAccess.Count -gt 0) {
            $requiredResourceAccess.Add(@{
                resourceAppId  = $resourceAppId
                resourceAccess = $roleAccess.ToArray()
            })
        }
    }

    # ── Create or skip existing App Registration ─────────────────────────
    $existingApp = Get-MgApplication -Filter "displayName eq '$displayName'" -ErrorAction SilentlyContinue

    if ($existingApp) {
        Write-Log -Message "  App already exists: AppId=$($existingApp.AppId) — skipping creation." -Level WARN
        $app = $existingApp
    }
    else {
        if ($PSCmdlet.ShouldProcess($displayName, 'Create App Registration')) {
            Write-Log -Message "  Creating App Registration..." -Level INFO

            $app = New-MgApplication `
                -DisplayName            $displayName `
                -Description            $AppDef.Description `
                -SignInAudience         'AzureADMyOrg' `
                -RequiredResourceAccess $requiredResourceAccess.ToArray() `
                -ErrorAction            Stop

            Write-Log -Message "  Created: AppId=$($app.AppId) | ObjectId=$($app.Id)" -Level SUCCESS
            Start-Sleep -Seconds 3  # allow Entra propagation
        }
        else {
            Write-Log -Message "  [WhatIf] Would create App Registration: $displayName" -Level WARN
            return [PSCustomObject]@{
                ShortName            = $AppDef.ShortName
                DisplayName          = $displayName
                AppId                = 'WHATIF-APP-ID'
                ObjectId             = 'WHATIF-OBJECT-ID'
                Certificates         = @()
                Secrets              = @()
                ExchangeRoleRequired = ($AppDef.ContainsKey('ExchangeRoleRequired') -and $AppDef.ExchangeRoleRequired)
            }
        }
    }

    $result = [PSCustomObject]@{
        ShortName            = $AppDef.ShortName
        DisplayName          = $displayName
        AppId                = $app.AppId
        ObjectId             = $app.Id
        Certificates         = [System.Collections.Generic.List[object]]::new()
        Secrets              = [System.Collections.Generic.List[object]]::new()
        ExchangeRoleRequired = ($AppDef.ContainsKey('ExchangeRoleRequired') -and $AppDef.ExchangeRoleRequired)
    }

    # ── Self-Signed Certificate ──────────────────────────────────────────
    if ($CreateSelfSignedCert) {
        try {
            $certInfo = New-SelfSignedAppCertificate `
                -Subject     $displayName `
                -ExpiryYears $CertificateExpiryYears `
                -OutputPath  $CertificateOutputPath `
                -Password    $CertificatePassword `
                -WhatIf:($WhatIfPreference)

            Add-CertificateToApp `
                -AppObjectId    $app.Id `
                -AppDisplayName $displayName `
                -CertBase64     $certInfo.CertBase64 `
                -Thumbprint     $certInfo.Thumbprint `
                -Subject        $certInfo.Subject `
                -NotAfter       $certInfo.NotAfter `
                -WhatIf:($WhatIfPreference)

            $result.Certificates.Add($certInfo)
        }
        catch {
            Write-Log -Message "  Certificate creation failed for '$displayName': $_" -Level ERROR
        }
    }

    # ── Import Existing Certificate ──────────────────────────────────────
    if ($ImportCert) {
        try {
            $certInfo = Get-ExistingCertificate -Path $CertificatePath

            Add-CertificateToApp `
                -AppObjectId    $app.Id `
                -AppDisplayName $displayName `
                -CertBase64     $certInfo.CertBase64 `
                -Thumbprint     $certInfo.Thumbprint `
                -Subject        $certInfo.Subject `
                -NotAfter       $certInfo.NotAfter `
                -WhatIf:($WhatIfPreference)

            $result.Certificates.Add($certInfo)
        }
        catch {
            Write-Log -Message "  Certificate import failed for '$displayName': $_" -Level ERROR
        }
    }

    # ── Client Secret ────────────────────────────────────────────────────
    if ($CreateClientSecret) {
        try {
            $secretInfo = Add-ClientSecretToApp `
                -AppObjectId    $app.Id `
                -AppDisplayName $displayName `
                -ExpiryYears    $SecretExpiryYears `
                -WhatIf:($WhatIfPreference)

            $result.Secrets.Add($secretInfo)
        }
        catch {
            Write-Log -Message "  Client secret creation failed for '$displayName': $_" -Level ERROR
        }
    }

    return $result
}

#endregion

#region ── Admin Consent ───────────────────────────────────────────────────────

function Grant-AppAdminConsent {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [string]    $AppId,
        [string]    $DisplayName,
        [hashtable] $AppDef
    )

    Write-Log -Message "  Granting admin consent for: $DisplayName" -Level INFO

    if ($PSCmdlet.ShouldProcess($DisplayName, 'Grant admin consent')) {

        # Ensure a Service Principal exists for this app
        $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
        if (-not $sp) {
            Write-Log -Message "  Creating Service Principal for AppId: $AppId" -Level INFO
            $sp = New-MgServicePrincipal -AppId $AppId -ErrorAction Stop
            Start-Sleep -Seconds 5
        }

        $resourceGroups = $AppDef.Permissions | Group-Object -Property ResourceAppId

        foreach ($group in $resourceGroups) {
            $resourceSP = Get-GraphServicePrincipal -AppId $group.Name
            if (-not $resourceSP) {
                Write-Log -Message "  Resource SP not found for consent: $($group.Name)" -Level WARN
                continue
            }

            foreach ($perm in $group.Group) {
                $role = $resourceSP.AppRoles | Where-Object {
                    $_.Value -eq $perm.Scope -and $_.AllowedMemberTypes -contains 'Application'
                }

                if (-not $role) {
                    Write-Log -Message "    Skipping consent (role not resolved): $($perm.Scope)" -Level WARN
                    continue
                }

                $alreadyGranted = Get-MgServicePrincipalAppRoleAssignment `
                    -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue |
                    Where-Object { $_.AppRoleId -eq $role.Id }

                if ($alreadyGranted) {
                    Write-Log -Message "    Already consented: $($perm.Scope)" -Level DEBUG
                    continue
                }

                try {
                    New-MgServicePrincipalAppRoleAssignment `
                        -ServicePrincipalId $sp.Id `
                        -PrincipalId        $sp.Id `
                        -ResourceId         $resourceSP.Id `
                        -AppRoleId          $role.Id `
                        -ErrorAction        Stop | Out-Null

                    Write-Log -Message "    Consented: $($perm.Scope)" -Level SUCCESS
                }
                catch {
                    Write-Log -Message "    Consent failed for '$($perm.Scope)': $_" -Level WARN
                }
            }
        }
    }
    else {
        Write-Log -Message "  [WhatIf] Would grant admin consent for: $DisplayName" -Level WARN
    }
}

#endregion

#region ── Exchange Directory Role Assignment ───────────────────────────────────

function Grant-ExchangeDirectoryRole {
    <#
    .SYNOPSIS
        Assigns a direct active Entra directory role to a Service Principal.

    .DESCRIPTION
        Uses New-MgRoleManagementDirectoryRoleAssignment to create a permanent active role
        assignment on the target Service Principal. This is not a PIM-eligible assignment —
        the role is active immediately with no activation step required.

        Called for the two apps that require Exchange.ManageAsApp to function:
            - M365DSC - Security, Compliance & Exchange
            - M365DSC - O365 & Other Services

    .PARAMETER ServicePrincipalId
        Object ID of the Service Principal to assign the role to.

    .PARAMETER ServicePrincipalName
        Display name used in log output.

    .PARAMETER RoleName
        The built-in Entra role to assign: GlobalReader or ExchangeAdministrator.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)] [string] $ServicePrincipalId,
        [Parameter(Mandatory)] [string] $ServicePrincipalName,
        [Parameter(Mandatory)] [ValidateSet('GlobalReader', 'ExchangeAdministrator')]
        [string] $RoleName
    )

    # Map friendly name to the Entra built-in role template display name
    $roleDisplayName = switch ($RoleName) {
        'GlobalReader'           { 'Global Reader' }
        'ExchangeAdministrator'  { 'Exchange Administrator' }
    }

    Write-Log -Message "  Assigning role '$roleDisplayName' to '$ServicePrincipalName'..." -Level INFO

    if ($PSCmdlet.ShouldProcess($ServicePrincipalName, "Assign directory role '$roleDisplayName'")) {

        # Resolve the role definition ID from the tenant (template IDs are well-known but
        # using Get-MgRoleManagementDirectoryRoleDefinition ensures the role exists in this tenant)
        $roleDef = Get-MgRoleManagementDirectoryRoleDefinition `
            -Filter "displayName eq '$roleDisplayName'" `
            -ErrorAction Stop

        if (-not $roleDef) {
            Write-Log -Message "  Role definition not found: '$roleDisplayName' — skipping." -Level ERROR
            return
        }

        # Check for an existing active assignment to avoid duplicates
        $existingAssignment = Get-MgRoleManagementDirectoryRoleAssignment `
            -Filter "roleDefinitionId eq '$($roleDef.Id)' and principalId eq '$ServicePrincipalId'" `
            -ErrorAction SilentlyContinue

        if ($existingAssignment) {
            Write-Log -Message "  Role '$roleDisplayName' is already assigned to '$ServicePrincipalName' — skipping." -Level WARN
            return
        }

        try {
            # directoryScopeId '/' = tenant-wide scope (required for built-in directory roles)
            $assignment = New-MgRoleManagementDirectoryRoleAssignment `
                -PrincipalId      $ServicePrincipalId `
                -RoleDefinitionId $roleDef.Id `
                -DirectoryScopeId '/' `
                -ErrorAction      Stop

            Write-Log -Message "  Role '$roleDisplayName' assigned successfully. AssignmentId: $($assignment.Id)" -Level SUCCESS
        }
        catch {
            Write-Log -Message "  Failed to assign role '$roleDisplayName' to '$ServicePrincipalName': $_" -Level ERROR
        }
    }
    else {
        Write-Log -Message "  [WhatIf] Would assign role '$roleDisplayName' to '$ServicePrincipalName'." -Level WARN
    }
}

#endregion


function Write-Summary {
    [CmdletBinding()]
    param ([object[]]$Results)

    $d = '═══════════════════════════════════════════════════════════════════'

    Write-Log -Message $d -Level INFO
    Write-Log -Message '  M365DSC App Registration Setup - Complete' -Level INFO
    Write-Log -Message $d -Level INFO

    foreach ($r in $Results) {
        Write-Log -Message '' -Level INFO
        Write-Log -Message "  App      : $($r.DisplayName)" -Level INFO
        Write-Log -Message "  AppId    : $($r.AppId)"       -Level INFO
        Write-Log -Message "  ObjectId : $($r.ObjectId)"    -Level INFO

        foreach ($cert in $r.Certificates) {
            Write-Log -Message "  Cert Thumbprint : $($cert.Thumbprint)"                      -Level INFO
            Write-Log -Message "  Cert Expires    : $($cert.NotAfter.ToString('yyyy-MM-dd'))" -Level INFO
            if ($cert.PfxPath) {
                Write-Log -Message "  PFX Path        : $($cert.PfxPath)" -Level INFO
            }
        }

        foreach ($s in $r.Secrets) {
            Write-Log -Message "  Secret ID      : $($s.SecretId)"                         -Level INFO
            Write-Log -Message "  Secret Expires : $($s.ExpiresOn.ToString('yyyy-MM-dd'))" -Level INFO
        }

        if ($r.ExchangeRoleRequired) {
            if ($AssignExchangeRole) {
                Write-Log -Message "  Exchange Role  : $ExchangeRole (assigned)" -Level INFO
            }
            else {
                Write-Log -Message "  ACTION REQUIRED: Exchange role not yet assigned." -Level WARN
                Write-Log -Message "  Re-run with -AssignExchangeRole, or manually assign 'Global Reader'" -Level WARN
                Write-Log -Message "  or 'Exchange Administrator' to this app's Service Principal." -Level WARN
            }
        }
    }

    Write-Log -Message '' -Level INFO
    Write-Log -Message $d -Level INFO
    Write-Log -Message '  NEXT STEPS' -Level INFO
    Write-Log -Message $d -Level INFO
    Write-Log -Message '  1. If using certificates: upload PFX files to Azure Key Vault.' -Level INFO
    Write-Log -Message '  2. If Exchange role was not assigned: re-run with -AssignExchangeRole or assign manually.' -Level INFO
    Write-Log -Message '  3. Configure M365DSC workload authentication with each AppId + certificate thumbprint.' -Level INFO
    Write-Log -Message '  4. Test with Get-M365DSCCompiledPermissionList and a config export run.' -Level INFO
    Write-Log -Message '  5. Review flagged (star) permissions and test least-privilege alternatives.' -Level INFO
    Write-Log -Message '' -Level INFO
    Write-Log -Message "  Log file: $($script:LogFile)" -Level INFO
}

#endregion

#region ── Main ────────────────────────────────────────────────────────────────

function Invoke-Main {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Initialize-Log -Path $LogPath

    $d = '═══════════════════════════════════════════════════════════════════'
    Write-Log -Message $d -Level INFO
    Write-Log -Message '  New-M365DSCAppRegistrations.ps1  v1.2.0' -Level INFO
    Write-Log -Message "  Tenant            : $TenantId" -Level INFO
    Write-Log -Message "  SelfSignedCert    : $CreateSelfSignedCert" -Level INFO
    Write-Log -Message "  ImportCert        : $ImportCert" -Level INFO
    Write-Log -Message "  ClientSecret      : $CreateClientSecret" -Level INFO
    Write-Log -Message "  AdminConsent      : $GrantAdminConsent" -Level INFO
    Write-Log -Message "  AssignExchangeRole: $AssignExchangeRole$(if ($AssignExchangeRole) { " ($ExchangeRole)" })" -Level INFO
    Write-Log -Message "  WhatIf            : $($WhatIfPreference)" -Level INFO
    Write-Log -Message $d -Level INFO

    # Prompt for PFX password if needed and not already provided
    if ($CreateSelfSignedCert -and -not $CertificatePassword) {
        Write-Log -Message 'PFX password not provided — prompting interactively.' -Level INFO
        $CertificatePassword = Read-Host -Prompt 'Enter password for PFX certificate export' -AsSecureString
    }

    # Initialise Service Principal lookup cache
    $script:SPNCache = @{}

    Connect-ToGraph -TenantId $TenantId

    $appDefinitions = Get-AppDefinitions
    $results        = [System.Collections.Generic.List[object]]::new()

    foreach ($appDef in $appDefinitions) {
        try {
            $appResult = New-AppRegistration -AppDef $appDef -WhatIf:($WhatIfPreference)
            $results.Add($appResult)

            if ($GrantAdminConsent -and $appResult.AppId -ne 'WHATIF-APP-ID') {
                Grant-AppAdminConsent `
                    -AppId       $appResult.AppId `
                    -DisplayName $appResult.DisplayName `
                    -AppDef      $appDef `
                    -WhatIf:($WhatIfPreference)
            }

            # Assign Exchange directory role to Exchange-dependent apps if requested
            if ($AssignExchangeRole -and $appResult.ExchangeRoleRequired) {
                # Ensure the Service Principal exists (may have been created during consent step)
                $sp = Get-MgServicePrincipal -Filter "appId eq '$($appResult.AppId)'" -ErrorAction SilentlyContinue
                if (-not $sp) {
                    Write-Log -Message "  Creating Service Principal for role assignment: $($appResult.AppId)" -Level INFO
                    $sp = New-MgServicePrincipal -AppId $appResult.AppId -ErrorAction Stop
                    Start-Sleep -Seconds 5
                }

                Grant-ExchangeDirectoryRole `
                    -ServicePrincipalId   $sp.Id `
                    -ServicePrincipalName $appResult.DisplayName `
                    -RoleName             $ExchangeRole `
                    -WhatIf:($WhatIfPreference)
            }
        }
        catch {
            Write-Log -Message "Unhandled error for '$($appDef.DisplayName)': $_" -Level ERROR
        }
    }

    Write-Summary -Results $results

    Write-Log -Message 'Script completed.' -Level SUCCESS
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

Invoke-Main

#endregion

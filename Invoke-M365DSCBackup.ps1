#Requires -Version 5.1
#Requires -Modules Microsoft365DSC, Az.Storage, Az.Accounts

<#
.SYNOPSIS
    Runs M365DSC configuration backups for one or more workloads, generates HTML reports,
    uploads artifacts to Azure Blob Storage, and emails a summary on completion.

.DESCRIPTION
    Invoke-M365DSCBackup orchestrates the full M365DSC export workflow:
      - Exports tenant configuration for one or more named workloads using Export-M365DSCConfiguration
      - Stores each backup in a timestamped subfolder under a workload-specific parent folder
        e.g. <BackupRoot>\SPO_Configs\SPO_ODFB_Backup_2026-03-27_14-30-00\
      - Generates an HTML report per workload via New-M365DSCReportFromConfiguration and saves
        it under <BackupRoot>\Reports\<WorkloadName>_Report_<Timestamp>.html
      - Uploads all backup and report files to an Azure Storage Account blob container
      - Sends an HTML summary email via Microsoft Graph (Send Mail) listing success/failure
        per workload, resource counts, file paths, and blob URLs

    Intended to run as one batch job per workload (e.g. a separate job for Exchange, Purview,
    SharePoint, OneDrive, etc.). Each job is fully independent and produces its own backup
    folder, report, blob upload, and summary email.

    Authentication uses per-workload service principal credentials defined in the config file.
    Workloads that share a real-world SPN (e.g. Exchange + Purview share SPN 3; SharePoint +
    OneDrive share SPN 4; Intune + Defender share SPN 5) simply reference the same ApplicationId
    and CertificateThumbprint in their respective config entries — the jobs still run separately.

    Workload definitions and global settings can be supplied either via parameters or a JSON
    config file. Parameters always take precedence over the config file.

.PARAMETER ConfigFile
    Path to a JSON configuration file. See M365DSCBackup.config.json for the schema.
    Defaults to M365DSCBackup.config.json in the same directory as this script.
    If the default file is not found, all required values must be supplied via individual parameters.

.PARAMETER Workloads
    Array of workload names to back up. Each maps to its own folder, report, and blob path.
    Supported values:
      Entra, Teams, Exchange, Purview, SharePoint, OneDrive, Intune, Defender, O365Services
    Example: -Workloads @('Exchange')   # typical single-workload batch job invocation

.PARAMETER BackupRoot
    Root path where all backup folders and the Reports subfolder will be created.
    Example: C:\M365DSCBackups

.PARAMETER TenantId
    Entra ID Tenant ID (GUID format) used for Azure Storage and Graph Mail authentication.

.PARAMETER TenantName
    Tenant domain name (e.g. contoso.onmicrosoft.com) passed as -TenantId to
    Export-M365DSCConfiguration. M365DSC requires the domain name, not the GUID.

.PARAMETER StorageAccountName
    Azure Storage Account name for uploading backup and report artifacts.

.PARAMETER StorageContainerName
    Azure Blob container name within the storage account. Defaults to 'm365dsc-backups'.

.PARAMETER StorageResourceGroup
    Resource group containing the Azure Storage Account (used for Az context lookup).

.PARAMETER StorageSubscriptionId
    Azure Subscription ID containing the Storage Account.

.PARAMETER StorageSPAppId
    Application (client) ID of the service principal used for Azure Storage and Graph Mail.

.PARAMETER StorageSPCertThumbprint
    Certificate thumbprint for the Storage/Graph service principal.

.PARAMETER UseManagedIdentity
    Switch. If set, authenticates to Azure Storage using a Managed Identity instead of a
    service principal. Supports both system-assigned and user-assigned identities.
    When set, StorageSPAppId and StorageSPCertThumbprint are not required for the upload step.

.PARAMETER ManagedIdentityClientId
    Client ID of a user-assigned Managed Identity. Only applicable when -UseManagedIdentity
    is set. Omit to use the system-assigned identity on the host running this script.

.PARAMETER EmailFrom
    Sender address for the summary email (must be a licensed mailbox in the tenant).

.PARAMETER EmailTo
    Array of recipient addresses for the summary email.

.PARAMETER EmailSPAppId
    Application (client) ID of the dedicated service principal used for Graph mail send.
    Corresponds to Email.AppId in the config file.

.PARAMETER EmailSPCertThumbprint
    Certificate thumbprint for the mail send service principal.
    Corresponds to Email.CertThumbprint in the config file.

.PARAMETER SkipUpload
    Switch. If set, skips the Azure Storage upload step.

.PARAMETER SkipEmail
    Switch. If set, skips sending the summary email.

.PARAMETER SkipReport
    Switch. If set, skips generating the HTML configuration report.

.PARAMETER LogPath
    Directory for log files. Defaults to C:\softdist\Logs\M365DSCBackup.

.EXAMPLE
    # Run using config file only
    .\Invoke-M365DSCBackup.ps1 -ConfigFile .\M365DSCBackup.config.json

.EXAMPLE
    # Override workloads at runtime; everything else from config
    .\Invoke-M365DSCBackup.ps1 -ConfigFile .\M365DSCBackup.config.json -Workloads @('Exchange','Teams')

.EXAMPLE
    # Fully parameter-driven, skip upload and email (useful for dev/test)
    .\Invoke-M365DSCBackup.ps1 `
        -Workloads @('SharePoint','OneDrive') `
        -BackupRoot 'C:\M365DSCBackups' `
        -TenantId '00000000-0000-0000-0000-000000000000' `
        -SkipUpload -SkipEmail

.NOTES
    Author      : IT Infrastructure
    Version     : 1.6.2
    Requires    : Microsoft365DSC, Az.Storage, Az.Accounts modules
    Auth model  : Per-workload credentials in config file. Workloads that share a real-world SPN
                  (Exchange+Purview, SharePoint+OneDrive, Intune+Defender) use the same AppId and
                  CertThumbprint in their respective config entries.
    Workloads   : Entra, Teams, Exchange, Purview, SharePoint, OneDrive, Intune, Defender, O365Services
    Log path    : C:\softdist\Logs\M365DSCBackup\ (default)
    Known issue : New-M365DSCReportFromConfiguration requires DSCParser module; ensure it is
                  current. Report generation errors are non-fatal - backup files are still kept.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()] [string]   $ConfigFile = (Join-Path $PSScriptRoot 'M365DSCBackup.config.json'),
    [Parameter()] [string[]] $Workloads,
    [Parameter()] [string]   $BackupRoot,
    [Parameter()] [string]   $TenantId,
    [Parameter()] [string]   $TenantName,
    [Parameter()] [string]   $StorageAccountName,
    [Parameter()] [string]   $StorageContainerName,
    [Parameter()] [string]   $StorageResourceGroup,
    [Parameter()] [string]   $StorageSubscriptionId,
    [Parameter()] [string]   $StorageSPAppId,
    [Parameter()] [string]   $StorageSPCertThumbprint,
    [Parameter()] [switch]   $UseManagedIdentity,
    [Parameter()] [string]   $ManagedIdentityClientId,
    [Parameter()] [string]   $EmailFrom,
    [Parameter()] [string[]] $EmailTo,
    [Parameter()] [string]   $EmailSPAppId,
    [Parameter()] [string]   $EmailSPCertThumbprint,
    [Parameter()] [switch]   $SkipUpload,
    [Parameter()] [switch]   $SkipEmail,
    [Parameter()] [switch]   $SkipReport,
    [Parameter()] [string]   $LogPath = 'C:\softdist\Logs\M365DSCBackup'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Workload Definitions ────────────────────────────────────────────────
# Maps a friendly workload name to:
#   FolderPrefix   : Parent folder name under BackupRoot  (e.g. SPO_Configs)
#   BackupPrefix   : Timestamped subfolder prefix         (e.g. SPO_ODFB_Backup)
#   DisplayName    : Human-readable name used in logs and email reports
#   Components     : Default M365DSC component array for Export-M365DSCConfiguration
#
# Each workload is a separate, independently runnable job. Multiple workloads may share
# the same service principal credentials in the config file (e.g. Exchange and Purview
# share SPN 3; SharePoint and OneDrive share SPN 4; Intune and Defender share SPN 5)
# — just set the same ApplicationId and CertificateThumbprint for those workload entries.
# Components can be overridden per-workload in the config file.
$Script:WorkloadMap = @{

    # ── Entra ID & Identity ─────────────────────────────────────────────────────
    # SPN: Entra ID & Identity SP
    # Graph: Organization.Read.All, AccessReview.Read.All, Policy.Read.All,
    #        AdministrativeUnit.Read.All, RoleManagement.Read.Directory, Agreement.Read.All,
    #        Application.Read.All, CustomSecAttributeDefinition.Read.All,
    #        Policy.Read.ConditionalAccess, Policy.Read.AuthenticationMethod,
    #        Group.Read.All, UserAuthenticationMethod.Read.All, User.Read.All,
    #        Directory.Read.All, Policy.Read.DeviceConfiguration, Domain.Read.All,
    #        EntitlementManagement.Read.All, NetworkAccess.Read.All, Device.Read.All,
    #        ReportSettings.Read.All, PrivilegedEligibilitySchedule.Read.AzureADGroup,
    #        RoleManagementPolicy.Read.Directory, GroupSettings.Read.All,
    #        APIConnectors.Read.All, IdentityUserFlow.Read.All, LifecycleWorkflows.Read.All,
    #        ProgramControl.Read.All, Policy.Read.IdentityProtection,
    #        NetworkAccessPolicy.Read.All, RoleAssignmentSchedule.Read.Directory,
    #        RoleEligibilitySchedule.Read.Directory, IdentityProvider.Read.All
    Entra = @{
        FolderPrefix = 'Entra_Configs'
        BackupPrefix = 'Entra_Backup'
        DisplayName  = 'Entra ID & Identity'
        Components   = @(
            'AADAccessReviewDefinition', 'AADAccessReviewPolicy',
            'AADActivityBasedTimeoutPolicy', 'AADAdminConsentRequestPolicy',
            'AADAdministrativeUnit', 'AADAgreement', 'AADAppManagementPolicy',
            'AADApplication', 'AADAttributeSet', 'AADAuthenticationContextClassReference',
            'AADAuthenticationFlowPolicy', 'AADAuthenticationMethodPolicy',
            'AADAuthenticationMethodPolicyAuthenticator', 'AADAuthenticationMethodPolicyEmail',
            'AADAuthenticationMethodPolicyExternal', 'AADAuthenticationMethodPolicyFido2',
            'AADAuthenticationMethodPolicySms', 'AADAuthenticationMethodPolicySoftwareOath',
            'AADAuthenticationMethodPolicyTemporaryAccessPass',
            'AADAuthenticationMethodPolicyVoice', 'AADAuthenticationMethodPolicyX509',
            'AADAuthorizationPolicy', 'AADConditionalAccessPolicy',
            'AADCrossTenantAccessPolicy', 'AADCrossTenantAccessPolicyConfigurationDefault',
            'AADCustomSecurityAttributeDefinition', 'AADEntitlementManagementAccessPackage',
            'AADEntitlementManagementAccessPackageCatalog',
            'AADEntitlementManagementConnectedOrganization',
            'AADEntitlementManagementSettings', 'AADExternalIdentityPolicy',
            'AADGroup', 'AADGroupLifecyclePolicy', 'AADGroupsNamingPolicy',
            'AADGroupsSettings', 'AADIdentityGovernanceLifecycleWorkflow',
            'AADIdentityGovernanceLifecycleWorkflowSettings',
            'AADLifecycleWorkflowsSettings', 'AADNamedLocationPolicy',
            'AADNetworkAccessForwardingProfile', 'AADNetworkAccessForwardingPolicyLink',
            'AADNetworkAccessLocalNetworkGateway', 'AADNetworkAccessSettings',
            'AADNetworkAccessTenantSettings', 'AADPasswordResetPolicy',
            'AADPrivilegedIdentityManagementSettings', 'AADRoleAssignmentScheduleRequest',
            'AADRoleDefinition', 'AADRoleEligibilityScheduleRequest',
            'AADRoleSetting', 'AADSocialIdentityProvider',
            'AADTenantDetails', 'AADTokenIssuancePolicy', 'AADTokenLifetimePolicy',
            'AADUserFlowAttribute', 'AADVerifiedIdAuthority'
        )
    }

    # ── Microsoft Teams ─────────────────────────────────────────────────────────
    # SPN: Teams SP
    # Graph: Organization.Read.All, TeamSettings.ReadWrite.All (★ — test Read.All),
    #        TeamSettings.Read.All, ChannelSettings.Read.All, Group.Read.All,
    #        Team.ReadBasic.All
    Teams = @{
        FolderPrefix = 'Teams_Configs'
        BackupPrefix = 'Teams_Backup'
        DisplayName  = 'Microsoft Teams'
        Components   = @(
            'TeamsAppPermissionPolicy', 'TeamsAppSetupPolicy',
            'TeamsAudioConferencingPolicy', 'TeamsCallHoldPolicy',
            'TeamsCallingPolicy', 'TeamsCallParkPolicy', 'TeamsChannelsPolicy',
            'TeamsClientConfiguration', 'TeamsComplianceRecordingPolicy',
            'TeamsCortanaPolicy', 'TeamsEmergencyCallingPolicy',
            'TeamsEmergencyCallRoutingPolicy', 'TeamsFeedbackPolicy',
            'TeamsFilesPolicy', 'TeamsGuestCallingConfiguration',
            'TeamsGuestMeetingConfiguration', 'TeamsGuestMessagingConfiguration',
            'TeamsIPPhonePolicy', 'TeamsMeetingBroadcastConfiguration',
            'TeamsMeetingBroadcastPolicy', 'TeamsMeetingConfiguration',
            'TeamsMeetingPolicy', 'TeamsMessagingPolicy', 'TeamsMobilityPolicy',
            'TeamsNetworkRoamingPolicy', 'TeamsShiftsPolicy',
            'TeamsSurvivableBranchAppliance', 'TeamsSurvivableBranchAppliancePolicy',
            'TeamsTranslationRule', 'TeamsUnassignedNumberTreatment',
            'TeamsUpdateManagementPolicy', 'TeamsUpgradeConfiguration',
            'TeamsUpgradePolicy', 'TeamsVoiceRoute', 'TeamsVoiceRoutingPolicy'
        )
    }

    # ── Exchange Online ─────────────────────────────────────────────────────────
    # SPN: Security, Compliance & Exchange SP (shared with Purview below)
    # Graph: Organization.Read.All, Group.Read.All
    # Exchange: Exchange.ManageAsApp
    Exchange = @{
        FolderPrefix = 'Exchange_Configs'
        BackupPrefix = 'Exchange_Backup'
        DisplayName  = 'Exchange Online'
        Components   = @(
            'EXOAcceptedDomain', 'EXOAddressList', 'EXOAntiPhishPolicy',
            'EXOAntiPhishRule', 'EXOAtpPolicyForO365', 'EXODkimSigningConfig',
            'EXOHostedConnectionFilterPolicy', 'EXOHostedContentFilterPolicy',
            'EXOHostedContentFilterRule', 'EXOHostedOutboundSpamFilterPolicy',
            'EXOHostedOutboundSpamFilterRule', 'EXOMalwareFilterPolicy',
            'EXOMalwareFilterRule', 'EXOOrganizationConfig',
            'EXOOrganizationRelationship', 'EXOOutboundConnector',
            'EXORemoteDomain', 'EXOReportSubmissionPolicy',
            'EXOSafeAttachmentPolicy', 'EXOSafeAttachmentRule',
            'EXOSafeLinksPolicy', 'EXOSafeLinksRule', 'EXOSharedMailbox',
            'EXOSharingPolicy', 'EXOTransportRule'
        )
    }

    # ── Purview / Security & Compliance ────────────────────────────────────────
    # SPN: Security, Compliance & Exchange SP (shared with Exchange above)
    # Graph: Organization.Read.All, Group.Read.All
    # Exchange: Exchange.ManageAsApp
    Purview = @{
        FolderPrefix = 'Purview_Configs'
        BackupPrefix = 'Purview_Backup'
        DisplayName  = 'Purview & Compliance'
        Components   = @(
            'SCAutoSensitivityLabelPolicy', 'SCAutoSensitivityLabelRule',
            'SCCaseHoldPolicy', 'SCCaseHoldRule', 'SCComplianceCase',
            'SCComplianceCaseHold', 'SCComplianceSearch',
            'SCComplianceSearchAction', 'SCComplianceTag',
            'SCDLPCompliancePolicy', 'SCDLPComplianceRule',
            'SCFilePlanPropertyAuthority', 'SCFilePlanPropertyCategory',
            'SCFilePlanPropertyCitation', 'SCFilePlanPropertyDepartment',
            'SCFilePlanPropertyReferenceId', 'SCFilePlanPropertySubCategory',
            'SCLabel', 'SCLabelPolicy', 'SCProtectionAlert',
            'SCRetentionCompliancePolicy', 'SCRetentionComplianceRule',
            'SCRetentionEventType', 'SCSensitivityLabel', 'SCSensitivityLabelPolicy',
            'SCSupervisoryReviewPolicy', 'SCSupervisoryReviewRule'
        )
    }

    # ── SharePoint Online ───────────────────────────────────────────────────────
    # SPN: SharePoint & OneDrive SP (shared with OneDrive below)
    # Graph: Organization.Read.All, Domain.Read.All, Group.Read.All,
    #        SharePointTenantSettings.Read.All, User.Read.All
    # SharePoint: Sites.FullControl.All (★ — test Sites.Read.All), User.Read.All
    SharePoint = @{
        FolderPrefix = 'SPO_Configs'
        BackupPrefix = 'SPO_Backup'
        DisplayName  = 'SharePoint Online'
        Components   = @(
            'SPOAccessControlSettings', 'SPOApp', 'SPOBrowserIdleSignout',
            'SPOBuiltInDesignPackageVisibility', 'SPOHomeSite', 'SPOHubSite',
            'SPOOrgAssetsLibrary', 'SPOSearchResultSource', 'SPOSharingSettings',
            'SPOSiteDesign', 'SPOSiteDesignRights', 'SPOSiteGroup',
            'SPOSiteScript', 'SPOStorageEntity', 'SPOTenantCdnEnabled',
            'SPOTenantCdnPolicy', 'SPOTenantSettings', 'SPOTheme',
            'SPOUserProfileProperty'
        )
    }

    # ── OneDrive for Business ───────────────────────────────────────────────────
    # SPN: SharePoint & OneDrive SP (shared with SharePoint above)
    # SharePoint and OneDrive share the same SPN — use the same ApplicationId and
    # CertificateThumbprint for both entries in the config file.
    OneDrive = @{
        FolderPrefix = 'SPO_Configs'          # intentionally shares the SPO parent folder
        BackupPrefix = 'SPO_ODFB_Backup'
        DisplayName  = 'OneDrive for Business'
        Components   = @(
            'ODSettings'
        )
    }

    # ── Microsoft Intune ────────────────────────────────────────────────────────
    # SPN: Intune & Defender SP (shared with Defender below)
    # Graph: Organization.Read.All, Group.Read.All,
    #        DeviceManagementConfiguration.Read.All, CloudPC.Read.All,
    #        DeviceManagementApps.Read.All, DeviceManagementManagedDevices.Read.All,
    #        DeviceManagementServiceConfig.Read.All, DeviceManagementScripts.Read.All,
    #        DeviceManagementConfiguration.ReadWrite.All (★ — test Read.All),
    #        DeviceManagementRBAC.Read.All
    Intune = @{
        FolderPrefix = 'Intune_Configs'
        BackupPrefix = 'Intune_Backup'
        DisplayName  = 'Microsoft Intune'
        Components   = @(
            # Device Compliance
            'IntuneDeviceCompliancePolicyAndroid',
            'IntuneDeviceCompliancePolicyAndroidDeviceOwner',
            'IntuneDeviceCompliancePolicyAndroidWorkProfile',
            'IntuneDeviceCompliancePolicyiOs', 'IntuneDeviceCompliancePolicyMacOS',
            'IntuneDeviceCompliancePolicyWindows10',
            # Device Configuration
            'IntuneDeviceConfigurationAdministrativeTemplatePolicyWindows10',
            'IntuneDeviceConfigurationCustomPolicyWindows10',
            'IntuneDeviceConfigurationPolicyWindows10',
            'IntuneDeviceConfigurationSharedMultiDevicePolicyWindows10',
            # Enrollment
            'IntuneDeviceEnrollmentLimitRestriction',
            'IntuneDeviceEnrollmentPlatformRestriction',
            'IntuneDeviceEnrollmentStatusPageWindows10',
            'IntuneWindowsAutopilotDeploymentProfile',
            # App Management
            'IntuneAppConfigurationPolicy', 'IntuneAppProtectionPolicyAndroid',
            'IntuneAppProtectionPolicyiOS', 'IntuneMobileAppsManagedGooglePlayApp',
            'IntuneWindowsInformationProtectionPolicyWindows10MdmEnrolled',
            # Settings Catalog & Update
            'IntuneSettingCatalogASRRulesPolicyWindows10',
            'IntuneSettingCatalogCustomPolicyWindows10',
            'IntuneWindowsUpdateRingUpdatePolicy',
            'IntuneDeviceConfigurationEndpointProtectionPolicyWindows10',
            # RBAC & Org
            'IntuneDeviceCategory', 'IntuneRoleAssignment', 'IntuneRoleDefinition',
            'IntuneDeviceConfigurationPolicySets'
        )
    }

    # ── Microsoft Defender ──────────────────────────────────────────────────────
    # SPN: Intune & Defender SP (shared with Intune above)
    # Use the same ApplicationId and CertificateThumbprint as Intune in the config file.
    Defender = @{
        FolderPrefix = 'Defender_Configs'
        BackupPrefix = 'Defender_Backup'
        DisplayName  = 'Microsoft Defender'
        Components   = @(
            'IntuneEndpointDetectionAndResponsePolicy',
            'IntuneAntivirusPolicyWindows10SettingCatalog',
            'IntuneFirewallPolicyWindows10',
            'IntuneFirewallRulesHyperVPolicyWindows10',
            'IntuneFirewallRulesPolicyWindows10'
        )
    }

    # ── O365 & Other Services ───────────────────────────────────────────────────
    # SPN: O365 & Other Services SP
    # Covers Planner, Power Platform, Forms, Microsoft Fabric, Viva, ToDo, and other
    # tenant-level M365 org settings not covered by the workload-specific SPNs.
    # Graph: Organization.Read.All, PeopleSettings.Read.All, Application.Read.All,
    #        ExternalConnection.Read.All, Group.Read.All,
    #        Application.ReadWrite.All (★ — test Read.All),
    #        ReportSettings.Read.All, OrgSettings-Microsoft365Install.Read.All,
    #        OrgSettings-Forms.Read.All, OrgSettings-Todo.Read.All,
    #        OrgSettings-AppsAndServices.Read.All, OrgSettings-DynamicsVoice.Read.All,
    #        Tasks.Read.All
    # Exchange: Exchange.ManageAsApp
    O365Services = @{
        FolderPrefix = 'O365Services_Configs'
        BackupPrefix = 'O365Services_Backup'
        DisplayName  = 'O365 & Other Services'
        Components   = @(
            # Org-wide M365 settings
            'O365AdminAuditLogConfig', 'O365OrgCustomizationSetting',
            'O365SearchAndIntelligenceConfigurations',
            # Microsoft Planner
            'PlannerConfig',
            # Power Platform & External Connections
            'PowerPlatformEnvironment',
            # Microsoft Forms
            'OrgSettingsForms',
            # Microsoft ToDo
            'OrgSettingsToDo',
            # Microsoft 365 Apps & Services org settings
            'OrgSettingsAppsAndServices', 'OrgSettingsDynamicsVoice',
            'OrgSettingsMicrosoftInstall',
            # Microsoft Viva
            'OrgSettingsVivaInsights'
        )
    }
}
#endregion

#region ── Logging ─────────────────────────────────────────────────────────────
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')] [string] $Level = 'INFO'
    )
    $ts      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry   = "[$ts] [$Level] $Message"
    $color   = switch ($Level) {
        'INFO'    { 'Cyan'    }
        'WARN'    { 'Yellow'  }
        'ERROR'   { 'Red'     }
        'SUCCESS' { 'Green'   }
    }
    Write-Host $entry -ForegroundColor $color
    Add-Content -Path $Script:LogFile -Value $entry
}
#endregion

#region ── Config Loading ──────────────────────────────────────────────────────
function Import-BackupConfig {
    [CmdletBinding()]
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        throw "Config file not found: $Path"
    }
    $cfg = Get-Content $Path -Raw | ConvertFrom-Json
    Write-Log "Config file loaded: $Path"
    return $cfg
}

function Resolve-Setting {
    # Returns param value if set, otherwise falls back to config value, then default.
    param($ParamValue, $ConfigValue, $Default = $null)
    if ($ParamValue) { return $ParamValue }
    if ($null -ne $ConfigValue -and $ConfigValue -ne '') { return $ConfigValue }
    return $Default
}
#endregion

#region ── Azure Storage Upload ────────────────────────────────────────────────
function Connect-AzureStorage {
    [CmdletBinding()]
    param(
        [string]$TenantId,
        [string]$SubscriptionId,
        [switch]$UseManagedIdentity,
        [string]$ManagedIdentityClientId,
        [string]$AppId,
        [string]$CertThumbprint
    )
    if ($UseManagedIdentity) {
        $connectParams = @{
            Identity    = $true
            TenantId    = $TenantId
            ErrorAction = 'Stop'
        }
        if ($SubscriptionId)          { $connectParams['Subscription'] = $SubscriptionId }
        if ($ManagedIdentityClientId) {
            $connectParams['AccountId'] = $ManagedIdentityClientId
            Write-Log "Connecting to Azure via user-assigned Managed Identity (ClientId: $ManagedIdentityClientId)..."
        } else {
            Write-Log "Connecting to Azure via system-assigned Managed Identity..."
        }
        Connect-AzAccount @connectParams | Out-Null
    } else {
        Write-Log "Connecting to Azure (SP: $AppId)..."
        Connect-AzAccount -ServicePrincipal `
            -TenantId              $TenantId `
            -ApplicationId         $AppId `
            -CertificateThumbprint $CertThumbprint `
            -Subscription          $SubscriptionId `
            -ErrorAction           Stop | Out-Null
    }
    Write-Log "Azure connection established." -Level SUCCESS
}

function Send-ToAzureBlob {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]   $StorageAccountName,
        [string]   $ContainerName,
        [string[]] $FilePaths,
        [string]   $BlobPathPrefix,  # e.g. 'SPO_Configs/SPO_ODFB_Backup_2026-03-27'
        [string]   $ResourceGroup,
        [string]   $SubscriptionId
    )
    $uploadedUrls = @()
    try {
        if ($SubscriptionId) {
            Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
        }
        $getParams = @{ Name = $StorageAccountName; ErrorAction = 'Stop' }
        if ($ResourceGroup) { $getParams['ResourceGroupName'] = $ResourceGroup }
        $storageAccount = Get-AzStorageAccount @getParams
        if (-not $storageAccount) { throw "Storage account '$StorageAccountName' not found." }
        $ctx = $storageAccount.Context

        # Ensure container exists
        $container = Get-AzStorageContainer -Name $ContainerName -Context $ctx -ErrorAction SilentlyContinue
        if (-not $container) {
            if ($PSCmdlet.ShouldProcess($ContainerName, 'Create blob container')) {
                New-AzStorageContainer -Name $ContainerName -Context $ctx | Out-Null
                Write-Log "Created blob container: $ContainerName"
            }
        }

        foreach ($file in $FilePaths) {
            if (-not (Test-Path $file)) {
                Write-Log "File not found, skipping upload: $file" -Level WARN
                continue
            }
            $blobName = "$BlobPathPrefix/$(Split-Path $file -Leaf)"
            if ($PSCmdlet.ShouldProcess($blobName, "Upload to blob container '$ContainerName'")) {
                Set-AzStorageBlobContent -File $file -Container $ContainerName `
                    -Blob $blobName -Context $ctx -Force `
                    -WarningAction SilentlyContinue | Out-Null
                $url = "https://$StorageAccountName.blob.core.windows.net/$ContainerName/$blobName"
                Write-Log "Uploaded: $blobName" -Level SUCCESS
                $uploadedUrls += $url
            }
        }
    }
    catch {
        Write-Log "Azure Storage upload error: $_" -Level ERROR
    }
    return $uploadedUrls
}
#endregion

#region ── Graph Mail ──────────────────────────────────────────────────────────
function Send-BackupSummaryEmail {
    [CmdletBinding()]
    param(
        [string]   $TenantId,
        [string]   $AppId,
        [string]   $CertThumbprint,
        [string]   $FromAddress,
        [string[]] $ToAddresses,
        [string]   $HtmlBody,
        [string]   $Subject
    )
    Write-Log "Connecting to Graph for mail send (SP: $AppId)..."
    try {
        Connect-MgGraph -TenantId $TenantId -ClientId $AppId `
            -CertificateThumbprint $CertThumbprint -NoWelcome -ErrorAction Stop

        $toRecipients = $ToAddresses | ForEach-Object {
            @{ emailAddress = @{ address = $_ } }
        }

        $message = @{
            subject      = $Subject
            body         = @{ contentType = 'HTML'; content = $HtmlBody }
            toRecipients = $toRecipients
        }

        Send-MgUserMail -UserId $FromAddress -Message $message -SaveToSentItems:$false
        Write-Log "Summary email sent to: $($ToAddresses -join ', ')" -Level SUCCESS
    }
    catch {
        Write-Log "Email send failed: $_" -Level ERROR
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
#endregion

#region ── HTML Email Builder ──────────────────────────────────────────────────
function Build-SummaryHtml {
    param([hashtable[]]$Results, [string]$RunTimestamp, [string]$BackupRoot)

    $successCount = @($Results | Where-Object { $_.Status -eq 'Success' }).Count
    $failCount    = @($Results | Where-Object { $_.Status -eq 'Failed'  }).Count
    $totalCount   = @($Results).Count

    $rowsHtml = foreach ($r in $Results) {
        $statusColor = if ($r.Status -eq 'Success') { '#107C10' } else { '#A80000' }
        $statusBadge = "<span style='color:$statusColor;font-weight:bold'>$($r.Status)</span>"
        $blobLinks   = if ($r.BlobUrls) {
            ($r.BlobUrls | ForEach-Object { "<a href='$_'>$(Split-Path $_ -Leaf)</a>" }) -join '<br/>'
        } else { '—' }
        $errors = if ($r.Errors) { "<span style='color:#A80000'>$($r.Errors -join '<br/>')</span>" } else { '—' }

        @"
        <tr>
            <td style='padding:8px;border:1px solid #ddd'>$($r.Workload)</td>
            <td style='padding:8px;border:1px solid #ddd;text-align:center'>$statusBadge</td>
            <td style='padding:8px;border:1px solid #ddd;text-align:center'>$($r.ResourceCount)</td>
            <td style='padding:8px;border:1px solid #ddd;font-size:12px'>$($r.BackupPath)</td>
            <td style='padding:8px;border:1px solid #ddd;font-size:12px'>$blobLinks</td>
            <td style='padding:8px;border:1px solid #ddd;font-size:12px'>$errors</td>
        </tr>
"@
    }

    $overallColor = if ($failCount -eq 0) { '#107C10' } else { '#A80000' }
    $overallStatus = if ($failCount -eq 0) { 'All Succeeded' } else { "$failCount Failed / $successCount Succeeded" }

    return @"
<!DOCTYPE html>
<html>
<head><meta charset='utf-8'/></head>
<body style='font-family:Segoe UI,Arial,sans-serif;background:#f4f4f4;padding:20px'>
<div style='max-width:960px;margin:auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.15)'>

  <!-- Header -->
  <div style='background:#0078D4;padding:24px 32px'>
    <h1 style='margin:0;color:#fff;font-size:22px'>M365DSC Configuration Backup Summary</h1>
    <p style='margin:6px 0 0;color:#cce4f7;font-size:13px'>Run completed: $RunTimestamp</p>
  </div>

  <!-- Stats -->
  <div style='display:flex;gap:0;border-bottom:1px solid #eee'>
    <div style='flex:1;padding:20px 32px;border-right:1px solid #eee'>
      <div style='font-size:32px;font-weight:700;color:#0078D4'>$totalCount</div>
      <div style='font-size:13px;color:#555'>Workloads Processed</div>
    </div>
    <div style='flex:1;padding:20px 32px;border-right:1px solid #eee'>
      <div style='font-size:32px;font-weight:700;color:#107C10'>$successCount</div>
      <div style='font-size:13px;color:#555'>Succeeded</div>
    </div>
    <div style='flex:1;padding:20px 32px;border-right:1px solid #eee'>
      <div style='font-size:32px;font-weight:700;color:#A80000'>$failCount</div>
      <div style='font-size:13px;color:#555'>Failed</div>
    </div>
    <div style='flex:1;padding:20px 32px'>
      <div style='font-size:20px;font-weight:700;color:$overallColor;margin-top:6px'>$overallStatus</div>
      <div style='font-size:13px;color:#555'>Overall Result</div>
    </div>
  </div>

  <!-- Table -->
  <div style='padding:24px 32px'>
    <h2 style='font-size:16px;margin:0 0 12px;color:#333'>Workload Results</h2>
    <table style='width:100%;border-collapse:collapse;font-size:13px'>
      <thead>
        <tr style='background:#f0f0f0'>
          <th style='padding:10px 8px;border:1px solid #ddd;text-align:left'>Workload</th>
          <th style='padding:10px 8px;border:1px solid #ddd;text-align:center'>Status</th>
          <th style='padding:10px 8px;border:1px solid #ddd;text-align:center'>Resources</th>
          <th style='padding:10px 8px;border:1px solid #ddd;text-align:left'>Local Path</th>
          <th style='padding:10px 8px;border:1px solid #ddd;text-align:left'>Azure Blob</th>
          <th style='padding:10px 8px;border:1px solid #ddd;text-align:left'>Errors</th>
        </tr>
      </thead>
      <tbody>
        $($rowsHtml -join '')
      </tbody>
    </table>
  </div>

  <!-- Footer -->
  <div style='padding:16px 32px;background:#f9f9f9;border-top:1px solid #eee;font-size:12px;color:#888'>
    Backup Root: $BackupRoot &nbsp;|&nbsp; Log: $Script:LogFile
  </div>

</div>
</body>
</html>
"@
}
#endregion

#region ── Export Single Workload ──────────────────────────────────────────────
function Invoke-WorkloadBackup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]    $WorkloadName,
        [hashtable] $WorkloadDef,
        [string]    $BackupRoot,
        [string]    $Timestamp,
        [hashtable] $AuthParams,         # ApplicationId, CertificateThumbprint, TenantId
        [string[]]  $OverrideComponents, # from config file
        [string]    $StorageAccountName,
        [string]    $StorageContainerName,
        [string]    $StorageResourceGroup,
        [string]    $StorageSubscriptionId,
        [bool]      $DoUpload,
        [bool]      $DoReport,
        [string]    $ReportsDir
    )

    $result = [ordered]@{
        Workload      = $WorkloadName
        Status        = 'Failed'
        ResourceCount = 0
        BackupPath    = ''
        ReportPath    = ''
        BlobUrls      = @()
        Errors        = @()
        Duration      = ''
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # Resolve components (override wins over workload map default)
        $components = if ($OverrideComponents -and @($OverrideComponents).Count -gt 0) {
            @($OverrideComponents)
        } else {
            @($WorkloadDef.Components)
        }

        Write-Log "[$WorkloadName] Starting export | Components: $($components.Count)"

        # Build folder structure: <BackupRoot>\<FolderPrefix>\<BackupPrefix>_<Timestamp>
        $parentFolder = Join-Path $BackupRoot $WorkloadDef.FolderPrefix
        $backupFolder = Join-Path $parentFolder "$($WorkloadDef.BackupPrefix)_$Timestamp"

        if ($PSCmdlet.ShouldProcess($backupFolder, "Create backup folder")) {
            New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null
        }
        $result.BackupPath = $backupFolder
        Write-Log "[$WorkloadName] Backup folder: $backupFolder"

        # Run the export
        $exportParams = @{
            Components  = $components
            Path        = $backupFolder
            ErrorAction = 'Stop'
        }
        # Merge auth params
        $exportParams += $AuthParams

        if ($PSCmdlet.ShouldProcess($WorkloadName, 'Export-M365DSCConfiguration')) {
            Write-Log "[$WorkloadName] Running Export-M365DSCConfiguration..."
            Export-M365DSCConfiguration @exportParams
            Write-Log "[$WorkloadName] Export complete." -Level SUCCESS
        }

        # Count exported PS1 resources
        $ps1File = Get-ChildItem -Path $backupFolder -Filter '*.ps1' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($ps1File) {
            # Rough resource count: count DSC resource blocks
            $content = Get-Content $ps1File.FullName -Raw
            $result.ResourceCount = ([regex]::Matches($content, '^\s*\w+\s+"[^"]+"')).Count
        }

        # Generate HTML report
        if ($DoReport -and $ps1File) {
            Write-Log "[$WorkloadName] Generating HTML report..."
            $reportFileName = "$($WorkloadDef.BackupPrefix)_Report_$Timestamp.html"
            $reportPath     = Join-Path $ReportsDir $reportFileName

            New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null

            try {
                New-M365DSCReportFromConfiguration `
                    -Type            'HTML' `
                    -ConfigurationPath $ps1File.FullName `
                    -OutputPath      $reportPath `
                    -ErrorAction     Stop
                $result.ReportPath = $reportPath
                Write-Log "[$WorkloadName] Report saved: $reportPath" -Level SUCCESS
            }
            catch {
                $errMsg = "Report generation failed: $_"
                Write-Log "[$WorkloadName] $errMsg" -Level WARN
                $result.Errors += $errMsg
            }
        }

        # Upload to Azure Blob then purge local files
        if ($DoUpload) {
            Write-Log "[$WorkloadName] Uploading to Azure Blob..."
            $blobPrefix    = "$($WorkloadDef.FolderPrefix)/$($WorkloadDef.BackupPrefix)_$Timestamp"
            $filesToUpload = @(Get-ChildItem -Path $backupFolder -File | Select-Object -ExpandProperty FullName)
            if ($result.ReportPath -and (Test-Path $result.ReportPath)) {
                $filesToUpload += $result.ReportPath
            }

            $urls = @(Send-ToAzureBlob `
                -StorageAccountName $StorageAccountName `
                -ContainerName      $StorageContainerName `
                -FilePaths          $filesToUpload `
                -BlobPathPrefix     $blobPrefix `
                -ResourceGroup      $StorageResourceGroup `
                -SubscriptionId     $StorageSubscriptionId)
            $result.BlobUrls = $urls

            # Purge local files only if every file was successfully uploaded
            # (guard: uploaded URL count must match files attempted)
            $uploadedCount = @($result.BlobUrls).Count
            $attemptedCount = @($filesToUpload).Count
            if ($uploadedCount -gt 0 -and $uploadedCount -eq $attemptedCount) {
                Write-Log "[$WorkloadName] Upload complete ($uploadedCount/$attemptedCount files). Purging local copies..."
                try {
                    if ($PSCmdlet.ShouldProcess($backupFolder, 'Remove local backup folder')) {
                        Remove-Item -Path $backupFolder -Recurse -Force
                        Write-Log "[$WorkloadName] Purged backup folder: $backupFolder" -Level SUCCESS
                    }
                    if ($result.ReportPath -and (Test-Path $result.ReportPath)) {
                        if ($PSCmdlet.ShouldProcess($result.ReportPath, 'Remove local report file')) {
                            Remove-Item -Path $result.ReportPath -Force
                            Write-Log "[$WorkloadName] Purged report file: $result.ReportPath" -Level SUCCESS
                        }
                    }
                }
                catch {
                    # Purge failure is non-fatal — backup is safely in blob storage
                    Write-Log "[$WorkloadName] Warning: local purge failed: $_" -Level WARN
                }
            }
            elseif ($uploadedCount -lt $attemptedCount) {
                Write-Log "[$WorkloadName] Upload incomplete ($uploadedCount/$attemptedCount files uploaded) — local files retained." -Level WARN
            }
        }

        $result.Status = 'Success'
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log "[$WorkloadName] FAILED: $errMsg" -Level ERROR
        $result.Errors += $errMsg
    }
    finally {
        $stopwatch.Stop()
        $result.Duration = $stopwatch.Elapsed.ToString('mm\:ss')
        Write-Log "[$WorkloadName] Duration: $($result.Duration)"
    }

    return $result
}
#endregion

#region ── Main ────────────────────────────────────────────────────────────────
function Main {
    # ── Init log ──────────────────────────────────────────────────────────────
    $Script:RunTimestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $null = New-Item -ItemType Directory -Path $LogPath -Force
    $Script:LogFile = Join-Path $LogPath "M365DSCBackup_$($Script:RunTimestamp).log"
    Write-Log "=== M365DSC Backup Run Started: $Script:RunTimestamp ==="

    # ── Load config file ──────────────────────────────────────────────────────
    $cfg = $null
    if ($ConfigFile) {
        if (Test-Path $ConfigFile) {
            $cfg = Import-BackupConfig -Path $ConfigFile
        } else {
            Write-Log "Config file not found at default location: $ConfigFile — proceeding with parameters only." -Level WARN
        }
    }

    # ── Resolve effective settings (param > config > default) ─────────────────
    $cfgGlobal  = if ($cfg) { $cfg.Global  } else { $null }
    $cfgStorage = if ($cfg) { $cfg.Storage } else { $null }
    $cfgEmail   = if ($cfg) { $cfg.Email   } else { $null }

    $cfgBackupRoot            = if ($cfgGlobal)  { $cfgGlobal.BackupRoot              } else { $null }
    $cfgTenantId              = if ($cfgGlobal)  { $cfgGlobal.TenantId                } else { $null }
    $cfgTenantName            = if ($cfgGlobal)  { $cfgGlobal.TenantName              } else { $null }
    $cfgStorAcct              = if ($cfgStorage) { $cfgStorage.AccountName            } else { $null }
    $cfgStorContainer         = if ($cfgStorage) { $cfgStorage.ContainerName          } else { $null }
    $cfgStorRG                = if ($cfgStorage) { $cfgStorage.ResourceGroup          } else { $null }
    $cfgStorSubId             = if ($cfgStorage) { $cfgStorage.SubscriptionId         } else { $null }
    $cfgStorSPAppId           = if ($cfgStorage) { $cfgStorage.SPAppId                } else { $null }
    $cfgStorSPCert            = if ($cfgStorage) { $cfgStorage.SPCertThumbprint       } else { $null }
    $cfgMIClientId            = if ($cfgStorage) { $cfgStorage.ManagedIdentityClientId } else { $null }
    $cfgEmailFrom             = if ($cfgEmail)   { $cfgEmail.From                     } else { $null }
    $cfgEmailAppId            = if ($cfgEmail)   { $cfgEmail.AppId                    } else { $null }
    $cfgEmailCert             = if ($cfgEmail)   { $cfgEmail.CertThumbprint           } else { $null }

    $effectiveBackupRoot    = Resolve-Setting $BackupRoot             $cfgBackupRoot
    $effectiveTenantId      = Resolve-Setting $TenantId              $cfgTenantId
    $effectiveTenantName    = Resolve-Setting $TenantName            $cfgTenantName
    $effectiveWorkloads     = if ($Workloads) { [string[]]@($Workloads) } elseif ($cfgGlobal -and $cfgGlobal.DefaultWorkloads) { [string[]]$cfgGlobal.DefaultWorkloads } else { @() }
    $effectiveStorAcct      = Resolve-Setting $StorageAccountName    $cfgStorAcct
    $effectiveStorContainer = Resolve-Setting $StorageContainerName  $cfgStorContainer  'm365dsc-backups'
    $effectiveStorRG        = Resolve-Setting $StorageResourceGroup  $cfgStorRG
    $effectiveStorSubId     = Resolve-Setting $StorageSubscriptionId $cfgStorSubId
    $effectiveStorSPAppId   = Resolve-Setting $StorageSPAppId        $cfgStorSPAppId
    $effectiveStorSPCert    = Resolve-Setting $StorageSPCertThumbprint $cfgStorSPCert
    $effectiveUseMI         = $UseManagedIdentity.IsPresent -or ($cfgStorage -and [bool]$cfgStorage.UseManagedIdentity)
    $effectiveMIClientId    = Resolve-Setting $ManagedIdentityClientId $cfgMIClientId
    $effectiveEmailFrom     = Resolve-Setting $EmailFrom             $cfgEmailFrom
    $effectiveEmailTo       = if ($EmailTo) { $EmailTo } elseif ($cfgEmail -and $cfgEmail.To) { $cfgEmail.To } else { @() }
    $effectiveEmailSPAppId  = Resolve-Setting $EmailSPAppId          $cfgEmailAppId
    $effectiveEmailSPCert   = Resolve-Setting $EmailSPCertThumbprint $cfgEmailCert
    $doUpload              = -not $SkipUpload.IsPresent
    $doReport              = -not $SkipReport.IsPresent
    $doEmail               = if ($SkipEmail.IsPresent) { $false } elseif ($cfgEmail -and $null -ne $cfgEmail.Enabled) { [bool]$cfgEmail.Enabled } else { $true }

    # ── Validate required settings ────────────────────────────────────────────
    if (-not $effectiveBackupRoot)   { throw 'BackupRoot is required.' }
    if (-not $effectiveTenantId)     { throw 'TenantId is required (GUID format — used for Azure Storage and Graph auth).' }
    if (-not $effectiveTenantName)   { throw 'TenantName is required (domain format, e.g. contoso.onmicrosoft.com — used for M365DSC exports).' }
    if ($effectiveWorkloads.Count -eq 0) { throw 'No workloads specified. Use -Workloads or set DefaultWorkloads in config. Valid values: Entra, Teams, Exchange, Purview, SharePoint, OneDrive, Intune, Defender, O365Services' }

    # Validate workload names
    foreach ($wl in $effectiveWorkloads) {
        if (-not $Script:WorkloadMap.ContainsKey($wl)) {
            throw "Unknown workload '$wl'. Valid values: $($Script:WorkloadMap.Keys -join ', ')"
        }
    }

    # ── Connect to Azure Storage (if uploading) ───────────────────────────────
    if ($doUpload) {
        if (-not $effectiveStorAcct) {
            Write-Log "StorageAccountName not set — disabling upload for this run." -Level WARN
            $doUpload = $false
        } elseif ($effectiveUseMI) {
            Connect-AzureStorage `
                -TenantId                $effectiveTenantId `
                -SubscriptionId          $effectiveStorSubId `
                -UseManagedIdentity `
                -ManagedIdentityClientId $effectiveMIClientId
        } elseif (-not $effectiveStorSPAppId -or -not $effectiveStorSPCert) {
            Write-Log "Azure Storage parameters incomplete — disabling upload for this run." -Level WARN
            $doUpload = $false
        } else {
            Connect-AzureStorage `
                -TenantId       $effectiveTenantId `
                -AppId          $effectiveStorSPAppId `
                -CertThumbprint $effectiveStorSPCert `
                -SubscriptionId $effectiveStorSubId
        }
    }

    # ── Reports directory ─────────────────────────────────────────────────────
    $reportsDir = Join-Path $effectiveBackupRoot 'Reports'

    # ── Process each workload ─────────────────────────────────────────────────
    $allResults = @()

    foreach ($wl in $effectiveWorkloads) {
        Write-Log "──────────────────────────────────────────"
        Write-Log "Processing workload: $wl"

        $workloadDef = $Script:WorkloadMap[$wl]

        # Get per-workload auth from config (or fall back to global)
        $wlCfg = $null
        if ($cfg -and $cfg.Workloads) {
            $wlCfg = $cfg.Workloads | Where-Object { $_.Name -eq $wl } | Select-Object -First 1
        }

        $authAppId     = if ($wlCfg -and $wlCfg.ApplicationId)        { $wlCfg.ApplicationId }        else { if ($cfgGlobal) { $cfgGlobal.DefaultApplicationId } else { $null } }
        $authCert      = if ($wlCfg -and $wlCfg.CertificateThumbprint) { $wlCfg.CertificateThumbprint } else { if ($cfgGlobal) { $cfgGlobal.DefaultCertThumbprint } else { $null } }
        $overrideComps = if ($wlCfg -and $wlCfg.Components)            { @($wlCfg.Components) }         else { @() }

        if (-not $authAppId -or -not $authCert) {
            Write-Log "[$wl] Missing ApplicationId or CertificateThumbprint — skipping." -Level ERROR
            $allResults += [ordered]@{
                Workload      = $wl
                Status        = 'Failed'
                ResourceCount = 0
                BackupPath    = ''
                ReportPath    = ''
                BlobUrls      = @()
                Errors        = @('Missing service principal credentials in config.')
                Duration      = '00:00'
            }
            continue
        }

        $authParams = @{
            ApplicationId          = $authAppId
            CertificateThumbprint  = $authCert
            TenantId               = $effectiveTenantName
        }

        $result = Invoke-WorkloadBackup `
            -WorkloadName         $wl `
            -WorkloadDef          $workloadDef `
            -BackupRoot           $effectiveBackupRoot `
            -Timestamp            $Script:RunTimestamp `
            -AuthParams           $authParams `
            -OverrideComponents   $overrideComps `
            -StorageAccountName   $effectiveStorAcct `
            -StorageContainerName $effectiveStorContainer `
            -StorageResourceGroup $effectiveStorRG `
            -StorageSubscriptionId $effectiveStorSubId `
            -DoUpload             $doUpload `
            -DoReport             $doReport `
            -ReportsDir           $reportsDir

        $allResults += $result
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    Write-Log "──────────────────────────────────────────"
    $workloadResults = @($allResults | Where-Object { $_ -is [System.Collections.Specialized.OrderedDictionary] })
    $success = @($workloadResults | Where-Object { $_.Status -eq 'Success' }).Count
    $failed  = @($workloadResults | Where-Object { $_.Status -eq 'Failed'  }).Count
    Write-Log "Run complete | Success: $success | Failed: $failed" -Level $(if ($failed -eq 0) { 'SUCCESS' } else { 'WARN' })

    # ── Send email ────────────────────────────────────────────────────────────
    if ($doEmail) {
        if (-not $effectiveEmailFrom -or @($effectiveEmailTo).Count -eq 0 -or -not $effectiveEmailSPAppId -or -not $effectiveEmailSPCert) {
            Write-Log "Email parameters incomplete (From/To/AppId/CertThumbprint required) — skipping email." -Level WARN
        } else {
            $htmlBody = Build-SummaryHtml `
                -Results       $workloadResults `
                -RunTimestamp  $Script:RunTimestamp `
                -BackupRoot    $effectiveBackupRoot

            $subject = "M365DSC Backup - $Script:RunTimestamp `| $success/$($workloadResults.Count) Succeeded"

            Send-BackupSummaryEmail `
                -TenantId        $effectiveTenantId `
                -AppId           $effectiveEmailSPAppId `
                -CertThumbprint  $effectiveEmailSPCert `
                -FromAddress     $effectiveEmailFrom `
                -ToAddresses     $effectiveEmailTo `
                -HtmlBody        $htmlBody `
                -Subject         $subject
        }
    }

    Write-Log "=== M365DSC Backup Run Finished ==="

    # Return results for pipeline / batch tool consumption
    return $allResults
}

# Entry point
Main
#endregion

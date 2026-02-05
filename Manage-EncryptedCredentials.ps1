<#
.SYNOPSIS
    Creates and manages encrypted credential files for secure script authentication.

.DESCRIPTION
    This script provides a user-friendly interface to create, test, update, and delete
    encrypted credential configuration files. Credentials are encrypted using Windows DPAPI
    and can only be decrypted by the same user account on the same machine.
    
    Ideal for storing Azure app registration credentials, service account passwords,
    and other sensitive authentication data needed by automation scripts.

.PARAMETER ConfigPath
    Full path where the encrypted configuration file will be stored.
    Default: C:\Config\Credentials\

.PARAMETER ConfigName
    Name for the credential configuration (will be saved as ConfigName.xml)

.PARAMETER Action
    Action to perform: Create, Test, Update, Delete, List

.EXAMPLE
    .\Manage-EncryptedCredentials.ps1 -Action Create -ConfigName "GraphAPI"
    Creates a new encrypted credential file for Microsoft Graph API authentication.

.EXAMPLE
    .\Manage-EncryptedCredentials.ps1 -Action List
    Lists all encrypted credential files in the default directory.

.EXAMPLE
    .\Manage-EncryptedCredentials.ps1 -Action Test -ConfigName "GraphAPI"
    Tests the credential file to ensure it can be decrypted properly.

.NOTES
    File Name      : Manage-EncryptedCredentials.ps1
    Author         : IT Infrastructure Team
    Prerequisite   : PowerShell 5.1 or higher
    
    SECURITY NOTES:
    - Credentials are encrypted using DPAPI (Data Protection API)
    - Can only be decrypted by the SAME USER on the SAME MACHINE
    - For scripts running as SYSTEM, create credentials while running as SYSTEM
    - Store credential files in a secure location with restricted NTFS permissions
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Container)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
        $true
    })]
    [string]$ConfigPath = "C:\Config\Credentials",

    [Parameter()]
    [string]$ConfigName,

    [Parameter(Mandatory)]
    [ValidateSet('Create', 'Test', 'Update', 'Delete', 'List')]
    [string]$Action
)

#region Helper Functions

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Get-SecureInput {
    param(
        [string]$Prompt
    )
    Read-Host $Prompt -AsSecureString
}

function New-EncryptedCredentialFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        
        [Parameter(Mandatory)]
        [string]$Name,
        
        [Parameter(Mandatory)]
        [string]$Description,
        
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )
    
    try {
        # Build configuration object
        $Config = @{
            Name = $Name
            Description = $Description
            CreatedBy = $env:USERNAME
            CreatedOn = Get-Date
            MachineName = $env:COMPUTERNAME
            LastModified = Get-Date
            Credentials = @{}
        }
        
        # Encrypt each credential
        foreach ($Key in $Credentials.Keys) {
            if ($Credentials[$Key] -is [SecureString]) {
                $Config.Credentials[$Key] = $Credentials[$Key] | ConvertFrom-SecureString
            }
            else {
                $Config.Credentials[$Key] = $Credentials[$Key]
            }
        }
        
        # Export to XML
        $Config | Export-Clixml -Path $FilePath -Force
        
        # Set restrictive NTFS permissions
        $Acl = Get-Acl $FilePath
        $Acl.SetAccessRuleProtection($true, $false)  # Disable inheritance
        
        # Add Administrators
        $AdminRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            "BUILTIN\Administrators", 
            "FullControl", 
            "Allow"
        )
        
        # Add SYSTEM
        $SystemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            "NT AUTHORITY\SYSTEM", 
            "FullControl", 
            "Allow"
        )
        
        # Add current user
        $UserRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $env:USERNAME,
            "FullControl",
            "Allow"
        )
        
        $Acl.SetAccessRule($AdminRule)
        $Acl.SetAccessRule($SystemRule)
        $Acl.SetAccessRule($UserRule)
        Set-Acl -Path $FilePath -AclObject $Acl
        
        Write-ColorOutput "✓ Credential file created successfully: $FilePath" -Color Green
        Write-ColorOutput "  Created by: $($Config.CreatedBy) on $($Config.MachineName)" -Color Gray
        Write-ColorOutput "  Security: File permissions restricted to Administrators and SYSTEM" -Color Gray
        
        return $true
    }
    catch {
        Write-ColorOutput "✗ Failed to create credential file: $($_.Exception.Message)" -Color Red
        return $false
    }
}

function Import-EncryptedCredentialFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            throw "Credential file not found: $FilePath"
        }
        
        $Config = Import-Clixml -Path $FilePath
        
        # Decrypt credentials
        $DecryptedCreds = @{}
        foreach ($Key in $Config.Credentials.Keys) {
            if ($Config.Credentials[$Key] -match '^[0-9a-f]+$') {
                # This is an encrypted string
                try {
                    $DecryptedCreds[$Key] = $Config.Credentials[$Key] | ConvertTo-SecureString
                }
                catch {
                    throw "Failed to decrypt credential '$Key'. Ensure you're running as the user who created this file ($($Config.CreatedBy)) on the correct machine ($($Config.MachineName))."
                }
            }
            else {
                # Plain text value (like TenantId, AppId)
                $DecryptedCreds[$Key] = $Config.Credentials[$Key]
            }
        }
        
        $Config.Credentials = $DecryptedCreds
        return $Config
    }
    catch {
        Write-ColorOutput "✗ Failed to import credential file: $($_.Exception.Message)" -Color Red
        throw
    }
}

#endregion

#region Main Actions

function Invoke-CreateCredential {
    param([string]$ConfigPath, [string]$ConfigName)
    
    Write-ColorOutput "`n=== Create New Encrypted Credential ===" -Color Cyan
    Write-ColorOutput "This will create an encrypted credential file that can only be decrypted by:" -Color Yellow
    Write-ColorOutput "  User: $env:USERNAME" -Color Yellow
    Write-ColorOutput "  Machine: $env:COMPUTERNAME" -Color Yellow
    Write-Host ""
    
    # Get configuration name if not provided
    if (-not $ConfigName) {
        $ConfigName = Read-Host "Enter a name for this credential (e.g., GraphAPI, ExchangeOnline, ServiceAccount)"
    }
    
    $FilePath = Join-Path $ConfigPath "$ConfigName.xml"
    
    if (Test-Path $FilePath) {
        $Overwrite = Read-Host "Configuration '$ConfigName' already exists. Overwrite? (Y/N)"
        if ($Overwrite -ne 'Y') {
            Write-ColorOutput "Operation cancelled." -Color Yellow
            return
        }
    }
    
    # Get description
    $Description = Read-Host "Enter a description (what is this credential for?)"
    
    # Determine credential type
    Write-Host "`nSelect credential type:" -ForegroundColor Cyan
    Write-Host "  1. Azure App Registration (TenantId, AppId, ClientSecret)"
    Write-Host "  2. Service Account (Username, Password)"
    Write-Host "  3. API Key (single secret value)"
    Write-Host "  4. Custom (you specify the fields)"
    $CredType = Read-Host "Enter selection (1-4)"
    
    $Credentials = @{}
    
    switch ($CredType) {
        '1' {
            # Azure App Registration
            $Credentials['TenantId'] = Read-Host "Enter Tenant ID"
            $Credentials['AppId'] = Read-Host "Enter Application (Client) ID"
            $Credentials['ClientSecret'] = Get-SecureInput "Enter Client Secret"
            
            # Optional: Add certificate thumbprint
            $AddCert = Read-Host "Do you also want to store a certificate thumbprint? (Y/N)"
            if ($AddCert -eq 'Y') {
                $Credentials['CertificateThumbprint'] = Read-Host "Enter Certificate Thumbprint"
            }
        }
        '2' {
            # Service Account
            $Credentials['Username'] = Read-Host "Enter Username"
            $Credentials['Password'] = Get-SecureInput "Enter Password"
            $Credentials['Domain'] = Read-Host "Enter Domain (optional, press Enter to skip)"
        }
        '3' {
            # API Key
            $Credentials['ApiKey'] = Get-SecureInput "Enter API Key"
            $Credentials['ApiUrl'] = Read-Host "Enter API URL (optional)"
        }
        '4' {
            # Custom
            Write-ColorOutput "`nEnter custom credential fields (press Enter with empty name when done):" -Color Cyan
            do {
                $FieldName = Read-Host "Field name"
                if ($FieldName) {
                    $IsSecret = Read-Host "Is this a secret/password? (Y/N)"
                    if ($IsSecret -eq 'Y') {
                        $Credentials[$FieldName] = Get-SecureInput "Enter value for $FieldName"
                    }
                    else {
                        $Credentials[$FieldName] = Read-Host "Enter value for $FieldName"
                    }
                }
            } while ($FieldName)
        }
    }
    
    # Create the file
    $Success = New-EncryptedCredentialFile -FilePath $FilePath -Name $ConfigName -Description $Description -Credentials $Credentials
    
    if ($Success) {
        Write-Host "`n" -NoNewline
        Write-ColorOutput "Credential file created successfully!" -Color Green
        Write-ColorOutput "Location: $FilePath" -Color White
        Write-Host "`nTo use in your scripts, add this code:" -ForegroundColor Cyan
        Write-Host @"

# Import encrypted credentials
`$Config = Import-Clixml -Path "$FilePath"
`$Creds = @{}
foreach (`$Key in `$Config.Credentials.Keys) {
    if (`$Config.Credentials[`$Key] -is [string] -and `$Config.Credentials[`$Key] -match '^[0-9a-f]+$') {
        `$Creds[`$Key] = `$Config.Credentials[`$Key] | ConvertTo-SecureString
    } else {
        `$Creds[`$Key] = `$Config.Credentials[`$Key]
    }
}

"@ -ForegroundColor Gray
    }
}

function Invoke-TestCredential {
    param([string]$ConfigPath, [string]$ConfigName)
    
    if (-not $ConfigName) {
        $ConfigName = Read-Host "Enter credential name to test"
    }
    
    $FilePath = Join-Path $ConfigPath "$ConfigName.xml"
    
    Write-ColorOutput "`n=== Testing Credential: $ConfigName ===" -Color Cyan
    
    try {
        $Config = Import-EncryptedCredentialFile -FilePath $FilePath
        
        Write-ColorOutput "✓ Credential file decrypted successfully!" -Color Green
        Write-Host "`nConfiguration Details:" -ForegroundColor White
        Write-Host "  Name: $($Config.Name)" -ForegroundColor Gray
        Write-Host "  Description: $($Config.Description)" -ForegroundColor Gray
        Write-Host "  Created By: $($Config.CreatedBy) on $($Config.MachineName)" -ForegroundColor Gray
        Write-Host "  Created On: $($Config.CreatedOn)" -ForegroundColor Gray
        Write-Host "  Last Modified: $($Config.LastModified)" -ForegroundColor Gray
        
        Write-Host "`nAvailable Credential Fields:" -ForegroundColor White
        foreach ($Key in $Config.Credentials.Keys) {
            $Type = if ($Config.Credentials[$Key] -is [SecureString]) { "(Encrypted)" } else { "(Plain Text)" }
            Write-Host "  - $Key $Type" -ForegroundColor Gray
        }
        
        # Test decryption of secure strings
        $SecureFields = $Config.Credentials.Keys | Where-Object { $Config.Credentials[$_] -is [SecureString] }
        if ($SecureFields) {
            Write-Host "`nTesting secure field decryption..." -ForegroundColor White
            foreach ($Field in $SecureFields) {
                try {
                    $PlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Config.Credentials[$Field])
                    )
                    $Length = $PlainText.Length
                    Write-Host "  ✓ $Field : $Length characters (decrypted successfully)" -ForegroundColor Green
                }
                catch {
                    Write-Host "  ✗ $Field : Failed to decrypt" -ForegroundColor Red
                }
            }
        }
    }
    catch {
        Write-ColorOutput "✗ Test failed: $($_.Exception.Message)" -Color Red
    }
}

function Invoke-ListCredentials {
    param([string]$ConfigPath)
    
    Write-ColorOutput "`n=== Encrypted Credential Files ===" -Color Cyan
    Write-ColorOutput "Location: $ConfigPath`n" -Color Gray
    
    $Files = Get-ChildItem -Path $ConfigPath -Filter "*.xml" -ErrorAction SilentlyContinue
    
    if (-not $Files) {
        Write-ColorOutput "No credential files found." -Color Yellow
        return
    }
    
    $Results = @()
    foreach ($File in $Files) {
        try {
            $Config = Import-Clixml -Path $File.FullName
            $CanDecrypt = $true
            
            # Test if we can decrypt
            try {
                foreach ($Key in $Config.Credentials.Keys) {
                    if ($Config.Credentials[$Key] -match '^[0-9a-f]+$') {
                        $null = $Config.Credentials[$Key] | ConvertTo-SecureString
                    }
                }
            }
            catch {
                $CanDecrypt = $false
            }
            
            $Results += [PSCustomObject]@{
                Name = $Config.Name
                Description = $Config.Description
                CreatedBy = $Config.CreatedBy
                CreatedOn = $Config.CreatedOn
                Machine = $Config.MachineName
                CanDecrypt = $CanDecrypt
                FilePath = $File.FullName
            }
        }
        catch {
            $Results += [PSCustomObject]@{
                Name = $File.BaseName
                Description = "Error reading file"
                CreatedBy = "Unknown"
                CreatedOn = $File.CreationTime
                Machine = "Unknown"
                CanDecrypt = $false
                FilePath = $File.FullName
            }
        }
    }
    
    $Results | Format-Table -Property Name, Description, CreatedBy, Machine, CanDecrypt -AutoSize
    
    Write-ColorOutput "`nTotal: $($Files.Count) credential file(s)" -Color Gray
    Write-Host "Note: 'CanDecrypt' indicates if you can decrypt the file with your current user context." -ForegroundColor Yellow
}

function Invoke-UpdateCredential {
    param([string]$ConfigPath, [string]$ConfigName)
    
    if (-not $ConfigName) {
        $ConfigName = Read-Host "Enter credential name to update"
    }
    
    $FilePath = Join-Path $ConfigPath "$ConfigName.xml"
    
    if (-not (Test-Path $FilePath)) {
        Write-ColorOutput "✗ Credential '$ConfigName' not found." -Color Red
        return
    }
    
    Write-ColorOutput "`n=== Update Credential: $ConfigName ===" -Color Cyan
    
    try {
        $Config = Import-EncryptedCredentialFile -FilePath $FilePath
        
        Write-Host "Current fields:" -ForegroundColor White
        foreach ($Key in $Config.Credentials.Keys) {
            $Type = if ($Config.Credentials[$Key] -is [SecureString]) { "(Encrypted)" } else { "(Plain Text)" }
            Write-Host "  - $Key $Type" -ForegroundColor Gray
        }
        
        Write-Host "`nEnter the field name to update (or press Enter to cancel):" -ForegroundColor Cyan
        $FieldToUpdate = Read-Host "Field name"
        
        if (-not $FieldToUpdate) {
            Write-ColorOutput "Update cancelled." -Color Yellow
            return
        }
        
        if (-not $Config.Credentials.ContainsKey($FieldToUpdate)) {
            Write-ColorOutput "Field '$FieldToUpdate' not found in this credential." -Color Red
            return
        }
        
        $IsSecret = Read-Host "Is this a secret/password? (Y/N)"
        if ($IsSecret -eq 'Y') {
            $NewValue = Get-SecureInput "Enter new value for $FieldToUpdate"
            $Config.Credentials[$FieldToUpdate] = $NewValue | ConvertFrom-SecureString
        }
        else {
            $NewValue = Read-Host "Enter new value for $FieldToUpdate"
            $Config.Credentials[$FieldToUpdate] = $NewValue
        }
        
        $Config.LastModified = Get-Date
        $Config | Export-Clixml -Path $FilePath -Force
        
        Write-ColorOutput "✓ Credential updated successfully!" -Color Green
    }
    catch {
        Write-ColorOutput "✗ Update failed: $($_.Exception.Message)" -Color Red
    }
}

function Invoke-DeleteCredential {
    param([string]$ConfigPath, [string]$ConfigName)
    
    if (-not $ConfigName) {
        $ConfigName = Read-Host "Enter credential name to delete"
    }
    
    $FilePath = Join-Path $ConfigPath "$ConfigName.xml"
    
    if (-not (Test-Path $FilePath)) {
        Write-ColorOutput "✗ Credential '$ConfigName' not found." -Color Red
        return
    }
    
    Write-ColorOutput "`n=== Delete Credential: $ConfigName ===" -Color Cyan
    Write-ColorOutput "WARNING: This action cannot be undone!" -Color Red
    
    $Confirm = Read-Host "Are you sure you want to delete this credential? Type 'DELETE' to confirm"
    
    if ($Confirm -eq 'DELETE') {
        try {
            Remove-Item -Path $FilePath -Force
            Write-ColorOutput "✓ Credential deleted successfully." -Color Green
        }
        catch {
            Write-ColorOutput "✗ Delete failed: $($_.Exception.Message)" -Color Red
        }
    }
    else {
        Write-ColorOutput "Delete cancelled." -Color Yellow
    }
}

#endregion

#region Main Execution

try {
    # Ensure config directory exists
    if (-not (Test-Path $ConfigPath)) {
        New-Item -Path $ConfigPath -ItemType Directory -Force | Out-Null
        Write-ColorOutput "Created credential directory: $ConfigPath" -Color Green
    }
    
    # Execute action
    switch ($Action) {
        'Create' { Invoke-CreateCredential -ConfigPath $ConfigPath -ConfigName $ConfigName }
        'Test'   { Invoke-TestCredential -ConfigPath $ConfigPath -ConfigName $ConfigName }
        'Update' { Invoke-UpdateCredential -ConfigPath $ConfigPath -ConfigName $ConfigName }
        'Delete' { Invoke-DeleteCredential -ConfigPath $ConfigPath -ConfigName $ConfigName }
        'List'   { Invoke-ListCredentials -ConfigPath $ConfigPath }
    }
}
catch {
    Write-ColorOutput "`n✗ Error: $($_.Exception.Message)" -Color Red
    Write-ColorOutput $_.ScriptStackTrace -Color Gray
    exit 1
}

#endregion
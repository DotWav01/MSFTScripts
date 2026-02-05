#region Encrypted Credential Import
<#
    ENCRYPTED CREDENTIAL IMPORT TEMPLATE
    
    Copy this section into any script that needs to use encrypted credentials.
    
    Requirements:
    1. Credential file must be created using Manage-EncryptedCredentials.ps1
    2. Script must run as the SAME USER who created the credential file
    3. Script must run on the SAME MACHINE where credential was created
    
    To create a credential file, run:
    .\Manage-EncryptedCredentials.ps1 -Action Create -ConfigName "YourCredName"
#>

function Import-EncryptedCredential {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CredentialName,
        
        [Parameter()]
        [string]$CredentialPath = "C:\Config\Credentials"
    )
    
    $FilePath = Join-Path $CredentialPath "$CredentialName.xml"
    
    try {
        if (-not (Test-Path $FilePath)) {
            throw "Credential file not found: $FilePath"
        }
        
        Write-Verbose "Loading encrypted credential: $CredentialName"
        $Config = Import-Clixml -Path $FilePath
        
        # Decrypt credentials
        $DecryptedCreds = @{}
        foreach ($Key in $Config.Credentials.Keys) {
            if ($Config.Credentials[$Key] -is [string] -and $Config.Credentials[$Key] -match '^[0-9a-f]+$') {
                # Encrypted SecureString
                try {
                    $DecryptedCreds[$Key] = $Config.Credentials[$Key] | ConvertTo-SecureString -ErrorAction Stop
                }
                catch {
                    throw "Failed to decrypt credential field '$Key'. Ensure you're running as user '$($Config.CreatedBy)' on machine '$($Config.MachineName)'."
                }
            }
            else {
                # Plain text value
                $DecryptedCreds[$Key] = $Config.Credentials[$Key]
            }
        }
        
        Write-Verbose "Successfully loaded $($DecryptedCreds.Count) credential fields"
        return $DecryptedCreds
        
    }
    catch {
        Write-Error "Failed to import encrypted credential '$CredentialName': $($_.Exception.Message)"
        throw
    }
}

# Quick helper to convert SecureString to plain text (use sparingly!)
function ConvertFrom-SecureStringToPlainText {
    param([SecureString]$SecureString)
    
    try {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }
    finally {
        if ($BSTR) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        }
    }
}
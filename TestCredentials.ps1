#Requires -Version 5.1

#region Import Encrypted Credentials
function Import-EncryptedCredential {
    param([string]$CredentialName, [string]$CredentialPath = "C:\Config\Credentials")
    $FilePath = Join-Path $CredentialPath "$CredentialName.xml"
    if (-not (Test-Path $FilePath)) { throw "Credential file not found: $FilePath" }
    $Config = Import-Clixml -Path $FilePath
    $Creds = @{}
    foreach ($Key in $Config.Credentials.Keys) {
        if ($Config.Credentials[$Key] -is [string] -and $Config.Credentials[$Key] -match '^[0-9a-f]+$') {
            $Creds[$Key] = $Config.Credentials[$Key] | ConvertTo-SecureString
        } else { $Creds[$Key] = $Config.Credentials[$Key] }
    }
    return $Creds
}

function ConvertFrom-SecureStringToPlainText {
    param([SecureString]$SecureString)
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try { return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR) }
}
#endregion

# === YOUR SCRIPT STARTS HERE ===

try {
    # Load credentials
    $Creds = Import-EncryptedCredential -CredentialName "TestAppReg"
    
    # Connect to Microsoft Graph
    $ClientSecretCredential = New-Object PSCredential($Creds.AppId, $Creds.ClientSecret)
    Connect-MgGraph -TenantId $Creds.TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome
    
    Write-Host "✓ Connected to Microsoft Graph successfully" -ForegroundColor Green
    
    # Your code here...
    
}
catch {
    Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    # Cleanup
    if (Get-MgContext) {
        Disconnect-MgGraph | Out-Null
    }
}
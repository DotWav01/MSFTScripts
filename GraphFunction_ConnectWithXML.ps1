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

$Creds = Import-EncryptedCredential -CredentialName "TestAppReg"
$ClientSecretCredential = New-Object PSCredential($Creds.AppId, $Creds.ClientSecret)
Connect-MgGraph -TenantId $Creds.TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome
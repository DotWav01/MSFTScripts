<#
.SYNOPSIS
    Creates an encrypted credential file for app registration authentication.

.DESCRIPTION
    This helper script securely encrypts and stores app registration client secret
    using Windows DPAPI. The encrypted file can only be decrypted by the same user
    on the same machine.

.PARAMETER OutputPath
    Path where the encrypted credential file will be saved.

.PARAMETER ClientSecret
    The client secret from your app registration (will be prompted securely if not provided).

.EXAMPLE
    .\New-EncryptedAppCredential.ps1 -OutputPath "C:\Scripts\Config\AppCred_GraphReports.xml"
    
    Prompts for client secret and creates encrypted credential file.

.EXAMPLE
    $secret = Read-Host "Enter secret" -AsSecureString
    .\New-EncryptedAppCredential.ps1 -OutputPath "C:\Scripts\Config\AppCred_UserGroup.xml" -ClientSecret $secret
    
    Creates encrypted credential file using provided SecureString.

.NOTES
    Author: Alexander
    Version: 1.1
    
    The encrypted credential can ONLY be decrypted by:
    - The same user account that created it
    - On the same computer where it was created
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [SecureString]$ClientSecret
)

try {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "App Registration Credential Encryption" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Prompt for client secret if not provided
    if (-not $ClientSecret) {
        Write-Host "Enter the App Registration Client Secret" -ForegroundColor Yellow
        $ClientSecret = Read-Host "Client Secret" -AsSecureString
        
        # Verify secret was entered
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
        $plainCheck = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        
        if ([string]::IsNullOrWhiteSpace($plainCheck)) {
            throw "Client secret cannot be empty"
        }
        $plainCheck = $null
    }
    
    # Create credential object
    $credObject = [PSCustomObject]@{
        ClientSecret = $ClientSecret
        CreatedDate = Get-Date
        CreatedBy = $env:USERNAME
        ComputerName = $env:COMPUTERNAME
    }
    
    # Ensure output directory exists
    $outputDir = Split-Path -Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path -Path $outputDir)) {
        Write-Host "Creating directory: $outputDir" -ForegroundColor Yellow
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }
    
    # Export encrypted credential
    Write-Host "Encrypting and saving credential..." -ForegroundColor Yellow
    $credObject | Export-Clixml -Path $OutputPath -Force
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Credential Encrypted Successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
    Write-Host "`nFile Details:" -ForegroundColor Cyan
    Write-Host "  Output Path: $OutputPath" -ForegroundColor White
    Write-Host "  Created By: $($credObject.CreatedBy)" -ForegroundColor White
    Write-Host "  Computer: $($credObject.ComputerName)" -ForegroundColor White
    Write-Host "  Created: $($credObject.CreatedDate)" -ForegroundColor White
    
    Write-Host "`nSecurity Information:" -ForegroundColor Yellow
    Write-Host "  This credential file can ONLY be decrypted by:" -ForegroundColor Yellow
    Write-Host "    - User: $env:USERNAME" -ForegroundColor White
    Write-Host "    - Computer: $env:COMPUTERNAME" -ForegroundColor White
    
    Write-Host "`nNext Steps:" -ForegroundColor Cyan
    Write-Host "  1. Update your configuration JSON file with this credential path" -ForegroundColor White
    Write-Host "  2. Ensure the app registration has the required permissions" -ForegroundColor White
    Write-Host "  3. Test the credential with your automation script`n" -ForegroundColor White
    
}
catch {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "Error Creating Encrypted Credential" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Error $_.Exception.Message
    exit 1
}

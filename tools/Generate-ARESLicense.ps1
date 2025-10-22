<#
.SYNOPSIS
    Generate ARES license files compatible with AresLicenseValidator.dll
.DESCRIPTION
    Creates RSA-signed JSON license files for network deployment
.PARAMETER Company
    Company name
.PARAMETER Domain
    Windows domain name
.PARAMETER AuthorizedUsers
    Array of authorized users (format: DOMAIN\username)
.PARAMETER OutputPath
    Network path where license will be saved
.EXAMPLE
    .\Generate-ARESLicense.ps1 -Company "Acme Corp" -Domain "ACME" -AuthorizedUsers @("ACME\john.doe", "ACME\jane.smith")
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$Company,
    
    [Parameter(Mandatory=$false)]
    [string]$Domain,
    
    [Parameter(Mandatory=$false)]
    [string[]]$AuthorizedUsers,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxUsers,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [string]$PrivateKeyPath
)

#Requires -Version 5.1
Add-Type -AssemblyName System.Security

# === Configuration ===
$LICENSE_FOLDER = "ARES_Licenses"
$LICENSE_FILENAME = "ares_license.json"
$LICENSE_VERSION = "1.0"

# RSA Keys - IMPORTANT: Generate with GenerateRSAKeys function and keep private key secure
$RSA_PRIVATE_KEY = @"
<RSAKeyValue>
    <Modulus>YOUR_PRIVATE_KEY_MODULUS_HERE</Modulus>
    <Exponent>AQAB</Exponent>
    <P>YOUR_P_VALUE</P>
    <Q>YOUR_Q_VALUE</Q>
    <DP>YOUR_DP_VALUE</DP>
    <DQ>YOUR_DQ_VALUE</DQ>
    <InverseQ>YOUR_INVERSEQ_VALUE</InverseQ>
    <D>YOUR_D_VALUE</D>
</RSAKeyValue>
"@

# === Functions ===

function Write-ColorOutput {
    param([string]$Message, [ConsoleColor]$Color = [ConsoleColor]::White)
    Write-Host $Message -ForegroundColor $Color
}

function New-RSAKeyPair {
    <#
    .SYNOPSIS
        Generate RSA key pair for license signing
    #>
    try {
        $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider(2048)
        
        $privateKey = $rsa.ToXmlString($true)  # Include private parameters
        $publicKey = $rsa.ToXmlString($false)  # Public only
        
        Write-ColorOutput "`n=== RSA Key Pair Generated ===" Green
        Write-ColorOutput "`nPRIVATE KEY (Keep secure - for license generation only):" Yellow
        Write-ColorOutput $privateKey White
        Write-ColorOutput "`nPUBLIC KEY (Use in C# DLL code):" Cyan
        Write-ColorOutput $publicKey White
        
        # Save to files
        $privateKey | Out-File "rsa_private_key.xml" -Encoding UTF8
        $publicKey | Out-File "rsa_public_key.xml" -Encoding UTF8
        
        Write-ColorOutput "`nKeys saved to:" Green
        Write-ColorOutput "  - rsa_private_key.xml (KEEP SECURE)" Yellow
        Write-ColorOutput "  - rsa_public_key.xml (Use in DLL)" Cyan
        
        $rsa.Dispose()
        return @{
            PrivateKey = $privateKey
            PublicKey = $publicKey
        }
    }
    catch {
        Write-ColorOutput "Error generating RSA keys: $_" Red
        return $null
    }
}

function Get-EnvironmentHash {
    param(
        [string]$Company,
        [string]$Domain
    )
    
    $environmentData = "$Company|$Domain|ARES_LICENSE_v1"
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($environmentData))
    $base64Hash = [Convert]::ToBase64String($hashBytes)
    $sha256.Dispose()
    
    return $base64Hash.Substring(0, [Math]::Min(16, $base64Hash.Length))
}

function New-LicenseKey {
    $guid = [Guid]::NewGuid().ToString("N").ToUpper()
    return "$($guid.Substring(0,8))-$($guid.Substring(8,4))-$($guid.Substring(12,4))-$($guid.Substring(16,4))-$($guid.Substring(20))"
}

function Sign-LicenseData {
    param(
        [hashtable]$LicenseData,
        [string]$PrivateKey
    )
    
    try {
        $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
        $rsa.FromXmlString($PrivateKey)
        
        # Create JSON exactly as DLL will verify it (order matters!)
        # CRITICAL: Use [ordered] hashtable to maintain field order
        $dataToSign = [ordered]@{
            company = $LicenseData.company
            domain = $LicenseData.domain
            installed_by = $LicenseData.installed_by
            installation_date = $LicenseData.installation_date
            license_key = $LicenseData.license_key
            environment_hash = $LicenseData.environment_hash
            authorized_users = $LicenseData.authorized_users
            max_users = $LicenseData.max_users
        } | ConvertTo-Json -Compress
        
        $dataBytes = [System.Text.Encoding]::UTF8.GetBytes($dataToSign)
        $signatureBytes = $rsa.SignData($dataBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        
        $rsa.Dispose()
        return [Convert]::ToBase64String($signatureBytes)
    }
    catch {
        Write-ColorOutput "Error signing data: $_" Red
        return $null
    }
}

function New-LicenseFile {
    param(
        [string]$Company,
        [string]$Domain,
        [string[]]$AuthorizedUsers,
        [int]$MaxUsers,
        [string]$OutputPath,
        [string]$PrivateKey
    )
    
    Write-ColorOutput "`n=== Generating License ===" Cyan
    
    $installedBy = "$env:USERDOMAIN\$env:USERNAME"
    $installationDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $licenseKey = New-LicenseKey
    $environmentHash = Get-EnvironmentHash -Company $Company -Domain $Domain
    
    # Create license data structure
    # CRITICAL: Order must match exactly with C# DLL verification order
    $licenseData = [ordered]@{
        company = $Company
        domain = $Domain
        installed_by = $installedBy
        installation_date = $installationDate
        license_key = $licenseKey
        environment_hash = $environmentHash
        authorized_users = $AuthorizedUsers
        max_users = $MaxUsers
        version = $LICENSE_VERSION
    }
    
    Write-ColorOutput "Company:          $Company" White
    Write-ColorOutput "Domain:           $Domain" White
    Write-ColorOutput "Max Users:        $MaxUsers" White
    Write-ColorOutput "Authorized Users: $($AuthorizedUsers.Count)" White
    foreach ($user in $AuthorizedUsers) {
        Write-ColorOutput "  - $user" Gray
    }
    Write-ColorOutput "License Key:      $licenseKey" White
    Write-ColorOutput "Environment Hash: $environmentHash" White
    
    # Sign the license
    Write-ColorOutput "`nSigning license..." Yellow
    $signature = Sign-LicenseData -LicenseData $licenseData -PrivateKey $PrivateKey
    
    if (-not $signature) {
        throw "Failed to sign license"
    }
    
    $licenseData.signature = $signature
    
    # Convert to JSON
    $jsonContent = $licenseData | ConvertTo-Json -Depth 10
    
    # Ensure output directory exists
    $fullOutputPath = Join-Path $OutputPath $LICENSE_FOLDER
    if (-not (Test-Path $fullOutputPath)) {
        New-Item -ItemType Directory -Force -Path $fullOutputPath | Out-Null
    }
    
    # Save license file
    $licensePath = Join-Path $fullOutputPath $LICENSE_FILENAME
    $jsonContent | Out-File -FilePath $licensePath -Encoding UTF8 -Force
    
    Write-ColorOutput "`n=== License Generated Successfully ===" Green
    Write-ColorOutput "File location: $licensePath" Cyan
    Write-ColorOutput "`nDeploy this file to the network location accessible by all users." Yellow
    Write-ColorOutput "Path format: \\server\share\$LICENSE_FOLDER\$LICENSE_FILENAME" Gray
    
    return $licensePath
}

function Get-UserInput {
    param([string]$Prompt, [string]$Default = "")
    
    if ($Default) {
        $input = Read-Host "$Prompt (default: $Default)"
        if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
        return $input
    }
    else {
        do {
            $input = Read-Host $Prompt
        } while ([string]::IsNullOrWhiteSpace($input))
        return $input
    }
}

function Show-Menu {
    Clear-Host
    Write-ColorOutput "========================================" Cyan
    Write-ColorOutput "   ARES License Generator v2.0" Cyan
    Write-ColorOutput "   Compatible with AresLicenseValidator" Cyan
    Write-ColorOutput "========================================" Cyan
    Write-ColorOutput ""
}

# === Main Execution ===

try {
    Show-Menu
    
    # Check if RSA keys are configured
    if ($RSA_PRIVATE_KEY -match "YOUR_PRIVATE_KEY") {
        Write-ColorOutput "WARNING: RSA keys not configured!" Yellow
        Write-ColorOutput ""
        $generate = Read-Host "Generate new RSA key pair? (Y/N)"
        
        if ($generate -eq 'Y' -or $generate -eq 'y') {
            $keys = New-RSAKeyPair
            if ($keys) {
                Write-ColorOutput "`nUpdate this script with the private key and" Yellow
                Write-ColorOutput "update your C# DLL with the public key!" Red
                Write-ColorOutput "`nPress any key to exit..." Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                exit 0
            }
        }
        
        Write-ColorOutput "Cannot continue without RSA keys." Red
        exit 1
    }
    
    # Load private key from file if path provided
    if ($PrivateKeyPath -and (Test-Path $PrivateKeyPath)) {
        $RSA_PRIVATE_KEY = Get-Content $PrivateKeyPath -Raw
    }
    
    # Get parameters if not provided
    if (-not $Company) {
        $Company = Get-UserInput -Prompt "Company name"
    }
    
    if (-not $Domain) {
        $Domain = Get-UserInput -Prompt "Windows domain name" -Default $env:USERDOMAIN
    }
    
    if (-not $AuthorizedUsers -or $AuthorizedUsers.Count -eq 0) {
        Write-ColorOutput "`nEnter authorized users (format: DOMAIN\username)" White
        Write-ColorOutput "Examples: GRPL\e.duperey, ACME\john.doe, CORP\jane-smith" Gray
        Write-ColorOutput "Enter blank line when done" Gray
        $userList = @()
        do {
            $user = Read-Host "User $($userList.Count + 1)"
            if (-not [string]::IsNullOrWhiteSpace($user)) {
                # Validate format - allow alphanumeric, dots, hyphens, underscores
                if ($user -match '^[\w\-]+\\[\w\.\-]+$') {
                    $userList += $user
                    Write-ColorOutput "  Added: $user" Green
                } else {
                    Write-ColorOutput "  Invalid format. Use: DOMAIN\username (letters, numbers, dots, hyphens allowed)" Yellow
                }
            }
        } while (-not [string]::IsNullOrWhiteSpace($user))
        
        if ($userList.Count -eq 0) {
            throw "At least one user must be specified"
        }
        
        $AuthorizedUsers = $userList
    }
    
    if (-not $MaxUsers) {
        $MaxUsers = $AuthorizedUsers.Count
        $maxInput = Get-UserInput -Prompt "Max concurrent users" -Default $MaxUsers.ToString()
        $MaxUsers = [int]$maxInput
    }
    
    if (-not $OutputPath) {
        $defaultPath = "\\$env:USERDOMAIN-SRV\Shared"
        Write-ColorOutput "`nOutput path should be a network location accessible by all users" Yellow
        $OutputPath = Get-UserInput -Prompt "Network path" -Default $defaultPath
    }
    
    # Generate license
    $licensePath = New-LicenseFile `
        -Company $Company `
        -Domain $Domain `
        -AuthorizedUsers $AuthorizedUsers `
        -MaxUsers $MaxUsers `
        -OutputPath $OutputPath `
        -PrivateKey $RSA_PRIVATE_KEY
    
    Write-ColorOutput ""
    Write-ColorOutput "========================================" Cyan
    Write-ColorOutput "Press any key to exit..." Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
catch {
    Write-ColorOutput ""
    Write-ColorOutput "ERROR: $_" Red
    Write-ColorOutput ""
    Write-ColorOutput "Press any key to exit..." Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
# ARES PowerShell Tools

PowerShell utilities for ARES license management.

## Generate-ARESLicense.ps1

Generate RSA-signed JSON license files for network deployment.

### Prerequisites

- PowerShell 5.1 or higher
- .NET Framework 4.7.2+
- Administrator privileges (for network path access)
- RSA key pair (generated on first run)

### Quick Start

#### 1. Generate RSA Key Pair (First Time Only)
powershell:
.\Generate-ARESLicense.ps1
# Choose 'Y' when prompted to generate keys
This creates:

rsa_private_key.xml - Keep secure, used for license generation<br>
rsa_public_key.xml - Use in C# DLL code

CRITICAL: Update your C# DLL with the public key:

// In LicenseValidatorService.cs
```
private const string PUBLIC_KEY = @"<RSAKeyValue>
    <Modulus>YOUR_GENERATED_PUBLIC_KEY_HERE</Modulus>
    <Exponent>AQAB</Exponent>
</RSAKeyValue>";
```

#### 2. Update Script Configuration
Open Generate-ARESLicense.ps1 and replace $RSA_PRIVATE_KEY with the content of rsa_private_key.xml.

#### 3. Generate License
#Interactive Mode:
.\Generate-ARESLicense.ps1

#Command Line Mode:
```
.\Generate-ARESLicense.ps1 `
    -Company "Acme Corporation" `
    -Domain "ACME" `
    -AuthorizedUsers @("ACME\john.doe", "ACME\jane.smith", "ACME\bob.wilson") `
    -MaxUsers 3 `
    -OutputPath "\\fileserver\shared"
```

#Parameters
```
Parameter			Required	Default			Description	
Company				No*			-				Company name (used in environment hash)
Domain				No*			Current domain	Windows domain name
AuthorizedUsers		No*			-				Array of users (format: DOMAIN\username)
MaxUsers			No			User count		Maximum concurrent users
OutputPath			No*			-				Network path for license file
PrivateKeyPath		No			-				Path to RSA private key XML file
```

*If not provided via command line, the script will prompt interactively.

#License File Structure
The script generates: \\OutputPath\ARES_Licenses\ares_license.json<br>
```
{
  "company": "Acme Corporation",
  "domain": "ACME",
  "installed_by": "ACME\\admin",
  "installation_date": "2025-01-03 14:30:22",
  "license_key": "A1B2C3D4-E5F6-G7H8-I9J0-K1L2M3N4O5P6",
  "environment_hash": "xY9zK2mN5pQ8rT3w",
  "authorized_users": [
    "ACME\\john.doe",
    "ACME\\jane.smith",
    "ACME\\bob.wilson"
  ],
  "max_users": 3,
  "signature": "Base64EncodedRSASignature==",
  "version": "1.0"
}
```
#License Validation Process<br>
The DLL validates licenses by checking:

File Location - Searches network drives (Z: to K:) and common UNC paths<br>
JSON Format - Must be valid JSON with all required fields<br>
RSA Signature - Verifies data integrity using public key<br>
Windows Domain - Must match current user's domain<br>
Authorized Users - Current user must be in the list<br>
Environment Hash - SHA256(company|domain|ARES_LICENSE_v1)<br>

#Network Deployment
```
\\fileserver\shared\
└── ARES_Licenses\
    └── ares_license.json
```

#Required Permissions:<br>
Read access for all authorized users<br>
Write access only for license administrators<br>
Network share must be accessible from all client machines<br>

#Common Network Paths:
\\server\shared\ARES_Licenses\<br>
Z:\ARES_Licenses\ (mapped drive)<br>
\\fileserver\public\ARES_Licenses\<br>
\\nas\common\ARES_Licenses\<br>

#Security Best Practices

##Private Key Protection<br>
Store rsa_private_key.xml in a secure location<br>
Never commit private key to version control<br>
Restrict access to license administrators only<br>
Consider using a password-protected key store<br>


##License File Security<br>
Set NTFS read-only permissions for end users<br>
Enable file auditing for license access<br>
Regular backup of license files<br>
Monitor for unauthorized modifications<br>


##User Management<br>
Use Windows AD groups when possible<br>
Regular audit of authorized user list<br>
Remove users who leave the organization<br>
Document license assignment process<br>

#Troubleshooting<br>
Error: "RSA keys not configured"<br>
Generate new keys<br>
.\Generate-ARESLicense.ps1<br>
Choose 'Y' at prompt<br>
Update script and DLL with generated keys<br>

#Error: "Cannot continue without RSA keys"<br>
Ensure $RSA_PRIVATE_KEY in script is updated with valid key<br>
Or use -PrivateKeyPath parameter to load key from file<br>

#License validation fails in MicroStation<br>
Verify username format: DOMAIN\username (case-sensitive)<br>
Check network path accessibility from client machine<br>
Confirm Windows domain matches license domain<br>
Ensure user is in authorized_users array<br>
Verify license file permissions (read access required)<br>

#Error: "Domain mismatch"<br>
License domain must match current Windows domain exactly<br>
Check with: $env:USERDOMAIN in PowerShell<br>
Generate new license if domain has changed<br>

#Error: "User not authorized"<br>
Verify exact username with: whoami in cmd<br>
Check user format in license: DOMAIN\username<br>
Usernames are case-insensitive but domain must match<br>

#Advanced Usage<br>
Generate License with External Key File:<br>
```
powershell.\Generate-ARESLicense.ps1 `
    -PrivateKeyPath "C:\SecureKeys\ares_private.xml" `
    -Company "Acme Corp" `
    -Domain "ACME" `
    -AuthorizedUsers @("ACME\user1", "ACME\user2")
```

#Batch License Generation (Multiple Domains):<br>
```
powershell$licenses = @(
    @{Company="Acme Inc"; Domain="ACME"; Users=@("ACME\user1")},
    @{Company="Beta Corp"; Domain="BETA"; Users=@("BETA\user2")}
)

foreach ($lic in $licenses) {
    .\Generate-ARESLicense.ps1 `
        -Company $lic.Company `
        -Domain $lic.Domain `
        -AuthorizedUsers $lic.Users `
        -OutputPath "\\server\licenses\$($lic.Domain)"
}
```
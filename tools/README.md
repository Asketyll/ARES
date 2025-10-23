# ARES PowerShell Tools

PowerShell utilities for ARES license management.

## Generate-ARESLicense.ps1

Generate RSA-signed JSON license files for network deployment.

### Prerequisites

- PowerShell 5.1 or higher
- .NET Framework 4.7.2+
- Administrator privileges (for network path access)
- RSA key pair (generated on first run)

## Quick Start

### 1. Generate RSA Key Pair (First Time Only)
```powershell
.\Generate-ARESLicense.ps1
```


**Choose 'Y' when prompted to generate keys**

This creates:

- `rsa_private_key.xml` - Keep secure, used for license generation
- `rsa_public_key.xml` - Use in C# DLL code

**CRITICAL:** Update your C# DLL with the public key:
```csharp
// In LicenseValidatorService.cs
private const string PUBLIC_KEY = @"<RSAKeyValue>
    <Modulus>YOUR_GENERATED_PUBLIC_KEY_HERE</Modulus>
    <Exponent>AQAB</Exponent>
</RSAKeyValue>";
```

### 2. Update Script Configuration

Open `Generate-ARESLicense.ps1` and replace `$RSA_PRIVATE_KEY` with the content of `rsa_private_key.xml`.

### 3. Generate License

**Interactive Mode:**
```powershell
.\Generate-ARESLicense.ps1
```

**Command Line Mode:**
```powershell
.\Generate-ARESLicense.ps1 `
    -Company "Acme Corporation" `
    -Domain "ACME" `
    -AuthorizedUsers @("ACME\john.doe", "ACME\jane.smith", "ACME\bob.wilson") `
    -MaxUsers 3 `
    -OutputPath "\\fileserver\shared"
```

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| Company | No* | - | Company name (used in environment hash) |
| Domain | No* | Current domain | Windows domain name |
| AuthorizedUsers | No* | - | Array of users (format: DOMAIN\username) |
| MaxUsers | No | User count | Maximum concurrent users |
| OutputPath | No* | - | Network path for license file |
| PrivateKeyPath | No | - | Path to RSA private key XML file |

*If not provided via command line, the script will prompt interactively.

## License File Structure

The script generates: `\\OutputPath\ARES_Licenses\ares_license.json`
```json
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

## License Validation Process<br>

The DLL validates licenses by checking:

- **File Location** - Searches network drives (Z: to K:) and common UNC paths
- **JSON Format** - Must be valid JSON with all required fields
- **RSA Signature** - Verifies data integrity using public key
- **Windows Domain** - Must match current user's domain
- **Authorized Users** - Current user must be in the list
- **Environment Hash** - SHA256(company|domain|ARES_LICENSE_v1)

## Network Deployment
```
\\fileserver\shared\
└── ARES_Licenses\
    └── ares_license.json
```

### Required Permissions:

- Read access for all authorized users
- Write access only for license administrators
- Network share must be accessible from all client machines

### Common Network Paths

- `\\server\shared\ARES_Licenses\`
- `Z:\ARES_Licenses\` (mapped drive)
- `\\fileserver\public\ARES_Licenses\`
- `\\nas\common\ARES_Licenses\`

## Security Best Practices

### Private Key Protection

- Store `rsa_private_key.xml` in a secure location
- Never commit private key to version control
- Restrict access to license administrators only
- Consider using a password-protected key store

### License File Security

- Set NTFS read-only permissions for end users
- Enable file auditing for license access
- Regular backup of license files
- Monitor for unauthorized modifications

### User Management

- Use Windows AD groups when possible
- Regular audit of authorized user list
- Remove users who leave the organization
- Document license assignment process

## Troubleshooting

### Error: "RSA keys not configured"

Generate new keys:
```powershell
.\Generate-ARESLicense.ps1
```

Choose 'Y' at prompt and update script and DLL with generated keys.

### Error: "Cannot continue without RSA keys"

- Ensure `$RSA_PRIVATE_KEY` in script is updated with valid key
- Or use `-PrivateKeyPath` parameter to load key from file

### License validation fails in MicroStation

- Verify username format: `DOMAIN\username` (case-sensitive)
- Check network path accessibility from client machine
- Confirm Windows domain matches license domain
- Ensure user is in `authorized_users` array
- Verify license file permissions (read access required)

### Error: "Domain mismatch"

- License domain must match current Windows domain exactly
- Check with: `$env:USERDOMAIN` in PowerShell
- Generate new license if domain has changed

### Error: "User not authorized"

- Verify exact username with: `whoami` in cmd
- Check user format in license: `DOMAIN\username`
- Usernames are case-insensitive but domain must match

## Advanced Usage

### Generate License with External Key File
```powershell
.\Generate-ARESLicense.ps1 `
    -PrivateKeyPath "C:\SecureKeys\ares_private.xml" `
    -Company "Acme Corp" `
    -Domain "ACME" `
    -AuthorizedUsers @("ACME\user1", "ACME\user2")
```

### Batch License Generation (Multiple Domains)
```powershell
$licenses = @(
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
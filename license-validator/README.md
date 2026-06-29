# ARES License Validator

COM-visible .NET DLL used by MicroStation VBA to validate ARES network licenses.

It reads an RSA-signed JSON license file from a shared network location and verifies it
against the current Windows environment. There is no online activation server and no
hardware fingerprinting.

## What it validates

- **Network license file** - searches mapped drives (Z: down to K:) and common UNC paths for `ARES_Licenses\ares_license.json`
- **RSA signature** - verifies the license data with an embedded RSA public key (RSA + SHA256)
- **Windows domain** - current `UserDomainName` must match the license `domain`
- **Authorized users** - current `DOMAIN\user` must be in the license `authorized_users` list
- **Environment hash** - must equal `SHA256(company|domain|ARES_LICENSE_v1)` (first 16 Base64 chars)

License files are generated with `tools/Generate-ARESLicense.ps1`.

## COM interface

- ProgID: `ARES.LicenseValidator`
- Class: `AresLicenseValidator` (implements `IAresLicenseValidator`)

Public methods:

| Method | Returns | Description |
|--------|---------|-------------|
| `ValidateLicense()` | `Boolean` | `True` if the license file is found, correctly signed, and matches the current domain / user / environment |
| `GetLicenseInfo()` | `String` | Human-readable summary (company, domain, licensed users, install date/admin) |
| `GetLastError()` | `String` | Last validation error message |
| `GetCurrentUser()` | `String` | Current `DOMAIN\user` |
| `GetAuthorizedUserCount()` | `Integer` | `max_users` value from the license |

## Building

Requires .NET Framework 4.7.2 and Newtonsoft.Json 13.0.3 (restored via NuGet).

```bash
# Restore NuGet packages
nuget restore AresLicenseValidator.sln

# Build with MSBuild
msbuild AresLicenseValidator.sln /p:Configuration=Release

# Register COM DLL (also run automatically as a post-build step)
regasm AresLicenseValidator.dll /tlb /codebase
```

> The embedded RSA public key (`PUBLIC_KEY` in `Services/LicenseValidator.cs`) must match the
> private key used by `Generate-ARESLicense.ps1`, otherwise signature validation always fails.

## Usage in VBA

```vb
Private oLicenseValidator As Object
Set oLicenseValidator = CreateObject("ARES.LicenseValidator")

If oLicenseValidator.ValidateLicense() Then
    ' License valid
Else
    ' License invalid - oLicenseValidator.GetLastError() explains why
End If
```

## License

AGPL-3.0 - See [LICENSE](../LICENSE) file for details.

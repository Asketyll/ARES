# ARES License Validator

COM-visible DLL for license validation in MicroStation VBA.

## Features

- AES-256 encryption
- User and environment validation
- Hardware-based licensing
- Offline license validation

## Building
```bash
# Restore NuGet packages
nuget restore AresLicenseValidator.sln

# Build with MSBuild
msbuild AresLicenseValidator.sln /p:Configuration=Release

# Register COM DLL
regasm AresLicenseValidator.dll /tlb /codebase
```

## Usage in VBA
```vb
Private oLicenseValidator As Object
Set oLicenseValidator = CreateObject("ARES.LicenseValidator")

If oLicenseValidator Is Nothing Then
		' License invalid
        Exit Function
End If

If oLicenseValidator.ValidateLicense() Then
    ' License valid
Else
    ' License invalid
End If
```
# ARES License Validator

COM-visible DLL for license validation in MicroStation VBA.

## Features

- AES-256 encryption
- User and environment validation
- Hardware-based licensing
- Offline license validation

## Building

nuget restore AresLicenseValidator.sln
msbuild AresLicenseValidator.sln /p:Configuration=Release

# Register COM DLL
regasm AresLicenseValidator.dll /tlb /codebase

# Usage in VBA
Dim validator As New AresLicenseValidator.LicenseValidator
If validator.ValidateLicense("license-file-path") Then
    ' License valid
Else
    ' License invalid
End If
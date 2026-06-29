# ARES Installer

Windows installer application for the ARES MicroStation Add-in. It downloads the latest ARES release from GitHub, verifies it, installs the MVBA project to `C:\ARES\` (DLLs to `C:\ARES\Rsc\`), and registers the license-validator COM component.

## Features

- Multi-language UI (English/French), chosen at startup
- Administrator privilege enforcement (self-elevates via UAC)
- Prerequisite checks (.NET Framework 4.7.2+)
- Automatic download of the latest release from the GitHub releases API
- Mandatory SHA-256 integrity verification of downloaded assets (from the GitHub API digest)
- COM registration of `AresLicenseValidator.dll` via regasm (`/tlb /codebase`)
- Versioned install: backs up previous DLLs and records the version in the registry (`HKCU\Software\ARES\Version`)
- Bentley product selection after install (detects installed Bentley products from the registry)
- Progress bar and on-screen log

## Building

Run from the `installer/` directory:

```bash
# Restore NuGet packages
nuget restore AresInstaller.sln

# Build with MSBuild
msbuild AresInstaller.sln /p:Configuration=Release
```

## Development

Built with:

- .NET Framework 4.7.2
- C# Windows Forms (packaged as a single self-contained executable)

NuGet dependencies (see `AresInstaller/packages.config`):

- **Newtonsoft.Json** 13.0.4 — parses the GitHub releases API response
- **Fody** 6.9.3 + **Costura.Fody** 6.0.0 (build-time) — embed referenced assemblies into a single `AresInstaller.exe`

## License

AGPL-3.0 - See [LICENSE](../LICENSE) file for details.

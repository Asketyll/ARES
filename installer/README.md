# ARES Installer

Windows installer application for the ARES MicroStation Add-in. It downloads the latest ARES release from GitHub, verifies it, installs `ARES.mvba` to `C:\ARES\`, and copies the release resources (e.g. the custom-property `.dgnlib`) to `C:\ARES\Rsc\`. ARES has no licensing and no COM component to register.

## Features

- Multi-language UI (English/French), chosen at startup
- Administrator privilege enforcement (self-elevates via UAC)
- Prerequisite checks (.NET Framework 4.7.2+)
- Automatic download of the latest release from the GitHub releases API
- Mandatory SHA-256 integrity verification of downloaded assets (from the GitHub API digest)
- Records the installed version in the registry (`HKCU\Software\ARES\Version`); existing resources are overwritten
- Bentley product selection after install (detects installed Bentley products from the registry); configures the chosen product's `Personal.ucf` to auto-load `ARES.mvba` (`MS_VBAAUTOLOADPROJECTS`) and load the resource dgnlib (`MS_DGNLIBLIST` → `C:\ARES\Rsc`)
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

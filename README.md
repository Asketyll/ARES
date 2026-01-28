# ARES - MicroStation Productivity Add-in

Automatic length calculation and graphical linking tools for MicroStation VBA.

[![License: AGPL-3.0](https://img.shields.io/badge/License-AGPL%203.0-blue.svg)](LICENSE)
[![GitHub release](https://img.shields.io/github/release/Asketyll/ARES.svg)](https://github.com/Asketyll/ARES/releases/latest)

## Quick Install

```powershell
Invoke-WebRequest -Uri "https://github.com/Asketyll/ARES/releases/download/installer-v1.0.0/AresInstaller.exe" -OutFile "$env:TEMP\AresInstaller.exe"; Start-Process "$env:TEMP\AresInstaller.exe" -Verb RunAs
```

## Downloads

| Component | Release | Description |
|-----------|---------|-------------|
| [Installer](https://github.com/Asketyll/ARES/releases/tag/installer-v1.0.0) | `installer-v1.0.0` | Windows installer with automatic setup |
| [MVBA](https://github.com/Asketyll/ARES/releases/latest) | `v1.0.0` | MicroStation VBA source code |

## Components

| Directory | Description | Documentation |
|-----------|-------------|---------------|
| [`MVBA/`](MVBA/) | MicroStation VBA project - core add-in functionality | [README](MVBA/README.md) |
| [`installer/`](installer/) | C# Windows Forms installer application | [README](installer/README.md) |
| [`license-validator/`](license-validator/) | COM-visible DLL for license validation | [README](license-validator/README.md) |
| [`tools/`](tools/) | PowerShell scripts for license generation | [README](tools/README.md) |

## Features

- **Auto Lengths** - Automatic length calculation for linked graphical elements
- **License Management** - AES-256 encrypted, RSA-signed network licenses
- **Multi-language** - French/English interface support
- **Bulk Operation Detection** - Auto-suspend during merge/reprojection for performance

## Requirements

- Windows 10/11
- .NET Framework 4.7.2+
- MicroStation Connect Edition, OpenCities Map PowerView, or Atlas/Eras
- Administrator privileges for installation

## Usage

1. Run `AresInstaller.exe` as Administrator
2. Choose language (English/Français)
3. Click Install and wait for completion
4. Load `C:\ARES\ARES.mvba` in MicroStation

## Development

See component READMEs for build instructions:
- [MVBA Development](MVBA/README.md#development)
- [Installer Build](installer/README.md#building)
- [License Validator Build](license-validator/README.md#building)

## Contributing

Contributions welcome. Please submit a Pull Request.

## License

AGPL-3.0 - See [LICENSE](LICENSE) for details.

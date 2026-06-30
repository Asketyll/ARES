# ARES - MicroStation Productivity Add-in

Automatic length calculation and graphical linking tools for MicroStation VBA.

[![License: AGPL-3.0](https://img.shields.io/badge/License-AGPL%203.0-yellow.svg)](LICENSE) &nbsp;&nbsp; [![GitHub release](https://img.shields.io/github/release/Asketyll/ARES.svg)](https://github.com/Asketyll/ARES/releases/latest) &nbsp;&nbsp; [![Wiki](https://img.shields.io/badge/docs-wiki-brightgreen.svg)](https://github.com/Asketyll/ARES/wiki)

## Quick Install

```powershell
Invoke-WebRequest -Uri "https://github.com/Asketyll/ARES/releases/download/installer-v1.1.0/AresInstaller.exe" -OutFile "$env:TEMP\AresInstaller.exe"; Start-Process "$env:TEMP\AresInstaller.exe" -Verb RunAs
```

## Downloads

| Component | Release | Description |
|:---------:|:-------:|:------------|
| [Installer](https://github.com/Asketyll/ARES/releases) | ![Installer release](https://img.shields.io/github/v/release/Asketyll/ARES?filter=installer-%2A&display_name=tag&label=&color=555) | Windows installer with automatic setup |
| [MVBA](https://github.com/Asketyll/ARES/releases/latest) | ![MVBA release](https://img.shields.io/github/v/release/Asketyll/ARES?filter=v%2A&display_name=tag&label=&color=555) | MicroStation VBA source code |

## Components

| Directory | Tech | Role | Docs |
|:--------:|:----:|:-----|:----:|
| [`MVBA/`](MVBA/) | VBA 7.1 | The MicroStation add-in (core functionality) | [README](MVBA/README.md) |
| [`installer/`](installer/) | C# WinForms | Windows installer; writes `HKCU\Software\ARES\Version` | [README](installer/README.md) |

## Features

User features — full key-in reference and configuration variables live in the **[Wiki](https://github.com/Asketyll/ARES/wiki)** ([Version FR](https://github.com/Asketyll/ARES/wiki/Accueil)):

| Feature | Description | Docs |
|:-------:|:------------|:----:|
| Auto Lengths | Automatic length calculation for linked graphical elements (+ color sync) | [EN](https://github.com/Asketyll/ARES/wiki/Auto-Lengths)&nbsp;·&nbsp;[FR](https://github.com/Asketyll/ARES/wiki/Longueurs-Auto) |
| Zoning | Buffer zone generation around elements (configurable distance, level, style, color, weight) | [EN](https://github.com/Asketyll/ARES/wiki/Zoning)&nbsp;·&nbsp;[FR](https://github.com/Asketyll/ARES/wiki/Zonage) |
| Zone Export | Element lengths (partial or full) inside zone polygons, exported per group to Excel | [EN](https://github.com/Asketyll/ARES/wiki/Zone-Export)&nbsp;·&nbsp;[FR](https://github.com/Asketyll/ARES/wiki/Export-de-Zone) |
| Region Split | Split a closed region into two with a single datapoint on its boundary | [EN](https://github.com/Asketyll/ARES/wiki/Region-Split)&nbsp;·&nbsp;[FR](https://github.com/Asketyll/ARES/wiki/Decoupe-de-Region) |

System:
- **Multi-language** - French/English interface ([EN](https://github.com/Asketyll/ARES/wiki/System-and-Config)&nbsp;·&nbsp;[FR](https://github.com/Asketyll/ARES/wiki/Systeme-et-Config))
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

## Contributing

Contributions welcome. Please submit a Pull Request.

## License

AGPL-3.0 - See [LICENSE](LICENSE) for details.

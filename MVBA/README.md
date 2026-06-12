# ARES - MicroStation VBA Project

This directory contains the complete MVBA (MicroStation Visual Basic for Applications) source code.

## Structure

- **Command/** - User-facing command entry points
- **Components/** - Shared reusable modules (Length, GetElements, FileDialogs, Link, …)
- **Configuration/** - Configuration management (ARESConfigClass, LangManager, Config)
- **Core/** - System core (ARESConstants, BootLoader, error handling)
- **EventHandlers/** - Event management (DGN open/close, element changes)
- **Features/** - Business features, each in its own sub-folder:
  - **AutoLengths/** - Automatic length calculation with GUI
  - **Zoning/** - Buffer zone generation around elements with GUI
  - **ZoneExport/** - Element length export inside zones to Excel with GUI
- **Security/** - License validation, AES-256 encryption, UUID, RSA signing
- **Tests/** - Unit testing suite
- **Update/** - Automatic update checker

## Development

### Loading in MicroStation

1. Open MicroStation VBA editor
2. File → Import File
3. Import all `.bas`, `.cls`, and `.frm` files maintaining the structure
4. Compile as `ARES.mvba`
5. Create ARES License

### Dependencies

- Tested on MicroStation Connect Edition, OpenCities Map PowerView by Bentley Systems and Atlas/Eras by Sogelink
- MVBA 7.1 environment

## License

AGPL-3.0 - See [LICENSE](../LICENSE) file for details.

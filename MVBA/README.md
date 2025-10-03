# ARES - MicroStation VBA Project

This directory contains the complete MVBA (MicroStation Visual Basic for Applications) source code.

## Structure

- **Command/** - User commands interface
- **Components/** - Functional components (lengths, links, graphics)
- **Configuration/** - Configuration and settings management
- **Core/** - System core (constants, boot loader, error handling)
- **EventHandlers/** - Event management (DGN open/close, element changes)
- **LengthsFeature/** - Auto-lengths functionality with GUI
- **Security/** - Security and validation (encryption, UUID)
- **Tests/** - Unit testing suite

## Development

### Loading in MicroStation

1. Open MicroStation VBA editor (Alt+F11)
2. File → Import File
3. Import all .bas, .cls, and .frm files maintaining the structure
4. Compile as ARES.mvba

### Dependencies

- Tested on MicroStation Connect Edition, OpenCities Map PowerView by Bentley Systems and Atlas/Eras by Sogelink.
- MVBA 7.1 environment

## License

AGPL-3.0 - See [LICENSE](../LICENSE) file for details.

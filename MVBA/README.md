# ARES - MicroStation VBA Project

This directory contains the complete MVBA (MicroStation Visual Basic for Applications) source code.

> User-facing key-ins and configuration variables are documented in the wiki:
> https://github.com/Asketyll/ARES/wiki
> This README is a developer/technical reference and does not duplicate the wiki's per-key-in / per-variable tables.

## Structure

- **Command/** - User-facing key-in entry points (`Command.bas`)
- **Components/** - Shared reusable modules: `Geometry`, `Length`, `Link`, `StringsInEl`, `GetElements`, `CustomPropertyHandler`, `MicroStationDefinition`, `MSGraphicalInteraction`, `CellRedreaw`, `FileDialogs`
- **Configuration/** - Configuration management: `ARESConfigClass`, `ARES_MS_VAR_Class`, `Config`, `LangManager`
- **Core/** - System core: `BootLoader`, `ARESConstants`, `ElementInProcesseClass`, `ErrorHandlerClass`, `ColorDialog`
- **EventHandlers/** - Event management: `DGNOpenClose`, `ElementChangeHandler`, `IdleEventHandler`, `ReRegisterIdleHandler`
- **Features/** - Business features, each in its own sub-folder:
  - **AutoLengths/** - Automatic length calculation with GUI (`Auto_Lengths.cls` + forms)
  - **Zoning/** - Buffer-zone generation around elements, with GUI (`Zoning.bas` + form)
  - **ZoneExport/** - Element length export inside zones to Excel, with GUI (`ExportLengthInRegion.bas` + form)
  - **RegionSplit/** - Single-datapoint cut of a closed region into two regions (`RegionSplit.bas` engine + `RegionSplitLocate.cls` driver)
- **Security/** - License validation, AES-256 encryption, machine UUID, in-VBA environment/anti-piracy validation
- **Tests/** - Unit testing harness (deprecated / unmaintained)
- **Update/** - Automatic update checker (`UpdateChecker.bas` + form)

## Development

### Loading in MicroStation

1. Open MicroStation VBA editor
2. File → Import File
3. Import all `.bas`, `.cls`, and `.frm` files maintaining the structure
4. Compile as `ARES.mvba`
5. Create ARES License

### Dependencies

- Tested on MicroStation Connect Edition, OpenCities Map PowerView by Bentley Systems and Atlas/Eras by Sogelink
- VBA 7.1 environment

## License

AGPL-3.0 - See [LICENSE](../LICENSE) file for details.

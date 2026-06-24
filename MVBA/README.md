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
  - **Zoning/** - Buffer zone generation around elements with GUI. Key-ins: `RunZoning` (`ARES_Zoning_Distance`, default 2.0 m; rounded caps; merged) and `RunZoning2` (`ARES_Zoning2_Distance`, default 0.2 m; flat caps; per-element fusion only — no cross-element merge)
  - **ZoneExport/** - Element length export inside zones to Excel with GUI
  - **RegionSplit/** - Split a closed region (Shape / ComplexShape) into two regions with a single datapoint on its boundary. Key-in: `SplitRegion`. The cut runs perpendicular to the local boundary at the clicked point, across to the opposite boundary; both halves inherit the original's level + symbology. **`ComplexShape` boundaries with arcs are supported** — clicking an arc side cuts radially (perpendicular to the arc tangent, i.e. along the radius), while clicking a straight side cuts perpendicular to that segment. Config vars: `ARES_RegionSplit_Collinear_Tol` (default `0.000001`, geometric epsilon; also drives the cut-knife half-width), `ARES_RegionSplit_Stroke_Tol` (default `0.01`, max chordal deviation in master units when densifying an arc boundary side into a polyline — smaller ⇒ more chords ⇒ a closer-to-radial cut and a tighter entry foot; must be `> 0`), `ARES_RegionSplit_Keep_Original` (default `False`; `True` keeps the original alongside the two halves)
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

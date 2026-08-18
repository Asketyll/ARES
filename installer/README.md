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

## Resource upgrades (`ARES_Custom_Properties.dgnlib`)

The `.dgnlib` is a **versioned deliverable**, not a static asset: some MVBA features require ItemType definitions that ship inside it. Since epic 15 it carries **two** ItemTypeLibraries — the user-facing `ARES` library (one ItemType per custom property, authored by the site) and the internal **`ARES_SYS`** library (one ItemType `ARES_Render`, two String properties `SchemaVersion` + `Entries`) that Property Rendering stores its text bindings in.

Both are found through the existing `MS_DGNLIBLIST > c:/ares/rsc/*.dgnlib` wildcard, so no installer parameter changes. What does matter is the **upgrade path on a site that authored its own properties**: the deployed `.dgnlib` is overwritten by an install, so a site copy must be merged rather than replaced — re-author or import the site's `ARES` ItemTypes into the shipped file, keeping `ARES_SYS` intact.

Version-mismatch matrix, both directions fail-closed and neither corrupts data:

| Station | File | Behaviour |
|---|---|---|
| Old MVBA + old `.dgnlib` | file carrying render bindings | The bindings are inert data; the text keeps its last rendered values. Nothing is rewritten. |
| New MVBA + old `.dgnlib` (no `ARES_SYS`) | any | Property Rendering **self-disables**: no bind, no render, no partial write — one English line in the `.log` plus a translated status. |
| New MVBA + new `.dgnlib` | file written by a NEWER ARES (unknown `SchemaVersion`) | Refuses to interpret the binding and never rewrites it; one English log line + status. |

Note that a DWG/V7 round trip, or a save from a non-ARES station, strips Item Types altogether: the text survives as clean frozen values and the binding must be re-authored. That is the accepted cost of emitting scaffolding-free deliverables.

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

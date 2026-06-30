# ARES — MicroStation VBA Project

This directory contains the complete MVBA (MicroStation Visual Basic for Applications) source code of ARES. This README is the **developer/technical reference** — architecture, lifecycle, the element-processing flow, coding conventions, and review rules.

For the repository as a whole (installer), see the [main README](../README.md).

> **User-facing** key-ins and configuration variables are documented in the wiki:
> https://github.com/Asketyll/ARES/wiki
> This README does not duplicate the wiki's per-key-in / per-variable tables.

## Contents

- [Architecture constraints (VBA 7.1)](#architecture-constraints-vba-71)
- [Project structure](#project-structure)
- [Boot & lifecycle](#boot--lifecycle)
- [Core element-processing flow](#core-element-processing-flow-event--idle--process)
- [Features](#features)
- [Configuration system](#configuration-system)
- [Security & licensing](#security--licensing)
- [Shared components](#shared-components-key-apis)
- [Coding conventions](#coding-conventions)
- [Code review — mandatory blockers](#code-review--mandatory-blockers)
- [MVBA documentation & pitfalls](#mvba-documentation--pitfalls)
- [Development](#development)

## Architecture constraints (VBA 7.1)

- **Runtime**: single-threaded, event-driven, COM-based. No async/await, no threads, no parallelism — never design around them. "Deferred" work runs on MicroStation **idle events**, not background threads.
- **Platform**: MicroStation CONNECT Edition / OpenCities Map PowerView / Atlas; .NET Framework 4.7.2+.
- **Host**: the MVBA host may be VBA7 (64-bit, `PtrSafe` / `LongLong`) or legacy VBA6 — bitness-sensitive code branches on `#If VBA7`. Win32 handles must be `LongPtr` under VBA7.

## Project structure

| Layer | Modules | Role |
|-------|---------|------|
| **Core/** | `BootLoader.bas` | entry point; owns global state, init order, change-tracking lifecycle |
| | `ARESConstants.bas` | central home for shared compile-time `Public Const` (sentinels, delimiters, geometry factors). **Not** the config vars |
| | `ElementInProcesseClass.cls` | uniqueness queue of element IDs pending processing (global `ElementInProcesse`) |
| | `ErrorHandlerClass.cls` | central logger (global `ErrorHandler`); per-DGN log file with rotation |
| | `ColorDialog.bas` | Win32 color picker ↔ MS color-index helpers |
| **Configuration/** | `ARESConfigClass.cls` | registry of ~33 MS config vars (global `ARESConfig`); versioned import/export |
| | `ARES_MS_VAR_Class.cls` | one config var; assigning `.Value` **write-throughs to MicroStation** |
| | `Config.bas` | thin get/set wrapper over `ActiveWorkspace` config vars; undefined → `ARES_NAVD` |
| | `LangManager.bas` | EN/FR i18n dictionary; `GetTranslation(key, params…)`; `UserLanguage()` |
| **Security/** | `UUID.bas` | machine UUID from MAC + time (not strict RFC-4122); used only by its unit test |
| **EventHandlers/** | `ElementChangeHandler.cls` | `IChangeTrackEvents` — capture add/modify/delete + bulk detection (global `ChangeHandler`) |
| | `IdleEventHandler.cls` | `IEnterIdleEvent` — deferred batch processing |
| | `DGNOpenClose.cls` | `Application` events — re-init change tracking on file open/close |
| | `ReRegisterIdleHandler.cls` | one-shot idle handler that re-attaches change tracking after a bulk suspend |
| **Features/** | `AutoLengths/Auto_Lengths.cls` | write linked-geometry length into a text trigger (+ optional color sync) |
| | `Zoning/Zoning.bas` | buffer zones around elements (`RunZoning` / `RunZoning2`) |
| | `ZoneExport/ExportLengthInRegion.bas` | sum element length inside zones → Excel (`ExportLength`) |
| | `RegionSplit/RegionSplit.bas` + `RegionSplitLocate.cls` | knife-cut a region in two (`SplitRegion`) |
| **Command/** | `Command.bas` | the key-in surface; each key-in inits config/lang, delegates |
| **Components/** | `Geometry`, `Length`, `Link`, `StringsInEl`, `GetElements`, `CustomPropertyHandler`, `MicroStationDefinition`, `MSGraphicalInteraction`, `CellRedreaw`, `FileDialogs` | shared helpers (see [Shared components](#shared-components-key-apis)) |
| **Update/** | `UpdateChecker.bas` | GitHub-releases self-update of every release asset (`.mvba` → C:\ARES, others → C:\ARES\Rsc) via elevated PowerShell + per-asset SHA-256 verify |
| **Tests/** | `UnitTesting.bas` | custom test harness, no native MVBA framework (**deprecated / unmaintained**) |

## Boot & lifecycle

Entry point: **`BootLoader.OnProjectLoad`**

1. `InitializeErrorHandler()` first (everything else logs through it).
2. `InitializeDGNHandlers()` → `New DGNOpenClose` (wires `Application` open/close events).
3. `InitializeInitialIdleHandler()` → an `IdleEventHandler` that sets the caption, initializes translations + config, checks for updates, then removes itself.

**Global objects** (declared in `BootLoader.bas`): `ChangeHandler` (ElementChangeHandler), `ErrorHandler` (ErrorHandlerClass), `ElementInProcesse` (ElementInProcesseClass), `ARESConfig` (ARESConfigClass).

**Teardown** `OnProjectUnload`: reset flags, then `Set … = Nothing` in dependency order (ErrorHandler last).

## Core element-processing flow (Event → Idle → Process)

Two phases. Capture is **synchronous** in the MicroStation event thread; processing is **deferred to idle** so it never blocks modeling.

**Phase 1 — Capture** (`ElementChangeHandler.IChangeTrackEvents_ElementChanged`)

```
If DetectAndSuspendBulkOperation() Then Exit Sub
If Not ARESConfig.IsInitialized Then Exit Sub
If Not IsAnyFeatureEnabled() Then Exit Sub
Select Case Action
  Add / Modify : ShouldQueueElement() -> ElementInProcesse.Add -> EnsureIdleHandlerRegistered
  Delete       : ShouldQueueForDeletion() -> for each linked el: Add -> EnsureIdleHandlerRegistered
```

**Phase 2 — Deferred** (`IdleEventHandler.IEnterIdleEvent_EnterIdle`)

```
mbIsProcessing re-entrance guard
First idle ever -> HandleInitialIdle (caption, translations, config, UpdateChecker) then self-remove
Ensure ARESConfig.Initialize
If ElementInProcesse.HasElements -> ProcessPendingElements   else CleanupHandler
```

**`ChangeHandler.ProcessElement(oElement, Depth=0)`** — `MAX_DEPTH = 20` recursion cap:

- **Branch 1 — text/cell**: if `AUTO_LENGTH And UPDATE_LENGTH`, scan triggers (`ARES_LENGTH_TRIGGER` split by `ARES_VAR_DELIMITER`); RECALCULATE (strip stale numeric between trigger parts) then SEARCH; if found → `AutoLengths.Initialize` + `UpdateLengths`. Else if `ONLY_COLOR` → sync color from linked geometry (FillMode = 2 special case).
- **Branch 2 — geometric**: `Link.GetLink` → for each linked text element: `ElementInProcesse.Add`, then recursive `ProcessElement(linkedEl, Depth + 1)`.

**Idle-handler lifecycle invariants** (hard-won — do not regress):

- ONE persistent `IdleEventHandler` instance, created lazily, only added/removed from the idle list — never `New`-ed per batch (recreating caused ObjPtr reuse + stale-handler desync, so the queue never drained).
- `CleanupHandler` guards queue-drain/flag-reset with `ChangeHandler.IsActiveIdleHandler(Me)` so an orphan (the boot one-shot, or a handler left after `ChangeHandler` was replaced) cannot disturb the live one.
- Real attachment state (`mbChangeTrackingAttached`) is decoupled from the bulk "suspended" flag. `ReinitChangeTrackingForNewFile` is the single source of truth on file open; `AttachChangeTracking` / `DetachChangeTracking` are idempotent.

**Bulk-operation detection** (`DetectAndSuspendBulkOperation`): frequency-based — `ARES_Bulk_Threshold` calls within `ARES_Bulk_Interval` ms → `SuspendAndScheduleResume` (detach tracking, clear queue, register `ReRegisterIdleHandler`, which re-attaches on the next idle). Self-heals an orphaned suspend flag. Suppressed during ARES's own idle writes (`IsIdleProcessingActive`).

## Features

Key-in names and configuration variables are in the [wiki](https://github.com/Asketyll/ARES/wiki); below is the implementation summary.

- **Auto Lengths** (`Auto_Lengths.cls`; runs inside the flow + key-in `ForceUpdateLength`): measures linked geometry (`Link.GetLink` → `Length.GetLength`) and writes the rounded value into the text trigger (e.g. `(Xx_m)` → `(12.3_m)`) via `StringsInEl`. Multiple differing lengths → modeless selection form → `OnElementSelected`. Optional color sync. This is the only feature wired into the live `ElementInProcesse` / `ChangeHandler` pipeline; it scrupulously `Remove`s from the queue on every path.
- **Zoning** (`Zoning.bas`; `RunZoning` = merged / round caps, `RunZoning2` = unmerged / flat caps): offset buffer per element type (Line → stadium, Arc → annular/pie sector, Ellipse → donut, Cell → rotated rounded rect), optional `GetRegionUnion` fuse. Round-vs-flat caps via `CapRoundAt` (flat only at true free ends).
- **Zone Export** (`ExportLengthInRegion.bas`; `ExportLength`): for each element, `Length.GetPartialLengthInsideZones` (length inside zone polygons), aggregated by Style/Level/Color, written to a new `.xlsx`. The Excel COM lifecycle is carefully managed (never quits the user's own session).
- **Region Split** (`RegionSplit.bas` engine + `RegionSplitLocate.cls` single-click driver; `SplitRegion`): cut a Shape/ComplexShape in two from one datapoint, perpendicular (or radial on arcs) across the interior via a thin "knife" rectangle + `GetRegionDifference`. **Anti-destructive invariant**: build, add and validate BOTH halves before deleting the original (`ARES_RegionSplit_Keep_Original` keeps it).

## Configuration system

- `ARESConfig` (global) holds ~33 vars; each `ARES_MS_VAR_Class` write-throughs to MicroStation on `.Value` assignment. First run seeds defaults (`Config.GetVar` == `ARES_NAVD` → write default).
- **Key vs field name differ**: the MS key `ARES_Round` maps to the in-code handle `ARESConfig.ARES_ROUNDS`. Always check both.
- **Import/export** format: `KEY=VALUE:::DEFAULT=…:::MODIFIED=…`, header `# Version: <ARES_CONFIG_VERSION>` (currently `1.0.1`, distinct from the MVBA product release).
- List-typed values are split on `ARES_VAR_DELIMITER` = `|`.
- Shared compile-time constants live in `ARESConstants.bas`; module-private implementation constants (e.g. `MAX_DEPTH`, `EPSILON`, URLs) stay local to their module.

Developer-relevant / automatic vars (not all are surfaced by a key-in):

| Key | Default | Used by |
|-----|---------|---------|
| `ARES_Bulk_Threshold` / `ARES_Bulk_Interval` | 10 / 1000 ms | bulk-operation detection |
| `ARES_Round` | 2 | default rounding for length computations |

The full user-facing list is in the wiki: [Configuration Variables](https://github.com/Asketyll/ARES/wiki/Configuration-Variables).

## Security & licensing

None. ARES has **no licensing or copy protection** (removed) — all features run unconditionally; there is no license file, secret, or COM DLL. The `.mvba` is freely copyable. This was a deliberate decision: client-side protection in an open-source, GitHub-distributed MVBA only deterred a zero-effort copy and was not worth its complexity.

## Shared components (key APIs)

- **Geometry.bas** — pure stateless 2D: `Perp2D(A, B, Dist)` (left-hand perpendicular, zero on degenerate), `NormalizeAngle(delta, direction)` (sweep interval for `CreateArcElement2`). Shared by Zoning + RegionSplit.
- **Length.bas** — `GetLength(El, RND…)`: **Shape/ComplexShape return the LONGEST SIDE, not the perimeter** (for labeling). `GetPartialLengthInsideZones(el, zones())`: ray-cast inside/outside spans; closed zones use the perimeter. `RND = 255` (`ARES_RND_ERROR_VALUE`) is reserved.
- **Link.bas** — `GetLink(El, ReturnMe, FilterByTypes, MaxCount)`: elements in the same graphic group (skips `ARES_DEFAULT_GRAPHIC_GROUP_ID`).
- **StringsInEl.bas** — `GetSetTextsInEl(el, txt, Triggers, Color)`: get/set text in Text/TextNode/Cell. GET on a single `TextElement` returns a **per-character** array (trap). TextNode writes per sub-element (the `TextLine` write property is buggy).
- **GetElements.bas** — `ByEE(Levels, Range, CellName, GraphicGroup, ElTypes, Colors, LineStyles, LineWeights)`: composable `ElementScanCriteria` factory; `GetLevel(name, canCreate)`.
- **CustomPropertyHandler.bas** — ItemType/EC custom properties in one "ARES" library: "Commune" (free text) + "Coupe Type" (value from `ARES_Coupe_Type_List`). MVBA cannot author native dropdowns → the list is config-driven + a UserForm ComboBox.
- **MicroStationDefinition.bas** — `StringToMsdElementType`, `IsValidElementType`, `IsRasterElement`.
- **MSGraphicalInteraction.bas** — `ZoomEl`, `HighlightEl` (transient).
- **CellRedreaw.bas** — rebuilds ATLAS leader-label cell geometry after a text edit (cells in `ARES_Cell_Is_Label_Name`).
- **FileDialogs.bas** — Save/Open dialogs by shelling PowerShell WinForms via a temp `.bat`; config import/export UI.

## Coding conventions

### Module header

Every module starts with this 4-line header (`Module:` for `.bas`, `Class Module:` for `.cls`, `UserForm:` for `.frm`):

```vba
' Module: <name>
' Description: <one line>
' License: This project is licensed under the AGPL-3.0.
' Dependencies: <comma-separated modules, or None>
```

Every module declares `Option Explicit`.

### Naming conventions

**Module/class members** — `m` + type letter (classes use `m*` too, never `p`):

| object | boolean | double | long | longlong | string | integer | byte |
|:------:|:-------:|:------:|:----:|:--------:|:------:|:-------:|:----:|
| `mo` | `mb` | `md` | `ml` | `mll` | `ms` | `mi` | `mby` |

**Locals** — `o` object · `b` boolean · `d` double · `s` string · `n` quantity/index (Long); `i`/`j`/`k` loop counters; `el`/`els` element(s). Any other `Long` takes a meaningful name (`tickNow`, `deltaMs`). **No bare `l` prefix** (avoids `1`/`I` confusion).

- Procedures: `Is…()` / `Has…()` / `Should…()` predicates; `Get…()` / `Set…()`; `Initialize…()` / `Ensure…()` / `Detect…()` / `Process…()`.
- Singleton classes use the `…Class` suffix (`ARESConfigClass`); event-handler classes have no suffix (`ElementChangeHandler`).
- `ARES_` + UPPER_SNAKE for constants and config-var handles; the config handle (`ARES_ROUNDS`) and the MS key string (`ARES_Round`) intentionally differ.

**Accepted exceptions** (do NOT "fix" these): Win32 struct fields in `ColorDialog` (`lStructSize`, `lCustData` — native API layout); a few internal members whose generic names collide as substrings with function/type names (`CellRedreaw.Delta`, `UnitTesting.TestResults`/`TestCount`/`TestElement`).

### Error handling

Standard pattern for every Public procedure:

```vba
Public Function FunctionName() As ReturnType
    On Error GoTo ErrorHandler
    ' ...
    Exit Function
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Module.FunctionName"
End Function
```

- **`HandleError(Description, Number, Source, Optional ModuleName)`** — the 4th arg is the call-site location (`"Module.Proc"`), printed as `[ModuleName]` in the log. `Number = 0` → informational formatting (no `Error N`); `Number > 0` → `Error N (Source)` + a critical MsgBox **in VBA design mode only**. There is **no severity system**.
- **Informational / non-error logs**: call `HandleError "<message>", 0, "", "Module.Proc"` (Number 0, empty Source, location in the ModuleName slot). Do NOT pass a severity word — there is no severity channel.
- **Documented exceptions** (don't "fix"): `ErrorHandlerClass` can't call itself; the `UUID` module is fail-closed/silent; `Command.bas` mostly uses `ShowStatus "<x> failed"`; trivial/pure-math helpers and `Class_Terminate` use `On Error Resume Next`.

### Recurring patterns

1. **Guard / early-exit**: `If oElement Is Nothing Then Exit Sub` / `If Not ARESConfig.IsInitialized Then Exit Sub` / feature-flag checks first.
2. **Find-or-create**: `FindByName` → `If Nothing Then CreateNew`.
3. **Idempotent idle-handler registration**: single persistent instance, add/remove from the idle list only. Anti-GC: hold one-shot handlers (e.g. `moReRegHandler`) as Private members, not `Dim` locals.
4. **Near-origin precision workaround** (Zoning + RegionSplit): region boolean ops (`GetRegionUnion` / `GetRegionDifference`) misbehave at large coordinates — translate operands by `-Range.High`, operate, translate results back.
5. **Config write-through**: assigning `ARES_MS_VAR_Class.Value` persists to MicroStation; `Config.RemoveValue` is a soft delete (sets `""`).
6. **Add-before-symbology**: a new element's Level/symbology can only be set once it is a model member — `AddElement` first, then set Level/Color/Style/Weight, then `Rewrite`.

## Code review — mandatory blockers

A review MUST flag these as **BLOCKER** (not nit):

- Module without `Option Explicit`.
- `Public` Function/Sub without `On Error GoTo ErrorHandler` + the standard `HandleError` (unless a documented exception above).
- Any MVBA type/object/method/property used without verifying its signature in the [MVBA docs](https://github.com/Asketyll/mvba-docs-in-md-for-ai) (the specific page, not just the index).
- `Element.ID` (DLong) used without DLong helpers; an element modified without `.Rewrite`.
- A new geometry feature that deletes/overwrites before both replacement elements are validated and added (anti-destructive invariant).

## MVBA documentation & pitfalls

- **API reference**: a Markdown mirror of the MVBA object model is published at
  https://github.com/Asketyll/mvba-docs-in-md-for-ai (categories `01-types-enums/` … `06-examples/`). In `01-types-enums/`, each type's page lists in its **Remarks** the methods usable with that type — start there, then open the specific method page. Verify any type/object/method/property against it before use.
- **High-frequency pitfalls**:
  - `Element.ID` is a **DLong** — use the DLong helpers, never treat it as a plain Long.
  - An element edited in place is not persisted until `.Rewrite`.
  - `GetSetTextsInEl` GET on a single `TextElement` returns a **per-character** array.
  - Region boolean ops are unstable far from the origin (see the near-origin workaround).
  - Win32 handles must be `LongPtr` under VBA7.

## Development

### Loading in MicroStation

1. Open the MicroStation VBA editor.
2. File → Import File.
3. Import all `.bas`, `.cls`, and `.frm` files, preserving the folder structure.
4. Compile as `ARES.mvba`.

### Dependencies

- Tested on MicroStation CONNECT Edition, OpenCities Map PowerView (Bentley Systems) and Atlas/Eras (Sogelink).
- VBA 7.1 environment.

## License

AGPL-3.0 — see [LICENSE](../LICENSE) for details.

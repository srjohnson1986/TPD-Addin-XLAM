# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A macro-enabled Excel add-in (`TPD_Addin.xlam`) for TPD, written in VBA. There is no compiler/CLI toolchain — "building" means running VBA macros inside Excel itself, and testing means exercising the ribbon commands interactively against sample workbooks. Targets Windows desktop Excel only (locally hosted files; Mac and OneDrive/SharePoint-web are out of scope).

The repo tracks VBA **source** (`/src`), not the compiled `.xlam`. The built add-in is published as a versioned asset on GitHub Releases, never committed.

## Repo layout

```
/src        VBA source of truth (.bas / .cls / .frm / .frx), organized into subfolders matching Rubberduck @Folder tags
/customUI   customUI14.xml (ribbon definition) + ribbon icons — lives outside /src, edited directly
/build      Local build output — gitignored, never committed (holds the base .xlam and build artifacts)
/docs       ARCHITECTURE.md (module-by-module map) and other dev docs
/tools      TPD_Builder.xlsm — the driver workbook that runs BuildAddinFromSource
```

## Development workflow (no CLI build/lint/test — everything happens in Excel/VBE)

Prerequisites: Windows + Excel, "Trust access to the VBA project object model" enabled (File → Options → Trust Center → Macro Settings), and [Rubberduck](https://rubberduckvba.com/) installed matching Office bitness.

**Making a change:**
1. Edit code in the VBE, test interactively in the open workbook.
2. Export to `/src` — either run `ExportAllVBAModules2` (in `modExport_VBAModules`) to refresh the whole tree, or export just the changed component(s) from the VBE Project Explorer (right-click → Export File, overwrite the matching `/src` file).
3. Check `src/export_log.txt` (written by `ExportAllVBAModules2`) to confirm what was exported and where.
4. Review the diff before committing — this is the code review step.

**Rebuilding a testable `.xlam`:**
1. Keep a known-good base file at `build/TPD_Addin_base.xlam` (gitignored — supplies worksheets, ribbon, styles, embedded logo shape that live outside `/src`).
2. From `tools/TPD_Builder.xlsm` (a separate driver workbook, not the add-in itself), run `BuildAddinFromSource` against `/src` and the base file → produces `build/TPD_Addin.xlam`.
3. Load it as an add-in (File → Options → Add-ins → Manage: Excel Add-ins → Browse) and exercise it against the sample EQ List fixtures.
4. If it checks out, it's a release candidate.

**Cutting a release:** bump the version (`PREF_VERSION` / `docProps`), tag `vX.Y.Z`, publish a GitHub Release with the built `.xlam` attached, add a `CHANGELOG.md` entry. That release's `.xlam` becomes the next `build/TPD_Addin_base.xlam`.

## Module organization convention

Every module starts with a Rubberduck folder annotation as a plain VBA comment (the leading `'` is required, or VBA throws a compile error on `@`):

```vb
'@Folder("TPD_Addin.EQList")
Option Explicit
```

This groups modules in Rubberduck's Code Explorer, and the export tooling sorts exported files into matching `/src` subfolders by the same tag. New modules should use an existing folder group, or propose a new one in `docs/ARCHITECTURE.md` if it genuinely doesn't fit.

## Architecture

One custom ribbon tab ("TPD") with two groups — **EQ List Tools** and **Schedule Tools** — defined in `customUI/customUI14.xml`. Ribbon buttons call thin wrapper subs in `modRibbonCallbacks`, which own the global `gRibbon` reference and `RibbonOnLoad` (triggers `modStartup.InitializeAddIn`, ensuring the hidden `_Resources` sheet, embedded logo, and `_Preferences` sheet exist).

Feature areas, mirroring the `@Folder("TPD_Addin.X")` groups (full module-by-module map in `docs/ARCHITECTURE.md` — read it before touching a module you haven't seen):

- **Ribbon** — `modRibbonCallbacks` bridges XML `onAction` callbacks to real entry points.
- **EQList** — "Create Customer EQ List" / "EQ Count" flows (`modMain_CustEQList`, `modMain_CountEquipmentRows`), plus the column-picker UserForm.
- **Schedule** — schedule header insertion; the main schedule automation module is currently a placeholder.
- **SplitExport** — "Split Sheet by Column" and "Save Each Sheet to XLSX" flows, plus their UserForms.
- **Preferences** — key/value settings backed by a hidden `_Preferences` sheet. All keys are centralized in `modPreferences_KeyMap` (`PREF_*` constants) rather than used as loose string literals — follow that pattern for any new preference.
- **Helpers** — shared utilities used by more than one feature area (sheet/workbook ops, header lookup, column filtering, string sanitizing, layout/formatting, checkbox-grid building).
- **Core** — add-in lifecycle/infra: `modPerformance.WithPerformance` (screen updating / events / calc mode wrapper around a routine), `modResources` (embeds default logo), `modStartup` (init sequencing), `modHelpers_Diagnostics` (`SheetExists2`), `modExport_VBAModules` (the export macros, `ExportAllVBAModules` / `ExportAllVBAModules2`).
- **Document** — code-behind for `ThisWorkbook`/`Sheet1`-`3`. These are document modules: unlike standard modules they can't be removed and re-imported normally — the build macro clears and re-pastes their code text instead of a plain `VBComponents.Import`.

## Design intent vs. bugs (don't "fix" these without asking)

- **Global preferences are shared across workbooks, not per-workbook** — "set once, applies everywhere" is intentional.
- **Re-running a command stacks its output** (e.g. running "Create Customer EQ List" twice adds a second sheet) — intentional.

## Known open defects

Tracked as GitHub Issues — [`bug` label](https://github.com/srjohnson1986/TPD-Addin-XLAM/labels/bug). `docs/ARCHITECTURE.md` carries inline pointers on the affected module rows. Two things worth knowing before you touch the relevant code:

- **GEN-04 ([#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7))** — fix `modPerformance.WithPerformance` by *adding* user-facing messaging; do **not** remove the `CleanUp` handler, or session-wide Excel settings (`ScreenUpdating`/`EnableEvents`/`Calculation`) are left broken on error.
- **Planned removal** — `CustEQListColumnPickerForm`, `PREF_CUSTEQ_IMAGE`, and `modHelpers_Image` are slated for removal once the one-click EQ List refactor (saved defaults instead of a per-run picker) lands.

## File-format conventions (`.gitattributes`)

- `.bas` / `.cls` / `.frm` / `.xml` are text, forced to CRLF (matches what the VBE always writes on export — keeps diffs to real code changes).
- `.frx` (UserForm binary resources) and `.xlam`/`.xlsm`/`.xlsx`/`.png`/`.ico`/`.bmp` are binary — never diff/merge/line-ending-convert these.

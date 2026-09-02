# TPD Add-in — Architecture

This is a module-level map of the add-in as of the `v2.2.0` release, organized by the same `@Folder("TPD_Addin.X")` grouping used in Rubberduck's Code Explorer and in the `/src` export layout. For the developer workflow (how to change, export, build, and release), see [CONTRIBUTING.md](../CONTRIBUTING.md).

## Overview

The add-in exposes one custom ribbon tab ("TPD") with two groups: **EQ List Tools** and **Schedule Tools**, defined in `customUI/customUI14.xml`. Ribbon buttons call thin wrapper subs in `modRibbonCallbacks`, which delegate to the feature modules described below. Startup (`modStartup.InitializeAddIn`) runs from `RibbonOnLoad` and makes sure the hidden `_Resources` sheet (embedded logo) and `_Preferences` sheet (saved settings) both exist.

## Naming: "header" vs. "heading"

The word "header" has two unrelated meanings in this codebase, so the naming convention splits them:

- **`header`** — the TPD title block at the top of a sheet (`EQUIPMENT LIST` / Customer / Project / PM / Date / Rev), e.g. `InsertEQHeaderBlock`, `InsertDefaultCustEQHeader`, `InsertDefaultCustScheduleHeader`.
- **`heading`** — the column heading row of the data table, e.g. `headingsRow`, `headingList`, `colHeading`, `GetHeadingList`, `FindHeadingIndex`.

A third case, the logo's positioning anchor row in `modHelpers_Logo`, is neither — it's called `anchorRow` since it can point at either a title-block row or a heading row depending on the caller. Use these names for any new code; don't reintroduce `header`/`headerRow` for column-heading concepts.

## TPD_Addin.Ribbon

| Module | Purpose |
|---|---|
| `modRibbonCallbacks` | Bridges `customUI14.xml`'s `onAction` callbacks to the real entry points (`RunCreateCustEQList`, `RunSplitSheetByColumn`, `RunExportSheetsToXLSX`, etc.). Owns the global `gRibbon` reference and `RibbonOnLoad`, which triggers add-in initialization. |

## TPD_Addin.EQList

| Module | Purpose |
|---|---|
| `modMain_CustEQList` | "Create Customer EQ List" entry point. Shows `CustEQListColumnPickerForm`, copies all rows to a new sheet, deletes unselected columns, formats the sheet, inserts the EQ header block and logo, and freezes panes. |
| `modMain_CustEQListHeader` | "Default EQ List Header" entry point — thin wrapper around `InsertDefaultCustEQHeader`. |
| `modMain_CountEquipmentRows` | "EQ Count" entry point. Inserts an `EQ COUNT` column and numbers every row except those marked `PARENT` or `INCLUDED` in the Purchased column, with leading-zero formatting. Known bug [#6](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/6): blank Purchased values are miscounted as equipment. |
| `CustEQListColumnPickerForm` | UserForm for choosing which columns to keep and the logo image path when creating a Customer EQ List. Builds its checkbox grid via `modHelpers_CheckboxLayout`, applies saved or default column selections, and saves preferences on OK. Known bug [#9](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/9): no validation that at least one column is selected. |

## TPD_Addin.Schedule

| Module | Purpose |
|---|---|
| `modMain_CustSchedule` | Currently an empty placeholder (just the folder tag) reserved for future Schedule automation logic. |
| `modMain_CustScheduleHeader` | "Default Schedule Header" entry point — thin wrapper around `InsertDefaultCustScheduleHeader`. |

## TPD_Addin.SplitExport

| Module | Purpose |
|---|---|
| `modMain_SplitSheets` | "Split Sheet by Column" entry point. Prompts for a group column and columns to keep via `splitSheetByColumnOptionsForm`, then creates or clears one sheet per unique value in the group column and copies matching rows into it. |
| `modMain_ExportSheets` | "Save Each Sheet to XLSX" entry point. Prompts for a filename suffix/date via `frmFilenameOptions`, then saves every worksheet in the workbook to its own `.xlsx` in a per-workbook export folder. Known bug [#8](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/8): hidden sheets (`_Preferences`, `_Resources`) are exported without an opt-in. |
| `modHelpers_Export` | Creates/locates the export folder next to the workbook (`EnsureExportFolder`) and copies a single worksheet out to its own `.xlsx` file (`ExportSheetToXLSX`). |
| `splitSheetByColumnOptionsForm` | UserForm for choosing the group-by column and which columns to keep when splitting a sheet. Remembers the last-used group column and column selection. Known bug [#10](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/10): validates the group column but not that at least one column is selected. |
| `frmFilenameOptions` | Small UserForm for appending free text and/or today's date to exported filenames. |

## TPD_Addin.Preferences

| Module | Purpose |
|---|---|
| `modPreferences` | Key/value preference storage backed by a hidden `_Preferences` worksheet. Also handles saving/loading comma-separated column lists. |
| `modPreferences_KeyMap` | Central list of every preference key (`PREF_CUSTEQ_IMAGE`, `PREF_SPLIT_GROUPCOL`, `PREF_EXPORT_APPEND`, etc.) so keys never float as loose string literals. |
| `modPreferences_Initializer` | Seeds default preference values the first time the add-in initializes on a workbook. |
| `frmSetTPDDefaults` | The in-progress unified defaults UserForm (planned tabs: Header Logo / Customer EQ List / Split Sheet / Export). Currently loads the embedded default logo into a preview image control on activation. |

## TPD_Addin.Helpers

Shared utilities used across more than one feature area.

| Module | Purpose |
|---|---|
| `modHelpers_Workbook` | Sheet/workbook utilities: first visible sheet, create-or-clear a sheet by name, next available sheet name, last-row/last-column lookups. |
| `modHelpers_Headers` | Reads a worksheet's heading row into an array; finds a heading's column index (searches right-to-left so the real "Vendor" column wins over decoy columns). |
| `modHelpers_Columns` | Collection-membership helper; deletes columns not in a selected-headings list; gets unique values in a column; copies filtered rows across sheets by column selection; copies all rows verbatim. |
| `modHelpers_Strings` | String cleanup/sanitizing helpers (sheet names, file names, cell value cleanup) used throughout the other modules. |
| `modHelpers_SheetSetup` | Orchestrates default header construction: inserts the EQ List header block, and inserts the default EQ List / Schedule headers (calling into `modHelpers_SheetFormatting` and `modHelpers_Logo`). Also owns `InsertRows` and `GetTodaysDate`. |
| `modHelpers_SheetFormatting` | Formats EQ/split sheets (borders, alignment, fonts), freezes panes, and autofits used columns. |
| `modHelpers_Logo` | Positions and resizes the logo shape: default-logo insertion/alignment, right-aligned logo insertion, and resizing a picture to fit a max row count. |
| `modHelpers_Forms` | Thin bridge functions that show `CustEQListColumnPickerForm` / `frmFilenameOptions` / `splitSheetByColumnOptionsForm` and hand back the user's selections to the caller. |
| `modHelpers_Image` | `SafeInsertLogoAtRight` (error-wrapped logo insertion) and `PastePicture` (grabs an image off the clipboard). Flagged as now-unused since logo handling moved to the embedded-shape approach on `_Resources`. |
| `modHelpers_CheckboxLayout` | Dynamically builds and arranges the checkbox grid shared by both column-picker forms. Known bug [#12](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/12): the fixed 3-column grid clips beyond ~30 headings. |
| `modHelpers_CheckboxSelection` | Reads which checkboxes are checked into a `Collection`, and re-checks boxes matching a previously saved column list. |

## TPD_Addin.Core

Add-in lifecycle and infrastructure — not directly user-facing.

| Module | Purpose |
|---|---|
| `modPerformance` | `WithPerformance` wrapper: disables `ScreenUpdating`/`EnableEvents` and sets manual calculation around a named internal routine, then restores them afterward. Known bug [#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7): errors inside the wrapped routine are swallowed without a user-facing message — the fix must keep the `CleanUp` handler (see Known open items). |
| `modResources` | Creates the hidden `_Resources` sheet and embeds the default logo shape on first run. |
| `modStartup` | `InitializeAddIn`, called from `RibbonOnLoad`; ensures both `_Resources` and `_Preferences` exist before anything else runs. |
| `modHelpers_Diagnostics` | Houses `SheetExists2` (known bug CORE-01 / [#11](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/11): its body assigns to `SheetExists`, not `SheetExists2`, so it always returns the default `False`). |
| `modExport_VBAModules` | The VBA export macros (`ExportAllVBAModules` / `ExportAllVBAModules2`) that support the source-control workflow itself. |

## TPD_Addin.Document

Document modules — code-behind tied to the workbook/sheet objects rather than standalone modules. These can't be removed and re-imported the way standard modules can (see CONTRIBUTING.md).

| Module | Purpose |
|---|---|
| `ThisWorkbook` | Workbook-level code-behind. Currently minimal. |
| `Sheet1`, `Sheet2`, `Sheet3` | Sheet-level code-behind. `Sheet2` has an empty `Worksheet_SelectionChange` stub; the others are currently empty. |

## Known open items

Open defects are tracked as GitHub Issues — see the [`bug` label](https://github.com/srjohnson1986/TPD-Addin-XLAM/labels/bug). The affected module rows above carry an inline pointer to the relevant issue. Historical defect IDs map to issues as: EQC-10 → [#6](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/6), GEN-04 → [#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7), EXP-07 → [#8](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/8), CEQ-06 → [#9](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/9), SPL-11 → [#10](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/10), CORE-01 → [#11](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/11), UI-01 → [#12](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/12).

Non-issue guidance to keep in mind when working these areas:

- **GEN-04 ([#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7)) constraint** — the fix is to add user-facing messaging to `modPerformance.WithPerformance`, *not* to remove the `CleanUp` handler. Removing it leaves `ScreenUpdating`/`EnableEvents`/`Calculation` broken for the rest of the Excel session on any error.
- **Planned removal** — `CustEQListColumnPickerForm`, `PREF_CUSTEQ_IMAGE`, and `modHelpers_Image` are slated for removal once the one-click EQ List refactor (using saved defaults instead of a per-run picker) lands.

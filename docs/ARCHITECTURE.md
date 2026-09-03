# TPD Add-in — Architecture

This is a module-level map of the add-in as of the `v2.3.0` release, organized by the same `@Folder("TPD_Addin.X")` grouping used in Rubberduck's Code Explorer and in the `/src` export layout. For the developer workflow (how to change, export, build, and release), see [CONTRIBUTING.md](../CONTRIBUTING.md).

## Overview

The add-in exposes one custom ribbon tab ("TPD") with two groups: **EQ List Tools** and **Schedule Tools**, defined in `customUI/customUI14.xml`. Ribbon buttons call thin wrapper subs in `modRibbonCallbacks`, which delegate to the feature modules described below. Startup (`modStartup.InitializeAddIn`) runs from `RibbonOnLoad` and stamps the running version into the registry-backed preference store. Saved settings live in the Windows registry per user (`modPreferences`), not in the workbook. The hidden `_Resources` sheet and its embedded default-logo shape come from the base `.xlam`, not from code.

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
| `modMain_CountEquipmentRows` | "EQ Count" entry point. `RunCountEquipmentRows` → `WithPerformance PERF_COUNT_EQ_ROWS` → `CountEquipmentRows_Internal` (validates the active worksheet via `RequireActiveWorkbook`) → `NumberEquipmentRows` inserts an `EQ COUNT` column and numbers every row except those marked `PARENT` or `INCLUDED` in the Purchased column, with leading-zero formatting sized to the highest number. A row with a blank Purchased value is a real equipment line whose status isn't filled in yet, so it is numbered by design ([#6](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/6), closed as working-as-intended). |
| `CustEQListColumnPickerForm` | UserForm for choosing which columns to keep and the logo image path when creating a Customer EQ List. `LoadColumns` builds the checkbox grid (`modHelpers_CheckboxLayout`) then applies the saved-or-default selection (`ApplyColumnSelection`, shared with the split picker — [#81](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/81)); OK saves preferences. `cmdOK_Click` does **not** check that at least one column is selected. Slated for removal in the one-click EQ List refactor ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25)). |

## TPD_Addin.Schedule

| Module | Purpose |
|---|---|
| `modMain_CustScheduleHeader` | "Default Schedule Header" entry point — thin wrapper around `InsertDefaultCustScheduleHeader`. |

The main Schedule-automation module doesn't exist yet — the empty `modMain_CustSchedule` placeholder was removed in [#48](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/48); recreate it with a `@Folder("TPD_Addin.Schedule")` tag when that work starts.

## TPD_Addin.SplitExport

| Module | Purpose |
|---|---|
| `modMain_SplitSheets` | "Split Sheet by Column" entry point. Prompts for a group column and columns to keep via `splitSheetByColumnOptionsForm`, then creates or clears one sheet per unique value in the group column and copies matching rows into it. |
| `modMain_ExportSheets` | "Save Each Sheet to XLSX" entry point. Prompts for a filename suffix/date via `frmFilenameOptions`, then saves each eligible worksheet to its own `.xlsx` in a per-workbook export folder. `ShouldExportSheet` skips `_`-prefixed internal sheets and very-hidden sheets always, and plain-hidden sheets unless the source-level `EXPORT_INCLUDE_HIDDEN_SHEETS` flag is turned on ([#8](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/8)). Per-sheet failures are collected and reported in one summary rather than aborting the run ([#15](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/15)). |
| `modHelpers_Export` | Creates/locates the export folder next to the workbook, named after the workbook file (`EnsureExportFolder`; returns `""` after a message if the workbook is unsaved or the folder can't be made), and copies a single worksheet out to its own `.xlsx` file (`ExportSheetToXLSX`, returns `""` on success or a reason string on failure). If a sheet's target name collides with an already-open workbook it gets a ` (export)` suffix; an existing target file on disk is overwritten without prompting ([#64](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/64)); any `SaveAs` failure is caught, the temporary `ws.Copy` workbook is closed without saving, and the reason is returned ([#15](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/15)). `ReportBatchOutcome` is the shared end-of-run summary box ("N done" + a list of `'name': reason` failures) used by both the Split and Export flows. |
| `splitSheetByColumnOptionsForm` | UserForm for choosing the group-by column and which columns to keep when splitting a sheet. `LoadColumns` builds the grid, applies the saved-or-default column selection (`ApplyColumnSelection`), populates the group dropdown, then `SelectGroupColumn` picks the start value (saved preference → "Vendor" → first column). Remembers the last-used group column and column selection. `cmdOK_Click` validates both the group column and (via `HasColumnSelection`) that at least one column is checked ([#10](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/10)). |
| `frmFilenameOptions` | Small UserForm for appending free text and/or today's date to exported filenames. |

## TPD_Addin.Preferences

| Module | Purpose |
|---|---|
| `modPreferences` | Per-user key/value preference storage in the Windows registry via `SaveSetting`/`GetSetting` (`HKCU\Software\VB and VBA Program Settings\TPD_Addin\Preferences`). Registry rather than a sheet inside the `.xlam` because Excel never saves an add-in on exit, so in-workbook writes were silently lost on quit ([#84](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/84)). `SavePref`/`LoadPref`/`DeletePref` plus `SaveColumnList`/`LoadColumnList` for comma-separated column lists; a missing key returns the caller's default. |
| `modPreferences_KeyMap` | Central list of every preference key (`PREF_CUSTEQ_IMAGE`, `PREF_SPLIT_GROUPCOL`, `PREF_EXPORT_APPEND`, etc.) so keys never float as loose string literals. |
| `modPreferences_Initializer` | `InitializePreferences` — stamps `PREF_VERSION` with `ADDIN_VERSION` when it changes (support/diagnostics). No value seeding: `LoadPref` returns the caller's default for any unset key, and each picker falls back to its own built-in column list. |
| `frmSetTPDDefaults` | The intended "Set TPD Defaults" UserForm ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25)). **Currently unreachable** — no ribbon button opens it, and its only code is a `UserForm_Activate` that previews the embedded logo. To be built out with tabs for Customer EQ List / Split Sheet / Customer Schedule / Header Logo. |

## TPD_Addin.Helpers

Shared utilities used across more than one feature area.

| Module | Purpose |
|---|---|
| `modHelpers_Workbook` | Sheet/workbook utilities: `RequireActiveWorkbook` (the standard "bail with a message if no workbook is open" guard), `GetFirstVisibleSheet` (guards no-workbook / no-visible-sheet, returns `Nothing`), create-or-clear a sheet by name, next available sheet name, last-row (`GetLastRow`, safe on an empty sheet) / last-column lookups, and `IsWorkbookOpen(fileName)`. |
| `modHelpers_Headers` | Reads a worksheet's heading row into an array; finds a heading's column index. `FindHeadingIndex` defaults to a right-to-left scan so the authoritative "Vendor" column wins over decoy columns; pass `preferRightmost:=False` for a plain left-to-right lookup ([#45](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/45)). |
| `modHelpers_Columns` | Collection-membership helper; deletes columns not in a selected-headings list; gets unique values in a column; `CopyFilteredRowsByColumns` copies matching rows across sheets projected onto a column selection (data only — formatting is the caller's job); `CopyEntireSheetRows` copies every row verbatim (whole-row copy carries row grouping across). |
| `modHelpers_Strings` | String sanitizing helpers: `SanitizeSheetName`, `SanitizeFileText`, and `NormalizeCellText` (first physical line, NBSP→space, collapsed spaces, trimmed — used wherever cell values are compared or deduped). |
| `modHelpers_SheetSetup` | Orchestrates default header construction: inserts the EQ List header block, and inserts the default EQ List / Schedule headers (calling into `modHelpers_SheetFormatting` and `modHelpers_Logo`). Owns `EQ_HEADER_ROW_COUNT` / `SCHEDULE_HEADER_ROW_COUNT`, `InsertRows`, and `GetTodaysDate`. |
| `modHelpers_SheetFormatting` | Formats EQ/split sheets (borders, alignment, fonts), freezes panes, and autofits used columns. |
| `modHelpers_Logo` | Positions and resizes the logo shape: default-logo insertion/alignment (`InsertDefaultLogo`) and resizing a picture to fit a max row count. Also `PastePicture` (clipboard image → `StdPicture`, used by `frmSetTPDDefaults` for the logo preview) and the file-path pair `InsertLogoAtRight` / `SafeInsertLogoAtRight` (called by the Customer EQ List flow — usually with an empty path now); the file-path pair is legacy of the file-path logo approach and slated for removal with the one-click EQ List refactor ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25)). |
| `modHelpers_CheckboxLayout` | Builds the checkbox grid shared by both column-picker forms — one box per heading, 3 columns filled top-to-bottom. When there are more headings than fit in 3 columns at the preferred height, the columns grow taller (never a 4th column off the right edge) and the frame gets a vertical scrollbar so every box stays reachable ([#47](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/47)). |
| `modHelpers_CheckboxSelection` | Reads which checkboxes are checked into a `Collection` (`GetSelectedColumns`), tests whether any are checked (`HasColumnSelection`), re-checks boxes matching a previously saved column list (`AutoSelectColumns`), and `ApplyColumnSelection` — the shared "check the saved list, else the built-in defaults" step both column pickers run after building their grid ([#81](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/81)). |

## TPD_Addin.Core

Add-in lifecycle and infrastructure — not directly user-facing.

| Module | Purpose |
|---|---|
| `modPerformance` | `WithPerformance` wrapper: disables `ScreenUpdating`/`EnableEvents` and sets manual calculation around a named internal routine, then restores them afterward. On an error inside the wrapped routine, `CleanUp` still restores all three settings and then surfaces `Err.Description` in a `MsgBox` ([#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7)). The routine is selected by a `PERF_*` string constant passed by the ribbon callback, so a mistyped name fails to compile rather than hitting the `Unknown action` branch at run time ([#36](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/36)). |
| `modStartup` | `ADDIN_VERSION` constant + `InitializeAddIn`, called from `RibbonOnLoad`; runs `InitializePreferences` (registry version stamp). (`_Resources` + the embedded logo come from the base `.xlam`; the code-side fallback `modResources` was removed in [#62](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/62).) |
| `modExport_VBAModules` | The VBA export macro `ExportAllVBAModules` that supports the source-control workflow itself. |

For "does this sheet exist in the add-in", use `SheetExists(ThisWorkbook, name)` from `modHelpers_Workbook` (the old single-purpose `modHelpers_Diagnostics.SheetExists2` was removed in [#50](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/50) — no callers, and it did nothing `SheetExists` didn't).

## TPD_Addin.Document

Document modules — code-behind tied to the workbook/sheet objects rather than standalone modules. These can't be removed and re-imported the way standard modules can (see CONTRIBUTING.md).

| Module | Purpose |
|---|---|
| `ThisWorkbook` | Workbook-level code-behind. Currently minimal. |
| `Sheet1`, `Sheet2`, `Sheet3` | Sheet-level code-behind. `Sheet2` has an empty `Worksheet_SelectionChange` stub; the others are currently empty. |

## Known open work

Tracked as GitHub Issues — no open `bug`-labelled issues. One open piece of work:

- **Set TPD Defaults dialog ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25))** — build out `frmSetTPDDefaults` (currently unreachable) into the settings UI: a ribbon button opens it; tabs for Customer EQ List / Split Sheet / Customer Schedule / Header Logo write the `PREF_*` keys. Feature flows then read prefs and run one-click (built-in lists as the fallback). `CustEQListColumnPickerForm`, `PREF_CUSTEQ_IMAGE`, and the `InsertLogoAtRight` / `SafeInsertLogoAtRight` file-path pair in `modHelpers_Logo` are removed as part of it. Future enhancement.

### Regressions to guard against (all fixed — don't undo them)

- **`modPerformance.WithPerformance` ([#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7) / [#36](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/36))** — the `CleanUp` handler must **always** restore `ScreenUpdating` / `EnableEvents` / `Calculation`, or they stay broken for the Excel session on any error. Callers pass a `PERF_*` constant, not a bare string.
- **Never call a `Sub` as `MySub (obj)` ([#29](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/29))** — the parens force by-value evaluation (the default property), so a `Worksheet`/`Workbook` argument raises run-time 438. Use `MySub obj` or `Call MySub(obj)`. The `Split Sheet by Column` loop ([#32](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/32)) now reports per-value failures in its summary, so a regression there is at least visible.

For the full history of fixed defects (EQC-10, GEN-04, EXP-07, SPL-11, CORE-01, EXP-08, CEQ-06, UI-01, #29–#68), see `CHANGELOG.md`.

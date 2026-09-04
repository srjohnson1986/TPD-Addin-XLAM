# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project aims to follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Releases are published on the [Releases page](https://github.com/srjohnson1986/TPD-Addin-XLAM/releases)
with the built `TPD_Addin.xlam` attached as an asset. See
[CONTRIBUTING.md](CONTRIBUTING.md) for the release process.

## [Unreleased]

- The Customer EQ List logo is now the embedded TPD logo, top-right of the
  heading row and scaled to the header block — the same
  `InsertDefaultLogo … "right-top"` call the "Default EQ List Header" command
  already uses ([#82](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/82)).
  Previously "Create Customer EQ List" inserted a user-chosen image file at its
  natural size (or nothing, if no path was set). The picker's logo-path field
  and Browse button are now unused controls with no code behind them — delete
  them in the VBE, or let #25 (which removes the whole form) take them. Removed
  the `InsertLogoAtRight` / `SafeInsertLogoAtRight` file-path pair (their
  right-align math was already duplicated inside `InsertDefaultLogo`) and the
  `PREF_CUSTEQ_IMAGE` preference key.
- Consolidated the two column-picker forms ([#81](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/81)),
  no behavior change: the "check the user's saved list, else the built-in
  defaults" step is now one shared `modHelpers_CheckboxSelection.ApplyColumnSelection`
  instead of a copy in each `LoadColumns`. Dropped the two unused optional
  parameters on `LayoutCheckboxes`, and removed a wasted `AutoSelectColumns`
  call each form made from `UserForm_Initialize` before its checkboxes existed.
  The split picker's group-column default (saved → "Vendor" → first) moved
  into one `SelectGroupColumn` helper that runs after the dropdown is filled,
  instead of being split across `UserForm_Initialize` and `LoadColumns`.
- Preferences (column-picker selections, Split group column, logo path) now
  persist across Excel sessions ([#84](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/84)).
  They were stored in a hidden `_Preferences` sheet **inside `TPD_Addin.xlam`**,
  but Excel never saves an add-in on exit, so every selection was discarded on
  quit — it only appeared to stick within a session. `modPreferences` now uses
  the Windows registry (`SaveSetting`/`GetSetting` under
  `HKCU\Software\VB and VBA Program Settings\TPD_Addin\Preferences`): always
  writable, written immediately, per Windows user, independent of which
  workbook is open. A key that was never saved falls back to the caller's
  default, and each picker to its built-in column list, so a fresh machine
  still opens with sensible defaults. The `_Preferences` sheet, `PrefSheet` /
  `GetPrefSheet`, and the default-value seeding in `modPreferences_Initializer`
  are all gone; `SaveColumnList` no longer crashes on an empty selection.
- Customer EQ List column picker now restores a saved non-default column
  selection instead of always reverting to the built-in defaults
  ([#79](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/79)) — `LoadColumns`
  read the preference with a string literal instead of the `PREF_CUSTEQ_COLUMNS`
  key constant, so it never found what OK had saved.
- Routed duplicated code through the existing shared helpers
  ([#80](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/80)): the EQ Count
  "PURCHASED" lookup, the inline last-column scans in the formatting and logo
  helpers, and the Split Sheet unique-value pass now use `FindHeadingIndex` /
  `GetLastCol` / `CollectionContainsText` / `GetUniqueValuesInColumn`; the Split
  and Export end-of-run summaries share one `ReportBatchOutcome`. No user-visible
  change, except the "PURCHASED" heading match no longer trims surrounding
  whitespace (it was already case-insensitive).
- Merged `modHelpers_Image` into `modHelpers_Logo` ([#77](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/77)) — both its
  functions (`SafeInsertLogoAtRight`, `PastePicture`) were logo-only, so a
  separate module earned nothing. No behavior change; callers use unqualified
  names. Retiring the `InsertLogoAtRight` / `SafeInsertLogoAtRight` file-path
  pair itself stays with the one-click EQ List refactor ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25)).
- `build/_base/TPD_Addin_base.xlam` is no longer tracked — the `.gitignore`
  `/build/*.xlam` rule missed the `_base/` subdir, so it had been committed
  since the initial `/build` setup despite the docs saying otherwise. Now
  `git rm --cached`'d and covered by `/build/**/*.xlam`; the base is a local
  per-contributor prerequisite (see CONTRIBUTING.md).
- Corrected the "cutting a release" docs (`CONTRIBUTING.md`, `CLAUDE.md`): the
  base file must be **stripped of standard modules and UserForms** — the
  builder can't re-import a form that already exists, so a full add-in `.xlam`
  is not a valid base. The old "that release's `.xlam` becomes the next base"
  line was wrong. Also ignore the per-component `.log` files the builder drops
  in `/src` on an import error.

## [2.3.0] - 2026-09-03

The first release cut from the `/src` VBA source (v2.2.0 and earlier were built
by hand). Almost entirely internal hardening, cleanup, and legibility — a few
user-visible behavior changes are noted under **Changed** / **Fixed**.

### Added

- `tools/Build-TPDAddin.ps1` — drives `TPD_Builder.xlsm`'s `BuildAddin` macro
  headlessly to rebuild `build/TPD_Addin.xlam` from `/src`. Paths default off the
  script's own location, so `powershell -ExecutionPolicy Bypass -File tools\Build-TPDAddin.ps1`
  works from any clone.

### Removed

- Four dead / redundant modules:
  `modHelpers_Diagnostics` (its one function `SheetExists2` had no callers and
  duplicated `SheetExists(ThisWorkbook, …)`),
  the empty `modMain_CustSchedule` placeholder
  ([#50](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/50),
  [#48](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/48)),
  `modHelpers_Forms` (three "show a picker form" bridges, none called), and
  `modResources` (`CreateResourcesSheetIfMissing` only ever no-op'd and its
  fallback loaded a logo from a hardcoded `C:\dev\` path)
  ([#61](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/61),
  [#62](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/62)).
  `_Resources` + the embedded logo now come purely from the base `.xlam`.
- Commented-out code across `/src` (a stale `CreateOrClearSheet` copy, an old
  `GetLastRow` overload, `'MsgBox "Ribbon callback fired"` debug lines) and the
  superseded `ExportAllVBAModules` v1 (hardcoded `C:\VBA_Export\` path);
  `ExportAllVBAModules2` renamed to `ExportAllVBAModules`.

### Changed

- Repo now tracks the VBA source under `/src` rather than the compiled `.xlam`;
  builds are published as versioned Releases with the `.xlam` attached.
- New `ADDIN_VERSION` constant in `modStartup` (`"2.3.0"`), stamped into
  `_Preferences` as `PREF_VERSION` by `modPreferences_Initializer` when it
  changes — replaces the vestigial hardcoded `"1.0"` seed.
- The seven informally-tracked defects were filed as GitHub Issues and the
  defect lists in `CLAUDE.md` / `docs/ARCHITECTURE.md` reduced to pointers.
- Final legibility sweep: removed a stray `Debug.Print "MATCHED DEFAULT: …"`
  from `splitSheetByColumnOptionsForm.LoadColumns`, replaced two `?`-for-arrow
  comments in the same sub with plain English, and corrected the
  `modHelpers_Image` banner / `ARCHITECTURE.md` row — both its functions do have
  live callers (`SafeInsertLogoAtRight` from the EQ List flow, `PastePicture`
  from `frmSetTPDDefaults`), they're just legacy of the file-path logo approach.
- Stripped trailing whitespace from 5 whitespace-only blank lines
  ([#63](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/63), partial):
  `modMain_CustEQList`, `modHelpers_Columns` (×2), `splitSheetByColumnOptionsForm`
  (×2). `git diff --ignore-all-space` is empty — no code change. The broader
  Rubberduck indent normalization is still open on #63.
- Review loose ends ([#58](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/58)),
  no behavior change: `ExportSheets_Internal` uses the shared
  `RequireActiveWorkbook` guard instead of its own inline check; new
  `SCHEDULE_HEADER_ROW_COUNT` constant (parallel to `EQ_HEADER_ROW_COUNT`);
  `WithPerformance`'s `action` param → `workProcName`,
  `CopyFilteredRowsByColumns`'s `idx` → `srcColIndex`, `modHelpers_Logo`'s
  `origW`/`origH` → `naturalWidth`/`naturalHeight`; `Option Explicit` added to
  `modResources`; module-purpose banners added to the remaining modules that
  lacked one (`modHelpers_Forms` — flagged as currently unused — `modHelpers_Headers`,
  `modHelpers_Strings`, `modHelpers_Export`, `modRibbonCallbacks`, `modResources`,
  the main flow and header-wrapper modules).
- Helper renames ([#49](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/49)),
  no behavior change: `modHelpers_Strings.CleanValue` → `NormalizeCellText` (the
  old name badly undersold what it does), `modHelpers_Columns.CopyAllRowsPreserveGroups`
  → `CopyEntireSheetRows` (with a comment noting the whole-row copy is what carries
  row grouping across). Both got a summary comment.
- `CopyFilteredRowsByColumns` ([#44](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/44)),
  no behavior change: the row-copy helper no longer bolds the heading row,
  activates/selects the sheet, freezes panes, or autofits — all of which its
  caller (`SplitOneGroup` → `FormatSplitSheet` + `SafeFreezePanes`) already did,
  so per group value the sheet was frozen and autofitted twice and left
  selected. Heading-row bold moved into `FormatSplitSheet`.
- `FindHeadingIndex` ([#45](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/45)),
  no behavior change: the right-to-left scan is now an explicit
  `Optional preferRightmost As Boolean = True` parameter with a comment
  explaining the duplicate-"Vendor"-column reason, instead of a silent default.
  `SplitSheetByColumn_DoWork` passes `preferRightmost:=True` at the call site.
- `WithPerformance` dispatch ([#36](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/36)),
  no behavior change: ribbon callbacks now pass a `PERF_*` string constant
  (declared in `modPerformance`) instead of a bare literal, and the `Select Case`
  matches on the same constants. A mistyped routine name is now a compile error
  at the call site rather than a run-time `"Unknown action"` message.
- Code-hygiene sweep ([#37](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/37)),
  no behavior change: added `Option Explicit` to `modStartup`; fixed a mojibake
  comment in `modHelpers_Strings`; gave `GetLastRow` an explicit `Public` and a
  guard so it returns row 1 instead of raising 91 on a completely empty sheet
  (also closes [#33](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/33) —
  `FormatEQSheet` now calls `GetLastRow` instead of its own inline `Find`);
  introduced `EQ_HEADER_ROW_COUNT` in `modHelpers_SheetSetup` to replace three
  hand-synced copies of the header-block row count; added a shared
  `RequireActiveWorkbook` helper to `modHelpers_Workbook`; collapsed the three
  `_Preferences` sheet accessors down to `GetPrefSheet`; and added module-purpose
  banners to the sheet/column/logo helpers.

### Fixed

- ([#64](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/64)):
  "Save Each Sheet to XLSX" re-run into a folder that already has the files no
  longer pops Excel's "replace existing file?" dialog once per sheet.
  `ExportSheetToXLSX` suppresses `Application.DisplayAlerts` around the `SaveAs`
  (restored immediately after, and in the failure handler); re-exporting
  silently replaces the prior files.
- ([#57](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/57)):
  "Split Sheet by Column" no longer loses rows when two distinct group values
  sanitize to the same sheet name (shared first 31 chars, or differing only in
  characters Excel forbids in sheet names). `SplitSheetByColumn_DoWork` tracks
  the names created this run and `SplitOneGroup` routes each through
  `UniqueRunName`, which appends ` (2)` / ` (3)` / … (trimming the base to stay
  within 31 chars). Re-running the command still reuses/clears sheets by name
  as before.
- ([#47](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/47)):
  The column-picker checkbox grid no longer clips past ~30 headings. It was
  a fixed 3-column layout that spilled a 4th (and 5th…) column off the right
  edge of a fixed frame; `LayoutCheckboxes` now caps at 3 columns, grows them
  taller as needed, and turns on a vertical scrollbar when the grid is taller
  than the frame. (This is the surviving half of the old UI-01.)
- ([#46](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/46)):
  "Save Each Sheet to XLSX" names its export folder after the workbook file
  (minus extension) instead of `Worksheets(1).Name` — the leftmost tab, which
  could be an internal sheet or shift as tabs are reordered. `MkDir` is now
  wrapped: a failure shows "Couldn't create the export folder: …" and the run
  stops cleanly instead of throwing.
- ([#35](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/35)):
  "EQ Count" now runs inside `WithPerformance` (`PERF_COUNT_EQ_ROWS`) so it no
  longer flickers or leaves a half-inserted `EQ COUNT` column with screen
  updating / calculation stuck if it errors partway. `CountEquipmentRows_Internal`
  validates the active worksheet via `RequireActiveWorkbook`; the numbering work
  moved to `NumberEquipmentRows` with `FindPurchasedColumn` split out. The dead
  `ws` parameter (the old sub overwrote it with `ActiveSheet`) is gone, the
  post-insert `purchasedCol` index is rebased once instead of `+ 1` at each use,
  and a heading-only sheet exits cleanly instead of formatting an empty range.
- ([#34](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/34)):
  "Create Customer EQ List" and "Split Sheet by Column" no longer raise a bare
  error 91 when run with no workbook open or from a workbook with no visible
  sheet. `GetFirstVisibleSheet` now guards both cases via `RequireActiveWorkbook`,
  shows a plain message, and returns `Nothing`; both callers check for `Nothing`
  and exit.
- ([#32](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/32)):
  "Split Sheet by Column" no longer reports `"N sheets created successfully."`
  when some group values failed. Per-value work moved into `SplitOneGroup`;
  `SplitSheetByColumn_DoWork` runs it under inline error handling, collects any
  failure as a `'value': reason` string, and `ReportSplitOutcome` shows a
  created/failed summary (same pattern as `ExportSheets_DoWork`). The old
  `LoopError` / `Debug.Print` / `Resume Next` handler is gone.
- ([#29](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/29)):
  "Split Sheet by Column" raised a silent run-time error 438 once per group
  value — `FormatSplitSheet` called `AutoFitUsedColumns (ws)`, whose stray
  parentheses forced VBA to evaluate the `Worksheet` by value (no default
  property). The `LoopError` handler logged it and continued, so output looked
  fine but the final autofit and freeze-panes were skipped. Removed the
  parentheses.
- GEN-04 ([#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7)):
  `modPerformance.WithPerformance` no longer swallows errors silently. On failure
  in a wrapped routine it still restores `ScreenUpdating`/`EnableEvents`/`Calculation`,
  then shows the error description in a message box.
- CORE-01 ([#11](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/11)):
  `modHelpers_Diagnostics.SheetExists2` assigned its result to the wrong
  identifier and always returned `False`; it now reports correctly.
- SPL-11 ([#10](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/10)):
  "Split Sheet by Column" now warns and keeps the form open if OK is clicked
  with no columns checked, instead of proceeding with an empty selection. Adds
  `HasColumnSelection` to `modHelpers_CheckboxSelection`.
- EXP-07 ([#8](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/8)):
  "Save Each Sheet to XLSX" no longer exports hidden sheets — the add-in's
  internal `_Preferences` / `_Resources` sheets, very-hidden sheets, or
  plain-hidden sheets. A source-level `EXPORT_INCLUDE_HIDDEN_SHEETS` flag can
  re-enable plain-hidden sheets if a build ever needs them.
- EXP-08 ([#15](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/15)):
  "Save Each Sheet to XLSX" with an empty filename suffix no longer fails and
  leaves an orphaned workbook open. `ExportSheetToXLSX` catches `SaveAs`
  failures, closes the temporary copy without saving, and reports which sheets
  failed and why in a single summary. A target name that collides with an
  already-open workbook gets a ` (export)` suffix.

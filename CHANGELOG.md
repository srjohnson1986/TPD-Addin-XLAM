# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project aims to follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Releases are published on the [Releases page](https://github.com/srjohnson1986/TPD-Addin-XLAM/releases)
with the built `TPD_Addin.xlam` attached as an asset. See
[CONTRIBUTING.md](CONTRIBUTING.md) for the release process.

## [Unreleased]

- Migrating the repo from tracking the compiled `.xlam` to tracking its VBA
  source under `/src`, with builds published as versioned Releases.
- Filed the seven informally-tracked defects as GitHub Issues (#6–#12) and
  reduced the defect lists in `CLAUDE.md` and `docs/ARCHITECTURE.md` to
  pointers, with inline issue links on the affected module rows.
- Filed EXP-08 ([#15](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/15)):
  "Save Each Sheet to XLSX" fails and leaves an orphaned workbook open when run
  with an empty filename suffix.
- Closed EQC-10 ([#6](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/6)) as
  working-as-intended: a blank Purchased value marks a real equipment line whose
  status isn't filled in yet, so "EQ Count" numbers it by design. Only `PARENT`
  and `INCLUDED` rows are skipped.
- Closed CEQ-06 ([#9](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/9)) and
  UI-01 ([#12](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/12)), folding
  both into a tracking issue for the one-click "Create Customer EQ List" refactor
  ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25)) — moving column
  and logo settings into saved defaults removes `CustEQListColumnPickerForm`, which
  subsumes both. Trimmed the now-resolved defect IDs out of the `CLAUDE.md` and
  `docs/ARCHITECTURE.md` "known defects" sections; no open `bug` issues remain.
- Filed a second review pass as issues [#32](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/32)–[#37](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/37):
  Split Sheet swallowing per-group failures (#32), `Find`-based last-row helpers
  raising 91 on empty sheets (#33), unguarded `GetFirstVisibleSheet` (#34),
  `CountNonParentOrIncludedRows` dead param / not perf-wrapped (#35),
  `WithPerformance` string dispatch (#36), and a code-hygiene sweep (#37).
- Removed commented-out code across `/src` (a stale `CreateOrClearSheet` copy and
  an old `GetLastRow` overload in `modHelpers_Workbook`, leftover
  `'MsgBox "Ribbon callback fired"` debug lines, an empty comment banner in
  `modMain_CountEquipmentRows`). Deleted the superseded `ExportAllVBAModules` v1
  (hardcoded `C:\VBA_Export\` path, skipped document modules) and renamed
  `ExportAllVBAModules2` to `ExportAllVBAModules`; docs updated to match.

### Changed

- Removed two dead modules
  ([#50](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/50),
  [#48](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/48)):
  `modHelpers_Diagnostics` (its only function, `SheetExists2`, had no callers
  and was just `SheetExists(ThisWorkbook, name)` with the workbook bound) and
  the empty `modMain_CustSchedule` placeholder. Recreate `modMain_CustSchedule`
  when Schedule automation actually starts.
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

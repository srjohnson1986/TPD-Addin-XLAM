# TPD Add-in — Architecture

A module-by-module map of the add-in as it is on `main`, grouped by the same `@Folder("TPD_Addin.X")` tags used in
Rubberduck's Code Explorer and in the `/src` folder layout. For the developer workflow (change, export, build,
release) see [CONTRIBUTING.md](../CONTRIBUTING.md).

## How this file is kept

- One `###` section per module. Bullets, one fact per line, each line short enough to read in a diff.
- Describe what the module does **today**. What it used to be called, what was removed and which PR did it
  belong in [CHANGELOG.md](../CHANGELOG.md) and `git log`, not here.
- Link an issue only where it explains *why* something non-obvious is the way it is. Links are reference-style:
  write `[#84]` in the text and add its URL to the list at the bottom of the file, in number order.
- A change to a module's job updates that module's section and nothing else.

## Overview

One custom ribbon tab ("TPD") with two groups, defined in `customUI/customUI14.xml`:

- **Sheet Tools** — the picker paths: "Create Customer EQ List", "Create Customer Schedule",
  "Split Sheet by Column", plus "Save Each Sheet to XLSX".
- **One-click** — "EQ List", "Schedule", "Split Sheets" (each runs straight from saved defaults, no picker)
  beside the large "Set Defaults" button.

How a click becomes work:

- Every button's `onAction` is a one-line callback in `modRibbonCallbacks`: `WithPerformance PERF_x`.
- `modPerformance.WithPerformance` runs that flow's `*_Internal` routine with screen updating, events and
  calculation managed.
- Set Defaults is the exception: `RunSetTPDDefaults` does real work, so it lives in `modSetDefaults`.
- The "EQ Count" flow has no ribbon button, but `modMain_CountEquipmentRows` and its `RunCountEquipmentRows`
  callback are kept ([#120]).

Other things worth knowing up front:

- Startup: `RibbonOnLoad` → `modStartup.InitializeAddIn`, which stamps the running version into the preference store.
- Saved settings live in the Windows registry per user (`modPreferences`), not in any workbook.
- The default header logo is the embedded `_Resources` / `DefaultLogo` shape from the base `.xlam`.
- A user can replace the logo in Set Defaults; their image is copied to `%APPDATA%\TPD_Addin\` and every flow
  uses it ([#95]).

## Naming: "header" vs. "heading"

"Header" has two unrelated meanings in this codebase, so the names split them:

- **`header`** — the TPD title block at the top of a sheet (`EQUIPMENT LIST` / Customer / Project / PM / Date / Rev),
  e.g. `InsertEQHeaderBlock`, `InsertDefaultCustScheduleHeader`.
- **`heading`** — the column heading row of the data table,
  e.g. `headingsRow`, `headingList`, `colHeading`, `GetHeadingList`, `FindHeadingIndex`.
- **`anchorRow`** — the logo's positioning row in `modHelpers_Logo`. It is neither: it can point at a title-block
  row or a heading row depending on the caller.

Use these names in new code. Don't reintroduce `header` / `headerRow` for column-heading concepts.

## TPD_Addin.Ribbon

### `modRibbonCallbacks`

- Holds every `customUI14.xml` `onAction` callback. Each is a single `WithPerformance PERF_*` line.
- Sheet Tools: `RunCreateCustEQList`, `RunCreateCustSchedule`, `SplitSheetByColumn`, `RunExportSheetsToXLSX`.
- One-click: `RunQuickEqList`, `RunQuickSchedule`, `RunQuickSplitSheets`.
- `RunCountEquipmentRows` is here too, although no button calls it ([#120]).
- The Split button's callback is `SplitSheetByColumn`, without the `Run` prefix its neighbours have. The name
  in `customUI/customUI14.xml` and the Sub's name must match; the static checks fail if they do not.
- Owns the global `gRibbon` reference and `RibbonOnLoad`, which triggers add-in initialization.

## TPD_Addin.EQList

### `modMain_CustEQList`

The Customer EQ List flow.

- `CreateCustEQList_Internal` (picker path) shows `frmCustEQListColumnPicker`.
- `CreateCustEQListFromDefaults_Internal` (one-click "EQ List") skips the picker and calls
  `modHelpers_Columns.ResolveOneClickColumns` with `PREF_EQLIST_COLUMNS` and the shipped default. It stops,
  after a message, when none of those columns are on the sheet.
- Both hand off to `CreateCustEQList_DoWork`.

What `CreateCustEQList_DoWork` does, in order:

1. Copies every row to a new `Customer EQ List` sheet **before** any formatting (`CopyEntireSheetRows`).
2. **Remove blank rows** toggle (`DeleteBlankRows`, ships on). Runs straight after the copy, judges emptiness
   across every source column, and comes before the status snapshot so row ordinals match the sheet ([#171]).
3. `StripCarriedRowFormatting` with `preserveStructuralRows:=True`, while all source columns are still present.
4. **Remove PARENT rows** toggle (`DeleteParentRows`) and the status snapshot (`PurchasedStatuses`), both while
   the Purchased column is still there.
5. `DeleteUnselectedColumnsByHeading`, then `ReorderColumnsByHeading`.
6. **Add EQ Count column** toggle (`NumberEquipmentRowsFromStatuses`, fed by the snapshot) while the heading row
   is still row 1.
7. **Plain PARENT rows** toggle (`ApplyPlainParentRows`: white fill, black text, not bold — overriding the
   grouping shade step 3 kept), then **Bold PARENT rows** toggle (`ApplyBoldParentRows`, ships off, [#174]).
   Both find their rows through `ParentDataRows` from the snapshot.
8. `FormatDataTable`; the **Add cell borders** toggle is its `addBorders` argument ([#122]).
9. **Add column filters** toggle (`AddHeadingRowFilter` on row 1), before `AutoFitUsedColumns` so the columns
   are widened to clear the dropdowns ([#147]).
10. Autofit columns and data-row heights. Steps 7–8 come before this so bold text is measured, not clipped.
11. Inserts the EQ header block and the logo (`InsertDefaultLogo … "right-top"`), then freezes panes. The
    title-block insert shifts the filter range down to the heading row.

Rules that hold across those steps:

- The four PARENT-row and EQ Count toggles need a `PURCHASED` heading. If it is absent they are skipped and
  `_DoWork` shows one `vbInformation` notice at the end (`SkippedPurchasedToggleNames`).
- Purchased reaches the output sheet only if it is in the chosen column list, whatever the toggles say.
- The seven toggles and their shipped values are defined once, in `modPreferences_Defaults`.

### `modMain_CountEquipmentRows`

The "EQ Count" flow, and the numbering the EQ List toggle reuses.

- Entry: `RunCountEquipmentRows` → `WithPerformance PERF_COUNT_EQ_ROWS` → `CountEquipmentRows_Internal`, which
  validates the active worksheet via `RequireActiveWorkbook`.
- `NumberEquipmentRows` inserts an `EQ COUNT` column and numbers every row except those marked `PARENT` or
  `INCLUDED` in the Purchased column. Numbers get leading zeros sized to the highest number.
- A row with a blank Purchased value **is** numbered. It is a real equipment line whose status is not filled in
  yet ([#6], closed as working as intended).
- `NumberEquipmentRows` is a thin sheet-reading wrapper over two routines:
  - `PurchasedStatuses(ws, headingRow, col)` snapshots the column to a 1-based `String()` array.
  - `NumberEquipmentRowsFromStatuses(ws, headingRow, statuses)` inserts the column and numbers from the array.
  - The split lets the EQ List toggle number a sheet after its Purchased column has been dropped ([#120]).
- The inserted `EQ COUNT` heading is `PasteSpecial xlPasteFormats`-matched to the neighbouring heading cell, so
  it is not left unformatted.
- `NumberEquipmentRowsFromStatuses` wraps two pure functions, which have Rubberduck unit tests ([#163]):
  - `ComputeEQCountValues(statuses)` — the `""` or next-number decision per row.
  - `EQCountNumberFormat(values)` — the leading-zero width.

## TPD_Addin.Schedule

### `modMain_CustSchedule`

The Customer Schedule flow — the schedule twin of `modMain_CustEQList` ([#112]).

- `CreateCustSchedule_Internal` (picker path, "Create Customer Schedule") shows `frmCustScheduleColumnPicker`.
- `CreateCustScheduleFromDefaults_Internal` (one-click "Schedule") calls the same
  `modHelpers_Columns.ResolveOneClickColumns` the EQ List one-click does, with `PREF_SCHEDULE_COLUMNS` and
  `DefaultScheduleColumns()`.
- Both call `CreateCustSchedule_DoWork`, which:
  1. Copies every row to a new `Customer Schedule` sheet **before** any formatting.
  2. Calls `StripCarriedRowFormatting` with its defaults for rows and `clearFontEmphasis:=False`.
  3. Drops unselected columns, then puts the rest in the chosen order (`ReorderColumnsByHeading`).
  4. `FormatDataTable wsNew, 1, "Tasks"`, then `AutoFitUsedColumns` and `AutoFitDataRowHeights`.
  5. `InsertDefaultCustScheduleHeader` — header rows, Date text and logo in one call.
  6. `SafeFreezePanes` at `SCHEDULE_HEADER_ROW_COUNT + 1`.
- The header builder and `SCHEDULE_HEADER_ROW_COUNT` live in `modHelpers_SheetSetup`.

## TPD_Addin.SplitExport

### `modMain_SplitSheets`

The Split Sheet by Column flow.

- `SplitSheetByColumn_Internal` (picker path) prompts via `frmSplitSheetOptions`.
- `SplitSheetByColumnFromDefaults_Internal` (one-click "Split Sheets") skips the picker. It resolves the group
  column from `PREF_SPLIT_GROUPCOL` and the columns from `PREF_SPLIT_COLUMNS`, each falling back to the shipped
  default, via `ResolveGroupColumn` / `ResolveColumnList`.
- Both call the module-private `RunSplit`, which runs `SplitSheetByColumn_DoWork` and reports the outcome. One
  shared tail, so the two paths cannot summarise a run differently.
- `SplitSheetByColumn_DoWork` creates or clears one sheet per unique value in the group column and copies the
  matching rows in.
- It shows a message and stops if the group column is not on the sheet. That is why this flow needs no
  `ResolveOneClickColumns`-style pre-check.
- `SplitOneGroup` builds one value's sheet and raises on failure; the caller records the failure against that
  value, so one bad value does not stop the run ([#32]).
- `UniqueRunName` adds a ` (2)` / ` (3)` suffix when two values sanitize to the same sheet name, trimming the
  base so the result stays within Excel's 31-character limit ([#57]). It is `Public` only so its unit tests
  can call it.

### `modMain_ExportSheets`

The "Save Each Sheet to XLSX" flow.

- Prompts for a filename suffix and date via `frmFilenameOptions`.
- Saves each eligible worksheet to its own `.xlsx` in a per-workbook export folder.
- `ShouldExportSheet` decides eligibility ([#8]):
  - `_`-prefixed internal sheets and very-hidden sheets are always skipped.
  - Plain-hidden sheets are skipped unless the source-level `EXPORT_INCLUDE_HIDDEN_SHEETS` flag is on.
- Per-sheet failures are collected and reported in one summary; they do not abort the run ([#15]).

### `modExport_Files`

File and folder work for the export flow.

- `EnsureExportFolder` creates or finds the export folder next to the workbook, named after the workbook file.
  It returns `""`, after a message, if the workbook is unsaved or the folder cannot be made.
- `ExportSheetToXLSX` copies one worksheet out to its own `.xlsx`. It returns `""` on success or a reason
  string on failure.
  - If the target name collides with an already-open workbook, the file gets a ` (export)` suffix.
  - An existing target file on disk is overwritten without prompting ([#64]).
  - Any `SaveAs` failure is caught, the temporary `ws.Copy` workbook is closed without saving, and the reason
    is returned ([#15]).
- `BuildExportSuffix(userAppend, includeDate)` builds the `<text> - <date>` suffix once.
- `PreviewExportFileName(sampleSheet, userAppend, includeDate)` renders the `<sheet> - <suffix>.xlsx` name a
  sheet would be saved as.
- The `frmFilenameOptions` live preview and the real export share both functions, so they cannot disagree.
- Both have unit tests in `/tests/unit`.

## TPD_Addin.Preferences

### `modPreferences`

Per-user key/value preference storage.

- Stored in the Windows registry via `SaveSetting` / `GetSetting`, under
  `HKCU\Software\VB and VBA Program Settings\TPD_Addin\Preferences`.
- Registry, not a sheet inside the `.xlam`: Excel never saves an add-in on exit, so in-workbook writes were
  silently lost on quit ([#84]).
- `SavePref` / `LoadPref` / `DeletePref` for single values. A missing key returns the caller's default.
- `SaveColumnList` / `LoadColumnList` for comma-separated column lists.
- `ResolveColumnList(keys, fallbackCsv)` and `ResolveGroupColumn(keys, fallback)` walk a caller-supplied key
  array in priority order and return the first non-empty saved value, else the shipped default ([#96]).
  - Pickers pass `Array(LastUsed*, DefaultUser*)`.
  - One-click flows pass `Array(DefaultUser*)`.
- Each of those two is a thin registry-reading wrapper around a pure sibling — `FirstNonEmptyColumnList` /
  `FirstNonEmptyValue` — that picks among already-read values. The pure ones have unit tests ([#163]).
- `SavePrefs` writes an array of key/value pairs. It returns `False` if any write failed; there is no rollback.
- `AllPrefsUnset` is `True` only if every listed key reads back empty.
- `LoadToggle(key, defaultOn)` / `SaveToggle(key, isOn)` store a boolean as `"1"` / `"0"`.

Support for the Set Defaults dialog:

- `NormalizeColumnList` — tabs / CR / LF become `, `, empties are dropped, duplicates are kept.
- `CollapseWhitespace` — a heading name's compare form: tabs to spaces, then
  `modHelpers_Strings.NormalizeCellText`.
- `SaveAllDefaults` and `DefaultsNeverSaved` — thin callers of `SavePrefs` / `AllPrefsUnset` over the one shared
  list of the four dialog keys, `DefaultsPrefKeys` ([#102]).

### `modPreferences_KeyMap`

Every preference key, as a `PREF_*` constant, so keys never float as loose string literals.

Column defaults are three layers ([#96]):

| Layer | Keys | Written by | Read by |
|---|---|---|---|
| Shipped | none — `modPreferences_Defaults` | — | everything, as the last fallback |
| User one-click (`DefaultUser*`) | `PREF_EQLIST_COLUMNS`, `PREF_SCHEDULE_COLUMNS`, `PREF_SPLIT_COLUMNS`, `PREF_SPLIT_GROUPCOL` | **only** the Set Defaults dialog | one-click buttons; pickers as their second choice |
| User picker (`LastUsed*`) | `PREF_EQLIST_PICKER_COLUMNS`, `PREF_SCHEDULE_PICKER_COLUMNS`, `PREF_SPLIT_PICKER_COLUMNS`, `PREF_SPLIT_PICKER_GROUPCOL` | the picker forms, on OK | pickers, as their first choice |

- Pickers resolve `LastUsed*` → `DefaultUser*` → shipped.
- One-click resolves `DefaultUser*` → shipped.

Other keys:

- `PREF_LOGO_PATH` — the user's chosen header logo ([#95]).
- `PREF_SETDEFAULTS_LAST_TAB` — the Set Defaults dialog's last-shown tab ([#135]).
- Seven EQ List behaviour toggles, stored `"1"` / `"0"`: `PREF_EQLIST_REMOVE_BLANK_ROWS`,
  `PREF_EQLIST_REMOVE_PARENT_ROWS`, `PREF_EQLIST_ADD_COUNT_COLUMN`, `PREF_EQLIST_PLAIN_PARENT_ROWS`,
  `PREF_EQLIST_BOLD_PARENT_ROWS`, `PREF_EQLIST_ADD_CELL_BORDERS`, `PREF_EQLIST_ADD_COLUMN_FILTERS`.
  - They are **not** in `DefaultsPrefKeys`, so they cannot trip the dialog's first-run notice.
- `PREF_VERSION` — the add-in version last run by this user.

### `modPreferences_Defaults`

What the add-in ships as, before a user saves anything. The single source for both pickers and one-click ([#98]).

- The shipped column lists, comma-separated: `DefaultEqListColumns`, `DefaultScheduleColumns`,
  `DefaultSplitColumns`, `DefaultSplitGroupColumn`.
- They seed the dialog's first run, back its per-page "Restore defaults", are the value `LoadPref` returns for
  an unsaved key, and are the final `ResolveColumnList` / `ResolveGroupColumn` fallback.

The EQ List toggle table (`EqListToggleTable`) holds one row per toggle:

| Toggle | Ships | Needs a `PURCHASED` column |
|---|---|---|
| Remove PARENT rows | off | yes |
| Add EQ Count column | off | yes |
| Plain PARENT rows | off | yes |
| Add cell borders | **on** | no |
| Add column filters | off | no |
| Remove blank rows | **on** | no |
| Bold PARENT rows | off | yes |

- Each row is a `key|shipsOn|display name|needsPurchased` string. Keep it a flat array of delimited strings:
  an earlier array-of-arrays version raised an unhandled error in the lookup and the cause was never found.
- Readers: `EqListToggleKeys` (every key, in checkbox order), `EqListToggleShipsOn`, `EqListToggleName`,
  `EqListToggleNeedsPurchased`, and `LoadEqListToggle` (saved value, else the shipped one).
- Adding a toggle = a `PREF_*` constant, a row here, a checkbox, and one line each in the form's
  `LoadEqListToggles` / `SaveEqListToggles`.
- The `TOGGLE_*` field-position constants sit at the top of the module. Module-level constants below the
  first procedure do not resolve.

### `modPreferences_Initializer`

- `InitializePreferences` stamps `PREF_VERSION` with `ADDIN_VERSION` when it changes (for support and
  diagnostics).
- It seeds no values: `LoadPref` returns the caller's default for any unset key.

### `modSetDefaults`

The "Set Defaults" ribbon entry point.

- `RunSetTPDDefaults` shows `frmSetTPDDefaults` modally.
- For the dialog's "Save & Run" button it then runs `WithPerformance gPendingDefaultsFlow` **after** `Unload`,
  so the perf wrapper's screen and calculation toggling happens with no modal form in memory ([#100]).
- `gPendingDefaultsFlow` is a public `String` the form sets to a `PERF_*_DEFAULTS` constant. It is cleared
  before the form shows and before the flow runs.
- `ClearTPDDefaultsStatusBar` is the `Application.OnTime` target that wipes the save confirmation. It must be
  a sub in a standard module, because `OnTime` cannot call into a form.
- **Its name is load-bearing.** The form schedules it by string literal — the only such reference in the
  codebase — so renaming it would silently stop the status bar clearing. `tools/Test-StaticChecks.ps1` checks
  that the name still resolves.

## TPD_Addin.Forms

Every UserForm lives here — `src/Forms` on disk, `TPD_Addin.Forms` in Code Explorer — and not beside the flow
that shows it, so all five are in one place. Which flow shows a form is in that flow's section above.

What is shared, and where it lives (all in Helpers):

- `modHelpers_DialogChrome` — brand band and About links, on all five forms.
- `modHelpers_ColumnPicker` — grid build, pre-fill, OK-guard and save, for the three pickers.
- `modHelpers_CheckboxLayout` / `modHelpers_CheckboxSelection` — the checkbox grid itself.

What stays on a form: its controls, its event handlers (a standard module cannot wire an event), and whatever
is genuinely its own.

Chrome common to all five forms:

- Controls `lblBrandBar`, `imgBrandLogo`, `lblTitle`, `lblHelp`, and the `lblVersion` / `lblUserGuide` links.
- `UserForm_Initialize` calls `ApplyDialogChrome` and `InitAboutLinks`.
- The two link `_Click` handlers call `modAbout`.

Common to the three column pickers:

- `LoadColumns` is one call to `InitColumnPicker` with that picker's keys and shipped list.
- `cmdOK_Click` hands the ticked columns to `CommitColumnPicker`, which refuses an empty selection ([#104])
  and writes the picker's own `LastUsed*` key — never the Set Defaults key.
- A `cmdSelectAll` / `cmdSelectNone` / `cmdRestoreColumns` row above `fraColumns` calls `SetAllColumns` /
  `RestoreColumnDefaults` ([#149]).
- Other controls: `fraColumns`, `cmdOK`, `cmdCancel`.

### `frmCustEQListColumnPicker`

Chooses which columns to keep for a Customer EQ List.

- Title: "Customer EQ List".
- Keys: `PREF_EQLIST_PICKER_COLUMNS` (its own), `PREF_EQLIST_COLUMNS` (the Set Defaults one).
- Shipped list: `DefaultEqListColumns()`.

### `frmCustScheduleColumnPicker`

Chooses which columns to keep for a Customer Schedule. A visual twin of the EQ List picker.

- Title: "Customer Schedule", set as a design-time caption.
- Keys: `PREF_SCHEDULE_PICKER_COLUMNS`, `PREF_SCHEDULE_COLUMNS`.
- Shipped list: `DefaultScheduleColumns()`.

### `frmSplitSheetOptions`

Chooses the group-by column and which columns to keep when splitting a sheet.

- Title: "Split Sheet by Column". The dropdown `cboGroupColumn` is labelled "Group rows by" (`lblGroupColumn`).
- Column grid keys: `PREF_SPLIT_PICKER_COLUMNS`, `PREF_SPLIT_COLUMNS`; shipped list `DefaultSplitColumns()`.
- `LoadColumns` calls `InitColumnPicker`, then fills the group dropdown itself.
- `SelectGroupColumn` selects
  `ResolveGroupColumn(Array(PREF_SPLIT_PICKER_GROUPCOL, PREF_SPLIT_GROUPCOL), DefaultSplitGroupColumn())`.
  - It matches on the `NormalizeCellText` compare form of both sides, so a whitespace-dirty saved value still
    resolves ([#103]).
  - It falls back to the first column if the value is not found.
- `cmdOK_Click` validates the group column itself, then calls `CommitColumnPicker` ([#10]) and saves
  `PREF_SPLIT_PICKER_GROUPCOL` alongside.
- The Select all / none / Restore defaults row acts on `fraColumns` only, never on `cboGroupColumn`.

### `frmFilenameOptions`

Appends free text and/or today's date to exported filenames.

- Outputs: `UserAppend`, `IncludeDate`, `Cancelled`.
- Title: "Filename Options". It uses the compact form of the shared chrome.
- Live **Preview**: `lblFullFilename` inside `fraPreview`, recomputed by `RefreshPreview` on every
  `txtAppend_Change` / `chkIncludeDate_Click` via `modExport_Files.PreviewExportFileName`.
- The caller seeds `frm.SampleSheetName = wb.ActiveSheet.Name` before `.Show`.
- The first preview is drawn in `UserForm_Activate`, not `UserForm_Initialize`. Assigning `SampleSheetName` is
  the caller's first member access, so `Initialize` runs before that value lands ([#154]).

### `frmSetTPDDefaults`

The "Set Defaults" dialog ([#25]).

Layout:

- A MultiPage (`mpgPages`) with EQ List / Schedule / Split Sheets / Logo tabs.
- Footer: **Save** (`cmdOK`) / **Save & Run** (`cmdSaveRun`) / **Cancel** ([#100]).
- `lblVersion` and `lblUserGuide` sit in the bottom corner, outside `mpgPages`, so they show on every tab.
- Each tab has a muted hint label under its field: `lblEqListHelp`, `lblScheduleHelp`, `lblSplitColsHelp`,
  `lblSplitGroupHelp`, `lblLogoHelp`.
- `SetHelperText` sets those captions and their grey colour in code, so the text stays visible in review
  instead of hiding in the `.frx` ([#143]).

Column tabs:

- Each is a MultiLine TextBox of plain comma-separated heading names.
- Text is normalized on paste (Ctrl+V caught on `KeyUp`) and on `Exit`.
- A per-tab "Restore defaults" button resets from `modPreferences_Defaults`.

EQ List behaviour toggles ([#97]):

- Seven checkboxes: `cbxRemoveBlankRows`, `cbxRemoveParentRows`, `cbxAddItemCountColumn`, `cbxPlainParentRows`,
  `cbxBoldParentRows`, `cbxAddCellBorders`, `cbxAddColumnFilters`.
- They sit in the frame `fraEqListOptions`. `cbxRemoveBlankRows` and `cbxBoldParentRows` are in its right-hand
  column.
- `LoadEqListToggles` fills them. `SaveEqListToggles` persists them via `modPreferences.SaveToggle`; it is
  called from `CommitDefaults` after the column keys.
- `cmdRestoreEqList_Click` resets each to its shipped value (`EqListToggleShipsOn`).

Saving:

- `cmdOK_Click` and `cmdSaveRun_Click` share `CommitDefaults()`: normalize → empty guard → staged logo →
  `SaveAllDefaults`. It returns `False` having already shown the reason.
- Save is blocked, with a message, if any of the three column-list fields is empty ([#104]).
  - An empty list is never meaningful: every consumer reads it as "use the built-in defaults", so
    "Restore defaults" is the way to get the built-in list.
  - The Split group-column field **is** allowed to be empty.
- On success there is a status-bar confirmation and no MsgBox. `ClearTPDDefaultsStatusBar` clears it later.
- A write failure keeps the form open and shows a `vbCritical` box.
- Cancel, Esc and the title-bar X all discard silently.
- Save & Run sets `modSetDefaults.gPendingDefaultsFlow` from `PerfConstForActivePage()` before `Me.Hide`, so the
  open tab's one-click flow runs after the form unloads.
- `SyncSaveRunVisibility` (called on load and from `mpgPages_Change`) hides `cmdSaveRun` on the Logo tab.

First-run notice ([#99]):

- `lblFirstRunNotice` ("these are the built-in defaults") is a **form-level** control above `mpgPages`, not a
  child of `pgEqList`.
- `ApplyFirstRunNoticeLayout` shows it only while `DefaultsNeverSaved()`. When shown, it moves `mpgPages` down
  and shrinks it by the banner's height so the tabs do not jump.

Remembered tab ([#135]):

- `mpgPages_Change` writes `PREF_SETDEFAULTS_LAST_TAB`.
- `SavedTabIndex` reads it back in `UserForm_Initialize` (clamped, default 0).
- `mInitializing` stops the load from bouncing through the change handler.

Logo tab ([#95], [#109]):

- **Choose image…** (`Application.FileDialog`; PNG / JPG / GIF / BMP) and **Use built-in logo** only *stage*
  a change.
- `imgLogoPreview` (zoomed) previews the staged image via `modGdiPlus.LoadPictureGDIP`, and `lblLogoStatus`
  says "On OK: …".
- Save applies it through `ApplyPendingLogo` → `modHelpers_Logo.SetUserLogo` / `ClearUserLogo`, *before* the
  column save. Cancel discards it.
- When the built-in logo is active, `imgLogoPreview` shows its own design-time picture, set in the VBE from
  `assets/tpdHeaderLogo.jpg` and cached in `mBuiltInLogoPic`. There is no reliable Shape-to-`Picture` path for
  the `_Resources` shape.

## TPD_Addin.Helpers

Shared utilities used by more than one feature area.

### `modHelpers_Workbook`

Sheet and workbook utilities.

- `RequireActiveWorkbook` — the standard "stop with a message if no workbook is open" guard.
- `GetFirstVisibleSheet` — guards no-workbook and no-visible-sheet; returns `Nothing` after a message.
- `CreateNewSheetAfterLast`, `CreateOrClearSheet` — add a named sheet, or reuse and empty an existing one.
- `GetNextAvailableSheetName` — the next free sheet name for a base name.
- `SheetExists(wb, name)` — use `SheetExists(ThisWorkbook, name)` for "does this sheet exist in the add-in".
- `IsWorkbookOpen(fileName)`.
- `GetLastRow` — the last row with content in *any* column. `Find`-based and safe on an empty sheet.
  - This is the "end of the data table" helper. Use a per-column `.End(xlUp)` only when you mean that one
    column ([#89]).
- `GetLastCol` — the last-column equivalent.

### `modHelpers_Headers`

- `GetHeadingList` reads a worksheet's heading row into an array.
- `FindHeadingIndex` finds a heading's column index.
  - It scans right-to-left by default, so the authoritative "Vendor" column wins over decoy columns ([#45]).
  - Pass `preferRightmost:=False` for a plain left-to-right lookup.

### `modHelpers_Columns`

Column selection, ordering and row copying.

- `CollectionContainsText` — collection-membership test.
- `DeleteUnselectedColumnsByHeading` — deletes columns not in a selected-headings list.
- `ReorderColumnsByHeading` — cut/inserts the kept columns into a chosen heading order. EQ List and Schedule
  call it after the delete, so the order set in Set Defaults is honoured ([#127]).
- `HeadingsPresentOnSheet` — the subset of a wanted list that actually appears on a sheet's heading row, so a
  default list matching nothing cannot delete every column ([#96]).
- `ResolveOneClickColumns(ws, defaultKey, shippedCsv, flowName, pickerCommand)` — for the one-click buttons:
  saved Set Defaults value → shipped list → narrowed to the sheet. Returns `Nothing` after one shared
  "none of your default *flowName* columns were found" message.
- `GetUniqueValuesInColumn` — the unique values in a column.
- `CopyFilteredRowsByColumns` — copies matching rows across sheets, projected onto a column selection in the
  selection's order ([#127]). **Values only**; formatting is the caller's job.
- `CopyEntireSheetRows` — copies every row verbatim. A whole-row copy carries row grouping across.

Four of these are thin `Worksheet`-facing wrappers around a pure sibling that takes and returns plain values
only. The pure ones have Rubberduck unit tests in `/tests/unit`, which `BuildAddin` never imports ([#163]):

| Sheet-facing | Pure |
|---|---|
| `HeadingsPresentOnSheet` | `HeadingsPresentIn` |
| `GetUniqueValuesInColumn` | `UniqueValues` |
| `CopyFilteredRowsByColumns` (its column map) | `BuildOrderedColumnMap` |
| `ReorderColumnsByHeading` (its move sequence) | `PlanColumnMoves` |

### `modHelpers_Strings`

- `SanitizeSheetName`, `SanitizeFileText` — make text safe for a sheet name or a file name.
- `NormalizeCellText` — first physical line, NBSP to space, collapsed spaces, trimmed. Used wherever cell
  values are compared or de-duplicated.
- All three have unit tests in `/tests/unit`.

### `modHelpers_SheetSetup`

Header-block construction.

- `InsertEQHeaderBlock` — the EQ title block. The EQ List flow adds the logo and freezes panes separately.
- `InsertDefaultCustScheduleHeader` — the schedule title block, Date text and logo in one call.
- Owns `EQ_HEADER_ROW_COUNT` / `SCHEDULE_HEADER_ROW_COUNT`, `InsertRows` and `GetTodaysDate`.

### `modHelpers_SheetFormatting`

Formats the generated EQ List, Schedule and Split sheets: borders, alignment, fonts, frozen panes, column
widths and row heights.

`FormatDataTable(ws, headingsRow, leftAlignHeading, addBorders)`:

- Centres every column except `leftAlignHeading`, which keeps its natural left alignment.
- `leftAlignHeading` defaults to `"Description"`; the Schedule flow passes `"Tasks"`.
- `addBorders` (default `True`) gates the thin data-range border, so the EQ List "Add cell borders" toggle
  can drop it ([#122]).

`StripCarriedRowFormatting(ws, headingsRow, preserveStructuralRows, clearFontEmphasis)`:

- Removes the formatting a row-for-row copy carries over from the source sheet: the fill on every row and, on
  data rows, the font emphasis too — bold, italic, underline, strikethrough, colour back to automatic ([#169]).
- The EQ List and Schedule `_DoWork` call it right after `CopyEntireSheetRows`, **before** any columns are
  dropped, so the `PURCHASED` / `INTERNAL ID` / `CUSTOMER ID` lookups still see the full source.
- `preserveStructuralRows:=True` (EQ List) leaves the rows the pure `IsStructuralRow` picks untouched, fill and
  font both:
  - With a `PURCHASED` column, that is `PARENT` rows only.
  - With no `PURCHASED` column, it falls back to rows with neither ID ([#170]).
- The default (Schedule) clears the fill on every row.
- The Schedule also passes `clearFontEmphasis:=False`: once every fill is gone, bold is what marks its
  summary rows.
- The rule is deliberately independent of row order, so the picker and one-click outputs agree ([#130]).

`AutoFitDataRowHeights(ws, headingsRow)`:

- Autofits row height from `headingsRow` through the last data row, so wrapped or multi-line text is not
  clipped ([#162]).
- Called right after `AutoFitUsedColumns`, because row height depends on the final column width.
- For EQ List and Schedule it runs **before** the header/logo block is inserted. That insert only adds blank
  rows above and shifts the rest down, so the heights move with their rows.
- Split Sheet has no header block, so `FormatSplitSheet` calls it directly.
- In Split output it currently does nothing for wrapped text: `CopyFilteredRowsByColumns` copies **values**
  only, so a source column's `WrapText` never arrives. EQ List and Schedule copy whole rows, which does carry
  `WrapText`.

Also here: `FormatSplitSheet`, `SafeFreezePanes`, `AutoFitUsedColumns`.

### `modGdiPlus`

- `LoadPictureGDIP(fileName)` loads any raster image (PNG / JPG / GIF / BMP / TIFF) into an
  `stdole.IPictureDisp` for an MSForms `Image` control. VBA's built-in `LoadPicture` cannot read PNG.
- It is the codebase's **only `Declare`**: isolated here and `#If VBA7` guarded.
- Every failure path returns `Nothing`; it never raises.
- It backs the Set Defaults logo preview ([#109]).

### `modHelpers_Logo`

Places and sizes the header logo, and manages the user's own logo.

- `InsertDefaultLogo` (alignment keywords, optional scale-to-max-rows):
  - inserts the user's image via `Shapes.AddPicture` when `EffectiveLogoPath()` is set;
  - otherwise copies the embedded `DefaultLogoShape()` — `_Resources` / `DefaultLogo`, resolved in that one
    place ([#105]).
- `ResizeImageToMaxRows` is the scaler.
- User-logo management ([#95]):
  - `SetUserLogo` — validates, copies into `UserLogoDir()` (`%APPDATA%\TPD_Addin\`) and saves `PREF_LOGO_PATH`.
  - `ClearUserLogo`.
  - `EffectiveLogoPath` — the copy if it still exists, else `""`.
  - `LogoPreviewPicture` — for the dialog.

### `modHelpers_CheckboxLayout`

Builds the checkbox grid the three column pickers share (`LayoutCheckboxes`).

- One box per heading, in **at most 3** columns, filled top-to-bottom then left-to-right (column-major).
- More headings than fit in 3 columns at `rowsPerColumn` height make the columns taller, never add a 4th
  column ([#47]).
- `NormalizeFrameScroll` gives `fraColumns` an **always-on** vertical scrollbar (`fmScrollBarsVertical` plus
  `KeepScrollBarsVisible`; `ScrollHeight` is the larger of content and interior height). A short list shows
  an idle bar and a long one scrolls; the form never grows ([#150]).
- Column width subtracts `SCROLLBAR_ALLOWANCE`, so the third column clears the bar.
- It never resizes the frame or the form.

### `modHelpers_CheckboxSelection`

Reads and sets which checkboxes are ticked.

- `GetSelectedColumns` — the ticked boxes, as a `Collection`.
- `HasColumnSelection` — whether any box is ticked.
- `ApplyColumnSelection(fra, resolved)` — ticks exactly the boxes in `resolved` and clears the rest. The
  caller has already run the list through `modPreferences.ResolveColumnList` ([#81], [#96]).
- `SetAllColumns(fra, isChecked)` — Select all / Select none.
- `RestoreColumnDefaults(fra, defaultUserKey, shippedCsv)` — resolves a **single** key (`DefaultUser*` →
  shipped), skipping the picker's `LastUsed*`, so it means the same as "Restore defaults" in Set Defaults.
- Neither of the last two saves anything. `cmdOK_Click` still writes `LastUsed*`.

### `modHelpers_ColumnPicker`

The behaviour the three picker forms share, written once.

- `InitColumnPicker(fra, headingList, rowsPerColumn, baseName, pickerKey, defaultKey, shippedCsv)`:
  - builds the grid (`LayoutCheckboxes`);
  - ticks the resolved selection: `ResolveColumnList(Array(pickerKey, defaultKey), shippedCsv)` →
    `ApplyColumnSelection`.
- `CommitColumnPicker(fra, pickerKey)` is the OK-button guard and save:
  - returns `False`, having shown the reason, when nothing is ticked;
  - otherwise writes the ticked set to the picker's own `LastUsed*` key and returns `True`.
  - The one-click `DefaultUser*` key is never touched here.

### `modHelpers_Reporting`

- `ReportBatchOutcome(summary, failures, failureNoun, errTitle, okTitle)` — the end-of-run summary box.
- It shows "N done" plus a list of `'name': reason` failures: a warning box when there are failures, an info
  box when there are none.
- Used by Split Sheet by Column and Save Each Sheet to XLSX.

### `modHelpers_DialogChrome`

Run-time chrome for all five dialogs, so they cannot drift apart ([#118]).

- `ApplyDialogChrome(brandBar, helpLabel)` — TPD-grey `lblBrandBar` fill with a blank caption, and the
  `#4A4A4A` help colour.
- `InitAboutLinks(lblVersion, lblUserGuide)` — the captions (`v<ADDIN_VERSION>` / "User guide"), the
  `ControlTipText` (`modAbout.AboutReleaseUrl()` / `ABOUT_USER_GUIDE_URL`) and the blue-underline link look.
- Both are `Nothing`-safe.
- Styling is done in code, not in the `.frx`, so it shows up in review.
- The two `_Click` handlers stay on each form, because a standard module cannot wire an event.

## TPD_Addin.Core

Add-in lifecycle and infrastructure. Not directly user-facing.

### `modPerformance`

- `WithPerformance` turns off `ScreenUpdating` and `EnableEvents` and sets manual calculation around one
  internal routine, then restores all three.
- On an error inside the routine, `CleanUp` still restores all three, then shows `Err.Description` in a
  `MsgBox` ([#7]).
- The routine is chosen by a `PERF_*` string constant, so a mistyped name fails to compile and never reaches
  the `Unknown action` branch at run time ([#36]).
- The EQ List, Schedule and Split flows each have a picker `PERF_*` and a one-click `PERF_*_DEFAULTS` constant,
  each with a `Case`.
- Adding a flow = a `PERF_*` constant, a `Case`, and a callback that passes it. `tools/Test-StaticChecks.ps1`
  checks all three line up.

### `modStartup`

- Holds the `ADDIN_VERSION` constant.
- `InitializeAddIn`, called from `RibbonOnLoad`, runs `InitializePreferences` (the registry version stamp).
- `_Resources` and the embedded logo come from the base `.xlam`, not from code.

### `modExport_VBAModules`

- `ExportAllVBAModules` exports every component, document modules included, to `/src`. It is the macro behind
  the source-control workflow.
- Each file goes into the subfolder named by its module's `@Folder` tag.
- It finds `/src` itself: it walks up from the workbook's folder to the repo root (a folder holding both `src`
  and `customUI`), and falls back to a folder picker.
- One failing component is logged and the export carries on. `export_log.txt` records what happened.

### `modAbout`

The outbound links on the dialogs ([#94]).

- Holds the two URLs in one place:
  - `AboutReleaseUrl` — this version's GitHub release page, `…/releases/tag/v<ADDIN_VERSION>`.
  - `ABOUT_USER_GUIDE_URL` — the user guide on the wiki.
- `OpenReleasePage` / `OpenUserGuide` open them through a private `ThisWorkbook.FollowHyperlink` wrapper that
  swallows errors.
- This is not a network call of the add-in's own: `FollowHyperlink` hands off to the default browser. The
  add-in never checks GitHub for updates.

## TPD_Addin.Document

Document modules — code-behind tied to the workbook and sheet objects. They cannot be removed and re-imported
the way standard modules can (see CONTRIBUTING.md).

### `ThisWorkbook`, `Sheet1`, `Sheet2`, `Sheet3`

- Workbook-level and sheet-level code-behind.
- All four are currently code-free: each holds only its `@Folder` tag.

## Regressions to guard against

All fixed. Don't undo them.

- **`modPerformance.WithPerformance`** ([#7], [#36]) — the `CleanUp` handler must **always** restore
  `ScreenUpdating` / `EnableEvents` / `Calculation`, or they stay broken for the Excel session after any
  error. Callers pass a `PERF_*` constant, not a bare string.
- **Never call a `Sub` as `MySub (obj)`** ([#29]) — the parens force by-value evaluation (the default
  property), so a `Worksheet` / `Workbook` argument raises run-time 438. Use `MySub obj` or `Call MySub(obj)`.
- **Split Sheet by Column** ([#32]) — the loop reports per-value failures in its summary, so a regression
  there is at least visible.

Open work is tracked in [GitHub Issues](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues). The history of
fixed defects is in [CHANGELOG.md](../CHANGELOG.md).

[#6]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/6
[#7]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7
[#8]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/8
[#10]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/10
[#15]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/15
[#25]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25
[#29]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/29
[#32]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/32
[#36]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/36
[#45]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/45
[#47]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/47
[#57]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/57
[#64]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/64
[#81]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/81
[#84]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/84
[#89]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/89
[#94]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/94
[#95]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/95
[#96]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/96
[#97]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/97
[#98]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/98
[#99]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/99
[#100]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/100
[#102]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/102
[#103]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/103
[#104]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/104
[#105]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/105
[#109]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/109
[#112]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/112
[#118]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/118
[#120]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/120
[#122]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/122
[#127]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/127
[#130]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/130
[#135]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/135
[#143]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/143
[#147]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/147
[#149]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/149
[#150]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/150
[#154]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/154
[#162]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/162
[#163]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/163
[#169]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/169
[#170]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/170
[#171]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/171
[#174]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/174

# Tests

Two kinds of test live here.

| Folder | What | How it runs |
| ------ | ---- | ----------- |
| `unit/` | Unit tests for the pure VBA logic (column lists, preferences, EQ Count numbering, which EQ List rows count as structural, sheet and file naming) | `tools\Test-Unit.ps1` headlessly, or [Rubberduck](https://rubberduckvba.com/)'s Test Explorer — see [Unit tests](#unit-tests) |
| `fixtures/` | **Synthetic** sample workbooks to run the ribbon flows against by hand | Open a fixture, click through the TPD ribbon, compare with the expectations below |

Everything in `fixtures/` is **made up**: company names, tags, part numbers, people and dates. There is no customer data in the repo and there must never be — see [CONTRIBUTING.md](../CONTRIBUTING.md#license-and-conduct).

## Unit tests

Two ways to run them. They give the same verdicts, because the tests do their own comparing (see "The assert object" below).

**Headless — the quick check.** From the repo root:

```
powershell -ExecutionPolicy Bypass -File tools\Test-Unit.ps1
```

It opens `build\TPD_Addin.xlam` read-only in a hidden Excel, imports this folder into it in memory (nothing is saved), runs every `'@TestMethod`, and exits non-zero if any fail. It refuses to run against a build older than `/src`; add `-Build` to rebuild first. It needs Excel and "Trust access to the VBA project object model", so it is a local check, not a CI job. `-PlanOnly` lists what would run without starting Excel.

- **It checks itself every run.** Three canary tests are added: one must pass, one must fail on an assert, one must fail by raising an error. If the runner reports anything else for them, the whole run is thrown out.
- **It cannot leave Excel stuck.** If Excel stops on a dialog nobody can see (a compile error, a `MsgBox` in the code under test, the VBE in break mode), the script waits `-TimeoutSeconds` (default 120), prints which test it was on and the dialog's text, ends the Excel process *it* started and no other, and exits 3. It does not retry.
- Leave the VBE's error trapping on **Break on Unhandled Errors** (the default). **Break on All Errors** stops on the errors the runner is meant to catch.

**Rubberduck's Test Explorer — the full check, before a release.** Import `/src` (document modules aside) and everything in `tests/unit` into a scratch workbook, open it, then **Rubberduck → Unit Tests → Test Explorer → Run All Tests**. The workbook must be open for Run All to find the tests.

### The assert object

Test modules get their assert from `modTestSupport.NewTestAssert()`, which returns a `clsTestAssert`:

```vb
Private Assert As clsTestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = NewTestAssert()
End Sub
```

Don't use `CreateObject("Rubberduck.PermissiveAssertClass")` directly. Rubberduck's assert only does anything while Test Explorer is driving; called any other way, `AreEqual` and `Fail` are silent no-ops and every test "passes". `clsTestAssert` makes the comparison itself, records a failure for the headless runner, and also passes it to Rubberduck's `Fail` so Test Explorer shows it.

It offers `AreEqual`, `AreNotEqual`, `IsTrue`, `IsFalse` and `Fail`. Equality is by value with VBA's usual numeric coercion, text is compared exactly (case-sensitive), and arrays are never equal — assert on their elements.

### Files

| File | What |
| ---- | ---- |
| `modTest_*.bas` | The tests. One module per area; each `'@TestMethod` sub is one test |
| `clsTestAssert.cls` | The assert object |
| `modTestSupport.bas` | `NewTestAssert`, and where failed asserts are recorded |

None of these are part of `/src`: `BuildAddin` never imports them and they never ship in the `.xlam`. A function has to be `Public` in a standard module for a test to call it.

## Fixtures

Run from the repo root to (re)generate them. They are built from scratch with a fixed seed, so the content is the same every run:

```
pip install openpyxl
python tests/fixtures/generate_fixtures.py
```

| File | Models | Used for |
| ---- | ------ | -------- |
| `fixtures/EQ_List_Sample.xlsx` | A Smartsheet-style equipment-list export: 25 columns, heading in row 1, a nested tree of PARENT and item rows (Excel row grouping, parents above children), row fills by type | Create Customer EQ List / EQ List, Split Sheet by Column / Split Sheets, Save Each Sheet to XLSX |
| `fixtures/Schedule_Sample.xlsx` | A Smartsheet-style project schedule: 16 columns (6 hidden), a task tree up to 6 levels deep, real dates, `%` complete, status | Create Customer Schedule / Schedule |

Edit the **generator**, not the `.xlsx`, then regenerate and commit both. (The `.xlsx` container isn't byte-identical between runs, so expect a binary diff even when nothing changed — only regenerate when you mean to.)

## EQ List fixture — what's in it

101 data rows under the heading row: **16 PARENT rows**, **70 item rows**, **1 blank row** inside the data (row 85) and **14 blank rows** at the end. Outline levels 0–4. The shipped default EQ List columns (`INTERNAL ID, Location, TYPE1, TYPE2, SIZE, RATING, CONNECTION TYPE, Description, Manufacturer, Model, Details`) are all present, alongside extras the default list should drop (`Status`, `OWNER`, `Vendor`, `Redbox Row 1–3`, `Filename`, `PO Number` …). `Purchased` has `PARENT` / `NO` / `N/R` / `INCLUDED`, so the three PARENT-row toggles on the Set Defaults EQ List page have something to act on.

Deliberately awkward inputs, and where to find them:

| Case | Where | What to check |
| ---- | ----- | ------------- |
| Wrapped, multi-line text | `Details`, rows 22, 43, 62 | After EQ List / Schedule / Split, those rows are tall enough to show all the text, header/logo rows keep their height ([#162](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/162)) |
| Embedded line break in a normal cell | `TYPE2`, row 13 | Survives the copy; doesn't break the filter dropdown or column widths |
| `For Approval` rows (yellow) | rows with `Status = For Approval` (3) | Fill treatment; `Status` isn't a default column, so it's dropped |
| `INCLUDED` purchase status | 1 item row | Treated as an item, not a PARENT |
| Bold + italic carried in from the source | item row 8 (whole row), row 31 (`SIZE` cell only) | EQ List output shows both plain, like Split Sheets ([#169](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/169)) |
| `For Approval` item with **no INTERNAL ID or CUSTOMER ID** | row 57 | EQ List output strips its yellow like every other item row: it's an item, not a grouping row ([#170](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/170)) |
| Numeric (not text) tag | one `INTERNAL ID` / `Redbox Row 1` cell | EQ Count numbering and header lookup still work |
| Sparse columns | `PO Number`, `PO Issue Date`, `Ship Date` empty; `Cost`, `Lead Time`, `CUSTOMER ID` mostly empty | Picker still lists them; empty columns don't break formatting |
| Blank row inside the data | row 85 | With **Remove blank rows** on (the default) it's gone from the EQ List output and gets no EQ COUNT number. With it off, it stays ([#171](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/171)) |
| Blank rows at the end | last 14 rows | Last-row detection (`GetLastRow`) ignores them |
| Group-column edge cases for **Split** | `Vendor`: see below | Sheet naming and collisions |

### `Vendor` values for Split Sheet by Column

Splitting on `Vendor` should produce **11 sheets** — 11 distinct values once whitespace and case are ignored. The sheet *names* in this table are checked by the unit tests in `unit/modTest_StringsAndExportNames.bas` (`SanitizeSheetName` and `UniqueRunName`). What those tests don't cover is the split itself — that the right rows land on each sheet and that there are 11 of them — so that part is still a manual check.

| Vendor value in the file | Expected sheet name | Why it's in the fixture |
| ------------------------ | ------------------- | ----------------------- |
| `Alder Valve Co` and ` alder valve co ` (leading/trailing spaces, lower case) | `Alder Valve Co` (one sheet, both rows' items) | Values are normalized before grouping |
| `Acme Valve/Fittings [West]: Ltd?` | `Acme Valve_Fittings _West__ Ltd` | Illegal sheet-name characters become `_`, then truncated to 31 |
| `International Combustion & Controls Corporation of America` | `International Combustion & Cont` | Over 31 characters |
| `International Combustion & Controls Corporation of Asia` | `International Combustion &  (2)` (note the two spaces: the base name is cut to 27 characters to make room for the ` (2)` suffix) | Truncates to the *same* 31 characters as the one above, which appears first in the sheet — exercises collision handling ([#57](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/57)) |
| 3 item rows with `Vendor` blank (plus the PARENT and blank rows) | appear on no vendor sheet | Blank group values are skipped |

The remaining 7 vendors are ordinary names and each gets one sheet. After a split, **Save Each Sheet to XLSX** should write one file per sheet, with the same sanitizing applied to the file names.

## Schedule fixture — what's in it

159 rows under the heading row: **27 summary rows**, **108 task rows** and **24 blank rows** at the end, nested up to 6 levels with Excel row grouping (parents above children) and the `Tasks` text indented by level. Summary rows are bold; the top level is shaded blue, some second-level rows light blue (one red-flagged), and the section rows grey — the fill covers the whole row.

The shipped default Schedule columns (`Status, % Complete, Tasks, Start Date, End Date`) are all present, alongside extras the default list should drop: `Duration`, `Completion Date`, `MANUFACTURER`, `Assigned To`, `Notes` and others.

| Case | Where | What to check |
| ---- | ----- | ------------- |
| **Hidden source columns** | `Computed Status`, `Predecessors`, `Baseline Start`, `Baseline Finish`, `Variance`, `Ready to Start` | The picker still lists them; choosing a hidden column puts it in the output (visible or not — note what the add-in does and whether that is what you want) |
| Real dates and percentages | `Start Date` / `End Date` (`MM/dd/yy`, working days only), `% Complete` (`#,##0%`) | Number formats survive the copy; columns are wide enough that nothing shows as `####` |
| Mixed-type column | `Duration` is `5d` / `2w` text on tasks and a mix of bare numbers and `Nd` on summary rows; one `Completion Date` cell is text, not a date | No type errors, consistent alignment |
| Dependencies as text | `Predecessors`: `12`, `12FS +3d`, `12FF`, `10, 12` | Treated as plain text; no formula or date coercion |
| Empty columns | `Computed Status`, `Notes` have no values | An all-empty column neither breaks formatting nor the column-width autofit |
| Deep nesting | outline levels 0–5 | Grouping and indentation survive the copy |
| Blank rows at the end | last 24 rows | Last-row detection ignores them |

All dates are in 2030 and all names (`Assigned To`, `MANUFACTURER`) are made up.

## Manual smoke test

1. Build and load the add-in (see [CONTRIBUTING.md](../CONTRIBUTING.md#rebuilding-a-testable-xlam-from-src)).
2. Open `EQ_List_Sample.xlsx` → **TPD › One-click › EQ List**, then **Create Customer EQ List** and pick columns. Check the new `Customer EQ List` sheet: only the chosen columns, the TPD header block and logo, grouping and fills carried over, wrapped rows tall enough, no clipped filter dropdowns.
3. On the same file: toggle the EQ List behaviour options in **Set Defaults** and rerun.
4. **Split Sheet by Column** on `Vendor` → check the 11 sheets above. Then **Save Each Sheet to XLSX**.
5. Open `Schedule_Sample.xlsx` → **Schedule** and **Create Customer Schedule**; check the chosen columns (including a hidden one), the date/percent formats, the indentation and the row heights.
6. Delete the generated output (every run adds a new sheet — that is intentional).

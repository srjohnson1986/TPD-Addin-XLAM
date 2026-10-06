# Tests

Two kinds of test live here.

| Folder | What | How it runs |
| ------ | ---- | ----------- |
| `unit/` | [Rubberduck](https://rubberduckvba.com/) unit tests for the pure VBA logic (column lists, preferences, EQ Count numbering) | Rubberduck's Test Explorer, in a workbook that has the add-in's modules loaded (see `docs/ARCHITECTURE.md`) |
| `fixtures/` | **Synthetic** sample workbooks to run the ribbon flows against by hand | Open a fixture, click through the TPD ribbon, compare with the expectations below |

Everything in `fixtures/` is **made up**: company names, tags, part numbers, people and dates. There is no customer data in the repo and there must never be — see [CONTRIBUTING.md](../CONTRIBUTING.md#license-and-conduct).

## Fixtures

Run from the repo root to (re)generate them. They are built from scratch with a fixed seed, so the content is the same every run:

```
pip install openpyxl
python tests/fixtures/generate_fixtures.py
```

| File | Models | Used for |
| ---- | ------ | -------- |
| `fixtures/EQ_List_Sample.xlsx` | A Smartsheet-style equipment-list export: 25 columns, heading in row 1, a nested tree of PARENT and item rows (Excel row grouping, parents above children), row fills by type | Create Customer EQ List / EQ List, Split Sheet by Column / Split Sheets, Save Each Sheet to XLSX |
| `fixtures/Schedule_Sample.xlsx` | A Smartsheet-style project schedule: 9 columns, 3-level task tree, real dates, `%` complete, status | Create Customer Schedule / Schedule |

Edit the **generator**, not the `.xlsx`, then regenerate and commit both. (The `.xlsx` container isn't byte-identical between runs, so expect a binary diff even when nothing changed — only regenerate when you mean to.)

## EQ List fixture — what's in it

100 data rows under the heading row: **16 PARENT rows**, **70 item rows** and **14 blank rows** at the end. Outline levels 0–4. The shipped default EQ List columns (`INTERNAL ID, Location, TYPE1, TYPE2, SIZE, RATING, CONNECTION TYPE, Description, Manufacturer, Model, Details`) are all present, alongside extras the default list should drop (`Status`, `OWNER`, `Vendor`, `Redbox Row 1–3`, `Filename`, `PO Number` …). `Purchased` has `PARENT` / `NO` / `N/R` / `INCLUDED`, so the three PARENT-row toggles on the Set Defaults EQ List page have something to act on.

Deliberately awkward inputs, and where to find them:

| Case | Where | What to check |
| ---- | ----- | ------------- |
| Wrapped, multi-line text | `Details`, rows 22, 43, 62 | After EQ List / Schedule / Split, those rows are tall enough to show all the text, header/logo rows keep their height ([#162](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/162)) |
| Embedded line break in a normal cell | `TYPE2`, row 13 | Survives the copy; doesn't break the filter dropdown or column widths |
| `For Approval` rows (yellow) | rows with `Status = For Approval` (2) | Fill treatment; `Status` isn't a default column, so it's dropped |
| `INCLUDED` purchase status | 1 item row | Treated as an item, not a PARENT |
| Numeric (not text) tag | one `INTERNAL ID` / `Redbox Row 1` cell | EQ Count numbering and header lookup still work |
| Sparse columns | `PO Number`, `PO Issue Date`, `Ship Date` empty; `Cost`, `Lead Time`, `CUSTOMER ID` mostly empty | Picker still lists them; empty columns don't break formatting |
| Blank rows at the end | last 14 rows | Last-row detection (`GetLastRow`) ignores them |
| Group-column edge cases for **Split** | `Vendor`: see below | Sheet naming and collisions |

### `Vendor` values for Split Sheet by Column

Splitting on `Vendor` should produce **11 sheets** — 11 distinct values once whitespace and case are ignored. Expectations below come from reading `modHelpers_Strings.SanitizeSheetName` and `modMain_SplitSheets`; confirm them on the first manual run and correct this table if the add-in disagrees.

| Vendor value in the file | Expected sheet name | Why it's in the fixture |
| ------------------------ | ------------------- | ----------------------- |
| `Alder Valve Co` and ` alder valve co ` (leading/trailing spaces, lower case) | `Alder Valve Co` (one sheet, both rows' items) | Values are normalized before grouping |
| `Acme Valve/Fittings [West]: Ltd?` | `Acme Valve_Fittings _West__ Ltd` | Illegal sheet-name characters become `_`, then truncated to 31 |
| `International Combustion & Controls Corporation of America` | `International Combustion & Cont` | Over 31 characters |
| `International Combustion & Controls Corporation of Asia` | `International Combustion &  (2)` (note the two spaces: the base name is cut to 27 characters to make room for the ` (2)` suffix) | Truncates to the *same* 31 characters as the one above, which appears first in the sheet — exercises collision handling ([#57](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/57)) |
| 3 item rows with `Vendor` blank (plus the PARENT and blank rows) | appear on no vendor sheet | Blank group values are skipped |

The remaining 7 vendors are ordinary names and each gets one sheet. After a split, **Save Each Sheet to XLSX** should write one file per sheet, with the same sanitizing applied to the file names.

## Schedule fixture — what's in it

36 task rows, 3 outline levels (phases, groups, tasks), summary rows bold and shaded. Start/End Date are real Excel dates (mm/dd/yy), `% Complete` is a real percentage, and `Status` has `Complete` / `In Progress` / `Not Started` / `At Risk`. All dates are in 2030, so nothing looks like a real project.

The source column order is `Tasks, Duration, Start Date, End Date, Predecessors, % Complete, Status, Assigned To, Comments` — **not** the default order (`Status, % Complete, Tasks, Start Date, End Date`) — so the output should come out in the *saved list's* order, not the source's ([#127](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/127)). `Comments` has a few long and multi-line cells for row-height autofit.

> The schedule layout is modelled on the add-in's default column list, not on a real export (there was no schedule sample to copy from). If a real schedule export has columns or structure that this misses, add them to the generator.

## Manual smoke test

1. Build and load the add-in (see [CONTRIBUTING.md](../CONTRIBUTING.md#rebuilding-a-testable-xlam-from-src)).
2. Open `EQ_List_Sample.xlsx` → **TPD › One-click › EQ List**, then **Create Customer EQ List** and pick columns. Check the new `Customer EQ List` sheet: only the chosen columns, the TPD header block and logo, grouping and fills carried over, wrapped rows tall enough, no clipped filter dropdowns.
3. On the same file: toggle the EQ List behaviour options in **Set Defaults** and rerun.
4. **Split Sheet by Column** on `Vendor` → check the 11 sheets above. Then **Save Each Sheet to XLSX**.
5. Open `Schedule_Sample.xlsx` → **Schedule** and **Create Customer Schedule**; check column order, date/percent formats and row heights.
6. Delete the generated output (every run adds a new sheet — that is intentional).

# TPD Add-in — User Guide

An Excel add-in that saves time and clicks when turning an internal Smartsheet
export into something ready to send to an external partner: customer EQ lists,
customer schedules, per-vendor RFQ splits, and one-file-per-sheet exports.

Windows desktop Excel only. Works on locally saved `.xlsx` files (files opened
from OneDrive/SharePoint in the browser are out of scope).

- [What it does](#what-it-does)
- [Install](#install)
- [Updating](#updating)
- [The TPD ribbon tab](#the-tpd-ribbon-tab)
- [Set your defaults first](#set-your-defaults-first)
- [The tools](#the-tools)
  - [Create Customer EQ List / EQ List](#create-customer-eq-list--eq-list-one-click)
  - [Create Customer Schedule / Schedule](#create-customer-schedule--schedule-one-click)
  - [Split Sheet by Column / Split Sheets](#split-sheet-by-column--split-sheets-one-click)
  - [Save Each Sheet to XLSX](#save-each-sheet-to-xlsx)
- [How defaults are resolved](#how-defaults-are-resolved)
- [Things that work this way on purpose](#things-that-work-this-way-on-purpose)
- [Troubleshooting](#troubleshooting)
- [Reporting a bug or requesting a feature](#reporting-a-bug-or-requesting-a-feature)
- [Glossary](#glossary)

---

## What it does

Every tool assumes the same starting point: you have exported an **EQ list** or a
**schedule** from Smartsheet as an `.xlsx` and opened it in Excel.

| Tool | What you get |
|---|---|
| **Create Customer EQ List** / **EQ List** | A new `Customer EQ List` sheet: your chosen columns only, a TPD title block with logo and today's date, deformatted rows, frozen header. |
| **Create Customer Schedule** / **Schedule** | The same idea for a schedule — a new `Customer Schedule` sheet with the schedule title block, date, logo, and frozen header. The `Tasks` column keeps its left alignment. |
| **Split Sheet by Column** / **Split Sheets** | One sheet per unique value in a column you pick (e.g. one sheet per **Vendor** for RFQs), each carrying just the columns you chose, with a bold frozen header row. |
| **Save Each Sheet to XLSX** | Saves every sheet in the workbook as its own `.xlsx` file, in a folder named after the workbook, with optional text/date appended to each filename. |

Each of the first three tools runs **two ways**:

- **The picker path** (buttons in the **Sheet Tools** group) opens a dialog so you
  choose columns for this run.
- **The one-click path** (buttons in the **One-click** group) skips the dialog and
  runs straight from your saved defaults.

---

## Install

### 1. Download

Grab the `TPD_Addin.xlam` attached to the newest release on the
**[Releases page](https://github.com/srjohnson1986/TPD-Addin-XLAM/releases)**.

### 2. Save it where Excel keeps add-ins

```
C:\Users\<your username>\AppData\Roaming\Microsoft\AddIns
```

`AppData` is hidden by default. In File Explorer, turn on
**View → Show → Hidden items** to see it. Saving it here is not strictly
required, but it keeps the add-in with every other Excel add-in and makes
updating a straight file swap.

### 3. Enable it in Excel

1. Open Excel.
2. **File → Options → Add-ins**.
3. At the bottom, set **Manage: Excel Add-ins** and click **Go…**.
4. Tick the box next to **TPD_Addin** and click **OK**.
   (If it is not listed, click **Browse…** and point it at the `.xlam` you saved.)
5. Confirm you now see a **TPD** tab on the ribbon.

The first time it loads, Excel may ask you to enable macros / content — allow it.

---

## Updating

1. Close Excel completely.
2. Download the new `TPD_Addin.xlam` from the Releases page.
3. Replace the old file in the `AddIns` folder, **keeping the filename exactly the
   same**. (Rename the old one first if you want a backup.)
4. Start Excel. The new version loads automatically.

Your saved defaults and logo are **not** stored in the `.xlam` — they live in your
Windows user profile — so they carry across an update untouched.

---

## The TPD ribbon tab

One tab, **TPD**, with two groups:

**Sheet Tools** — the picker paths and the export tool:

- Create Customer EQ List
- Create Customer Schedule
- Split Sheet by Column
- Save Each Sheet to XLSX

**One-click** — run straight from saved defaults, plus the settings button:

- EQ List
- Schedule
- Split Sheets
- **Set Defaults** (opens the settings dialog)

---

## Set your defaults first

Click **TPD → One-click → Set Defaults**. This modal dialog is where you tell the
add-in your usual columns and your logo, once, so you don't have to pick them
every run.

Tabs:

| Tab | What it holds |
|---|---|
| **EQ List** | Your default column list for the EQ List tools, plus five behaviour checkboxes (see below). |
| **Schedule** | Your default column list for the Schedule tools. |
| **Split Sheets** | Your default column list **and** the default group-by column for the Split tools. |
| **Logo** | Your header logo — **Choose image…** (PNG / JPG / GIF / BMP) or **Use built-in logo**. |

**Column tabs** hold a plain comma-separated list of heading names. You can paste
a heading row straight out of Excel — tabs and line breaks are converted to
commas automatically when you paste or leave the field. The order you list them
in is the order the generated columns come out (one-click paths).

Each tab shows a short hint under its field as a reminder of what to type.

- **Restore defaults** (per tab) resets that list to the built-in one.
- A column list can't be left empty — if you want the built-in list, use
  **Restore defaults**. (The Split group-column field *may* be left empty.)
- **Save** saves everything and confirms in the status bar. **Cancel**, **Esc**,
  or the **X** discard your changes.
- **Save & Run** saves, closes the dialog, and immediately runs the one-click
  flow for the tab you're on — EQ List, Schedule or Split Sheets — against the
  active sheet, so you can tweak a list and generate in one step. It's hidden on
  the **Logo** tab (nothing to run). Both **Save** and **Save & Run** write all
  four tabs, not just the one you're looking at.
- The dialog reopens on whichever tab you had open last, so you don't have to
  click back to it each time.

**Logo tab:** **Choose image…** or **Use built-in logo** stages a change — you see
it in the preview and a note saying what will happen. It is applied when you click
**Save** (along with the column changes) and used on every generated header from
then on. Your image is copied into `%APPDATA%\TPD_Addin\`, so it survives an Excel
restart and a version update.

Built-in column defaults, for reference:

- **EQ List / Split Sheets:** `INTERNAL ID, Location, TYPE1, TYPE2, SIZE, RATING, CONNECTION TYPE, Description, Manufacturer, Model, Details`
- **Schedule:** `Status, % Complete, Tasks, Start Date, End Date`
- **Split group column:** `Vendor`

### EQ List behaviour checkboxes

On the **EQ List** tab, below the column list. They apply to **every** EQ List
you generate — the picker *and* the one-click button:

| Checkbox | Effect | Default |
|---|---|---|
| **Remove PARENT rows** | Drops every row marked `PARENT` in the Purchased column. | off |
| **Add EQ Count column** | Adds the numbered **EQ COUNT** column as part of the build (same as running EQ Count afterwards). | off |
| **Plain PARENT rows** | White fill, black un-bold text on PARENT rows, instead of the grey grouping shade. | off |
| **Add cell borders** | The thin border around the data range. Turn off for a border-free sheet. | **on** |
| **Add column filters** | Puts Excel's AutoFilter dropdowns on the heading row. | off |

The first three read the **Purchased** column. If the sheet doesn't have one,
those three are skipped (you get a note) and the rest of the list is still
generated. Turning them on does **not** force Purchased into the output — it
appears only if it's in your column list.

---

## The tools

All of these assume the Smartsheet export is the active sheet.

### Create Customer EQ List / EQ List (one-click)

**Picker path — Sheet Tools → Create Customer EQ List:**

1. A dialog lists every column on the sheet as a checkbox. Columns from your saved
   selection are pre-ticked.
2. Tick/untick columns, click **OK** (at least one must be ticked). **Select all**
   / **Select none** / **Restore defaults** above the list are shortcuts —
   *Restore defaults* re-ticks your EQ List default (or the built-in list).
3. A new **Customer EQ List** sheet is added with:
   - only the columns you kept
   - all rows copied first, so Smartsheet's row grouping is preserved
   - a TPD title block (EQUIPMENT LIST / Customer / Project / PM / Date / Rev) with
     the logo top-right and today's date
   - source-row shading stripped, except on `PARENT` rows and rows with no
     `INTERNAL ID` and no `CUSTOMER ID`
   - the header rows frozen

The picker remembers *its own* last selection separately from the Set Defaults
value — clicking OK here does not change your one-click default.

The **EQ List behaviour checkboxes** in Set Defaults (remove PARENT rows, add EQ
Count, plain PARENT rows, borders, column filters) apply here too — they're read
at generate time, not stored with the picker.

**One-click path — One-click → EQ List:** same result, no dialog. Columns come
from your Set Defaults list (then the built-in list). Only the columns from that
list that actually exist on the sheet are kept; if none match, it stops with a
message rather than deleting everything. Columns come out in the order you set on
the EQ List tab.

### Create Customer Schedule / Schedule (one-click)

The schedule twin of the EQ List tools.

**Picker path — Sheet Tools → Create Customer Schedule:** column picker (with the
same **Select all** / **Select none** / **Restore defaults** shortcuts) → new
**Customer Schedule** sheet with the schedule title block, the date, the logo, and
frozen panes. The **Tasks** column keeps its natural left alignment; everything
else is centered.

**One-click path — One-click → Schedule:** runs from your Schedule-tab defaults,
no dialog.

### Split Sheet by Column / Split Sheets (one-click)

Built for making per-vendor RFQs, but works on any column.

**Picker path — Sheet Tools → Split Sheet by Column:**

1. Pick the **group-by column** from the dropdown (e.g. `Vendor`).
2. Tick the columns to copy into each split sheet. **Select all** / **Select
   none** / **Restore defaults** above the list are shortcuts and apply to this
   list only, not the group-by dropdown.
3. **OK.** One new sheet is created per unique non-blank value in the group
   column, each populated with the matching rows (your columns only, in your
   chosen order), a bold frozen header row, and autofit columns.
4. A summary reports how many sheets were created and lists any that failed.

Re-running reuses/clears sheets of the same name. Two values that would collide to
the same sheet name get ` (2)`, ` (3)`, … suffixes so no rows are lost.

**One-click path — One-click → Split Sheets:** uses the group column and column
list from the Split Sheets tab of Set Defaults.

### Save Each Sheet to XLSX

1. **Sheet Tools → Save Each Sheet to XLSX.**
2. A dialog asks for text to append to each filename (optional — e.g. a
   customer name) and whether to append today's date. A **Preview** underneath
   shows the file name the active sheet would be saved as, updating as you type.
3. **OK.** The add-in:
   - makes a folder next to the workbook, named after the workbook file
   - saves each eligible sheet into it as `<sheet name> - <your text>.xlsx`
   - skips the add-in's own internal sheets and hidden/very-hidden sheets
   - overwrites existing files of the same name without asking
   - reports any per-sheet failures in one summary at the end

The workbook must be **saved** first (the export folder is created next to it).

---

## How defaults are resolved

Column lists have three layers. Each tool walks them in order and uses the first
one that is set:

| Tool path | Looks at, in order |
|---|---|
| **Picker** (Create… / Split Sheet by Column) | 1. the picker's own last-used selection → 2. your Set Defaults list → 3. the built-in list |
| **One-click** (EQ List / Schedule / Split Sheets) | 1. your Set Defaults list → 2. the built-in list |

So: clicking **OK** in a picker only affects that picker next time. Changing a list
in **Set Defaults** affects the one-click buttons and pre-fills the pickers, but a
picker whose own last-used selection is set keeps showing that until you change it
there.

Settings are stored per Windows user in the registry
(`HKCU\Software\VB and VBA Program Settings\TPD_Addin\Preferences`) and the logo
under `%APPDATA%\TPD_Addin\` — not in the workbook and not in the `.xlam`.

---

## Things that work this way on purpose

- **Defaults are shared across all workbooks, not per-file.** "Set once, applies
  everywhere." They persist per Windows user.
- **Re-running a command stacks its output.** Running Create Customer EQ List
  twice gives you `Customer EQ List` and `Customer EQ List (1)`. This is
  deliberate — it never overwrites the previous run.

---

## Troubleshooting

**No TPD tab after installing.** Re-check step 3: File → Options → Add-ins →
Manage: Excel Add-ins → Go… → the TPD_Addin box must be ticked. If a startup
error disabled it, it may be under **Manage: Disabled Items**.

**Can't find the AppData folder.** File Explorer → **View → Show → Hidden items**.

**A one-click button says it couldn't find any of your columns.** The heading
names on the sheet don't match the names in your Set Defaults list. Open Set
Defaults and match them to the export's actual headings (or use **Restore
defaults**), or use the picker path for that run.

**The logo looks stretched or missing.** Set a fresh image in Set Defaults →
Logo, or switch to **Use built-in logo**. Very small source images will look soft
when scaled to the header block.

For anything about installing the add-in or showing hidden files, a quick web
search usually settles it fastest. If not, open an issue (see below).

---

## Reporting a bug or requesting a feature

Open an issue on the
**[Issues page](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues)**
(a free GitHub account is all you need). Include:

- the steps to reproduce
- what happened
- what you expected (or want) instead
- the add-in version (shown in the corner of the **Set Defaults** dialog)

Attaching the workbook you were working on helps a lot.

---

## Glossary

| Term | Meaning |
|---|---|
| **EQ list** | Equipment list — the Smartsheet export this add-in is mainly built around. |
| **Schedule** | A project schedule exported from Smartsheet. |
| **RFQ** | Request For Quote — one per vendor, which is what Split Sheet by Column produces. |
| **Header / title block** | The TPD block at the top of a generated sheet (EQUIPMENT LIST / Customer / Project / PM / Date / Rev). |
| **Heading row** | The row of column names for the data table. |
| **Picker** | A dialog that lets you choose columns (and, for Split, the group column) for one run. |
| **One-click** | Running a tool straight from your saved defaults with no dialog. |
| **Defaults** | Your saved column lists, group column, and logo — set in **Set Defaults**. |
| **Add-in (`.xlam`)** | A macro-enabled Excel file that adds commands to Excel without being a workbook you edit. |
| **VBA / macro** | The language and automated routines the add-in is built from. |

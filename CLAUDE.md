# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A macro-enabled Excel add-in (`TPD_Addin.xlam`) for TPD, written in VBA. There is no compiler/CLI toolchain — "building" means running VBA macros inside Excel itself, and testing means exercising the ribbon commands interactively against sample workbooks. Targets Windows desktop Excel only (locally hosted files; Mac and OneDrive/SharePoint-web are out of scope).

The repo tracks VBA **source** (`/src`), not the compiled `.xlam`. The built add-in is published as a versioned asset on GitHub Releases, never committed.

## Repo layout

```
/src        VBA source of truth (.bas / .cls / .frm / .frx), organized into subfolders matching Rubberduck @Folder tags.
            Every UserForm lives in /src/Forms (@Folder("TPD_Addin.Forms")), not beside the flow that shows it, so
            they are all in one place - see the TPD_Addin.Forms section of docs/ARCHITECTURE.md
/customUI   customUI14.xml (ribbon definition) + ribbon icons — lives outside /src, edited directly
/assets     Source brand/logo images not read by the build — a UserForm's Picture property embeds the
            bytes straight into its .frx when set in the VBE, so these are kept so the original art is
            in version control. tpdAddinLogo.* = the wide brand-bar logo; tpdHeaderLogo.jpg = the
            square-ish header logo (matches _Resources/DefaultLogo), loaded into frmSetTPDDefaults'
            imgLogoPreview for the built-in-logo preview (#109)
/build      Local build output — gitignored, never committed (holds the base .xlam and build artifacts)
/docs       ARCHITECTURE.md (module-by-module map) and other dev docs
/tools      TPD_Builder.xlsm — driver workbook whose BuildAddin macro rebuilds the add-in; Build-TPDAddin.ps1 runs it headlessly
```

## Development workflow (no CLI build/lint/test — everything happens in Excel/VBE)

Prerequisites: Windows + Excel, "Trust access to the VBA project object model" enabled (File → Options → Trust Center → Macro Settings), and [Rubberduck](https://rubberduckvba.com/) installed matching Office bitness.

**Making a change:**
1. Edit code in the VBE, test interactively in the open workbook.
2. Export to `/src` — either run `ExportAllVBAModules` (in `modExport_VBAModules`) to refresh the whole tree, or export just the changed component(s) from the VBE Project Explorer (right-click → Export File, overwrite the matching `/src` file).
3. Check `src/export_log.txt` (written by `ExportAllVBAModules`) to confirm what was exported and where.
4. Review the diff before committing — this is the code review step.

**Rebuilding a testable `.xlam`:**
1. Keep a known-good base file at `build/_base/TPD_Addin_base.xlam` (gitignored — supplies worksheets, ribbon, styles, embedded logo shape that live outside `/src`). Must be **stripped of standard modules + UserForms** (the builder can't re-import a form that already exists) — a full add-in `.xlam` is not a valid base. Rarely needs updating.
2. Run the `BuildAddin` macro in `tools/TPD_Builder.xlsm` (a separate driver workbook, not the add-in itself) → copies the base, imports all of `/src`, writes `build/TPD_Addin.xlam`, logs to `build/build.log`. Headless: `powershell -ExecutionPolicy Bypass -File tools\Build-TPDAddin.ps1`.
3. Load it as an add-in (File → Options → Add-ins → Manage: Excel Add-ins → Browse) and exercise it against the sample EQ List fixtures.
4. If it checks out, it's a release candidate.

A clean build is not a passing test — there's no compile step in the macro. For a headless compile check, open the built `.xlam` via COM and `Application.Run` a no-arg no-side-effect function (e.g. `GetTodaysDate`): VBA refuses to run any macro when the project has a compile error, so a clean return means the whole project compiled.

**Cutting a release:** bump `ADDIN_VERSION` in `modStartup`, tag `vX.Y.Z`, publish a GitHub Release with the built `.xlam` attached, add a `CHANGELOG.md` entry, update `docs/USER_GUIDE.md` for any user-visible change and re-sync it to the [wiki](https://github.com/srjohnson1986/TPD-Addin-XLAM/wiki) (`Home.md`). Refresh `build/_base/TPD_Addin_base.xlam` only if the release changed something outside `/src` (ribbon / sheets / logo) — and then from a code-stripped copy, not the release `.xlam` directly. Full steps in `CONTRIBUTING.md`.

## Module organization convention

Every module starts with a Rubberduck folder annotation as a plain VBA comment (the leading `'` is required, or VBA throws a compile error on `@`):

```vb
'@Folder("TPD_Addin.EQList")
Option Explicit
```

This groups modules in Rubberduck's Code Explorer, and the export tooling sorts exported files into matching `/src` subfolders by the same tag. New modules should use an existing folder group, or propose a new one in `docs/ARCHITECTURE.md` if it genuinely doesn't fit.

**UserForms are the one grouping that is by kind, not by feature:** every `.frm` carries `'@Folder("TPD_Addin.Forms")` and exports to `/src/Forms`, whichever feature area shows it. If you add a form, tag it `TPD_Addin.Forms` too — a form tagged with its feature area will export itself back out into that feature's folder.

## Architecture

One custom ribbon tab ("TPD") with two groups — **Sheet Tools** (the picker/custom paths: Create Customer EQ List / Create Customer Schedule / Split Sheet by Column / Save Each Sheet to XLSX) and **One-click** (the one-click "EQ List" / "Schedule" / "Split Sheets" buttons that run straight from saved defaults, beside "Set Defaults") — defined in `customUI/customUI14.xml`. The "EQ Count" flow has no ribbon button (removed for [#120](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/120)) but its code and `RunCountEquipmentRows` callback are kept. Every button's `onAction` is a one-line callback in `modRibbonCallbacks` handing `WithPerformance` the flow's `PERF_*` constant — that module holds them all (Set Defaults excepted: `RunSetTPDDefaults` does real work, so it stays in `modMain_SetTPDDefaults`) and owns the global `gRibbon` reference and `RibbonOnLoad` (triggers `modStartup.InitializeAddIn`, which stamps the running version into the registry-backed preference store; the `_Resources` sheet + embedded logo come from the base `.xlam`).

Feature areas, mirroring the `@Folder("TPD_Addin.X")` groups (full module-by-module map in `docs/ARCHITECTURE.md` — read it before touching a module you haven't seen):

- **Ribbon** — `modRibbonCallbacks` holds every XML `onAction` callback, each one line: `WithPerformance PERF_x`. Watch the naming wart: the Split button's `onAction` is `SplitSheetByColumn`, not `RunSplitSheetByColumn` like its neighbours — the ribbon XML ships inside the base `.xlam`, so renaming it means re-stamping the base file, not just editing `customUI/customUI14.xml`.
- **EQList** — "Create Customer EQ List" (picker) + "EQ List" (one-click, `CreateCustEQListFromDefaults_Internal`) / "EQ Count" flows (`modMain_CustEQList`, `modMain_CountEquipmentRows`). Its column-picker UserForm lives in `/src/Forms`, like every other form.
- **Schedule** — "Create Customer Schedule" (picker, `frmCustScheduleColumnPicker`) + "Schedule" (one-click, `CreateCustScheduleFromDefaults_Internal`) flows in `modMain_CustSchedule` (#112) — the twin of the EQ List flow (copy source → new `Customer Schedule` sheet before formatting, `InsertDefaultCustScheduleHeader` from `modHelpers_SheetSetup`). The picker form is a clone of `frmCustEQListColumnPicker`.
- **SplitExport** — "Split Sheet by Column" and "Save Each Sheet to XLSX" flows. Their UserForms live in `/src/Forms`.
- **Preferences** — per-user key/value settings in the Windows registry (`SaveSetting`/`GetSetting` under `HKCU\…\TPD_Addin\Preferences`), via `modPreferences`. Registry rather than a sheet inside the `.xlam` because Excel never saves an add-in on exit ([#84](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/84)). All keys are centralized in `modPreferences_KeyMap` (`PREF_*` constants) rather than used as loose string literals — follow that pattern for any new preference. A missing key falls back to the caller's default; the built-in column lists live once in `modPreferences_Defaults` (the single source for both pickers and one-click — #98). Column defaults are **three layers** (#96): shipped (`modPreferences_Defaults`) → user one-click default (`DefaultUser*` keys, written **only** by `frmSetTPDDefaults`) → user picker last-used (`LastUsed*` keys, written by the picker forms). `modPreferences.ResolveColumnList` / `ResolveGroupColumn` walk a caller's key array in order: pickers pass `Array(LastUsed*, DefaultUser*)`, one-click passes `Array(DefaultUser*)`. Clicking OK in a picker no longer touches the one-click default. `PREF_LOGO_PATH` stores the user's chosen header logo — a file copied to `%APPDATA%\TPD_Addin\` (#95).
- **Helpers** — shared utilities used by more than one feature area (sheet/workbook ops, header lookup, column filtering, string sanitizing, layout/formatting, checkbox-grid building, logo placement, the batch-run summary box `modHelpers_Reporting.ReportBatchOutcome`) — including the code the dialogs share, so they cannot drift apart: `modHelpers_DialogChrome` (brand band + About links, all five forms) and `modHelpers_ColumnPicker` (`InitColumnPicker` builds and pre-fills a picker grid, `CommitColumnPicker` is its OK-guard and `LastUsed*` save — the three pickers call one of each). `modGdiPlus.LoadPictureGDIP` is the **only `Declare` in the codebase** — a `#If VBA7`-guarded GDI+ picture loader for the logo preview (VBA's `LoadPicture` can't read PNG).
- **Core** — add-in lifecycle/infra: `modPerformance.WithPerformance` (screen updating / events / calc mode wrapper around a routine), `modStartup` (init sequencing), `modExport_VBAModules` (the export macro `ExportAllVBAModules`).
- **Document** — code-behind for `ThisWorkbook`/`Sheet1`-`3`. These are document modules: unlike standard modules they can't be removed and re-imported normally — the build macro clears and re-pastes their code text instead of a plain `VBComponents.Import`.

## Design intent vs. bugs (don't "fix" these without asking)

- **Global preferences are shared across workbooks, not per-workbook** — "set once, applies everywhere" is intentional. They persist per Windows user (registry), not per machine or per file.
- **Re-running a command stacks its output** (e.g. running "Create Customer EQ List" twice adds a second sheet) — intentional.

## Known open work

Tracked as GitHub Issues — **there are currently no open issues.**

The **Set TPD Defaults dialog ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25))** cluster is delivered and shipped in `v2.4.2`: `frmSetTPDDefaults` (EQ List / Schedule / Split Sheets / Logo tabs, Save / Save & Run / Cancel footer), the three picker + one-click flow pairs and their three-layer column defaults (#96 / #98 / #112), the user-settable header logo (#95 / #109), the EQ List behaviour toggles (#97 → #119–#123), the dialog visual refresh across all five forms (#118 / #143 / #99 / #100), the picker Select all / none / Restore defaults row (#149) and the picker grid scrolling (#150).

Two names are **load-bearing strings** that no compiler checks — rename either and something breaks silently:

- `ClearTPDDefaultsStatusBar` — `frmSetTPDDefaults` schedules it via `Application.OnTime` by string literal (the only such reference in the codebase). Rename it and the "TPD defaults saved" status bar never clears.
- Anything in a `customUI14.xml` `onAction` — the ribbon XML ships inside the base `.xlam`, so a rename needs the base re-stamped, not just the repo's XML edited. This is why the Split button's callback is `SplitSheetByColumn` rather than `RunSplitSheetByColumn`.

### Regressions to guard against (all fixed — don't undo them)

- **`modPerformance.WithPerformance`** ([#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7) / [#36](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/36)): the `CleanUp` handler must **always** restore `ScreenUpdating` / `EnableEvents` / `Calculation`, or they stay broken for the Excel session. Callers pass a `PERF_*` constant, not a bare string — a new flow needs a constant, a `Case` in the `Select`, and the `WithPerformance PERF_x` call in its callback.
- **Never call a `Sub` as `MySub (obj)`** ([#29](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/29)) — the parens force by-value evaluation (the default property), so a `Worksheet`/`Workbook` argument raises run-time 438. Use `MySub obj` or `Call MySub(obj)`. The `Split Sheet by Column` loop ([#32](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/32)) now reports per-value failures in its summary, so a regression there is at least visible.

## File-format conventions (`.gitattributes`)

- `.bas` / `.cls` / `.frm` / `.xml` are text, forced to CRLF (matches what the VBE always writes on export — keeps diffs to real code changes).
- `.frx` (UserForm binary resources) and `.xlam`/`.xlsm`/`.xlsx`/`.png`/`.ico`/`.bmp` are binary — never diff/merge/line-ending-convert these.

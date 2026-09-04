# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A macro-enabled Excel add-in (`TPD_Addin.xlam`) for TPD, written in VBA. There is no compiler/CLI toolchain — "building" means running VBA macros inside Excel itself, and testing means exercising the ribbon commands interactively against sample workbooks. Targets Windows desktop Excel only (locally hosted files; Mac and OneDrive/SharePoint-web are out of scope).

The repo tracks VBA **source** (`/src`), not the compiled `.xlam`. The built add-in is published as a versioned asset on GitHub Releases, never committed.

## Repo layout

```
/src        VBA source of truth (.bas / .cls / .frm / .frx), organized into subfolders matching Rubberduck @Folder tags
/customUI   customUI14.xml (ribbon definition) + ribbon icons — lives outside /src, edited directly
/assets     Source brand/logo images (e.g. tpdAddinLogo.png/.jpg) not read by the build — a UserForm's
            Picture property embeds the bytes straight into its .frx when set in the VBE, so these are
            kept only so the original art is in version control for future edits
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

**Cutting a release:** bump `ADDIN_VERSION` in `modStartup`, tag `vX.Y.Z`, publish a GitHub Release with the built `.xlam` attached, add a `CHANGELOG.md` entry. Refresh `build/_base/TPD_Addin_base.xlam` only if the release changed something outside `/src` (ribbon / sheets / logo) — and then from a code-stripped copy, not the release `.xlam` directly.

## Module organization convention

Every module starts with a Rubberduck folder annotation as a plain VBA comment (the leading `'` is required, or VBA throws a compile error on `@`):

```vb
'@Folder("TPD_Addin.EQList")
Option Explicit
```

This groups modules in Rubberduck's Code Explorer, and the export tooling sorts exported files into matching `/src` subfolders by the same tag. New modules should use an existing folder group, or propose a new one in `docs/ARCHITECTURE.md` if it genuinely doesn't fit.

## Architecture

One custom ribbon tab ("TPD") with two groups — **EQ List Tools** and **Schedule Tools** — defined in `customUI/customUI14.xml`. Ribbon buttons call thin wrapper subs in `modRibbonCallbacks`, which own the global `gRibbon` reference and `RibbonOnLoad` (triggers `modStartup.InitializeAddIn`, which stamps the running version into the registry-backed preference store; the `_Resources` sheet + embedded logo come from the base `.xlam`).

Feature areas, mirroring the `@Folder("TPD_Addin.X")` groups (full module-by-module map in `docs/ARCHITECTURE.md` — read it before touching a module you haven't seen):

- **Ribbon** — `modRibbonCallbacks` bridges XML `onAction` callbacks to real entry points.
- **EQList** — "Create Customer EQ List" / "EQ Count" flows (`modMain_CustEQList`, `modMain_CountEquipmentRows`), plus the column-picker UserForm.
- **Schedule** — schedule header insertion; the main schedule automation module is currently a placeholder.
- **SplitExport** — "Split Sheet by Column" and "Save Each Sheet to XLSX" flows, plus their UserForms.
- **Preferences** — per-user key/value settings in the Windows registry (`SaveSetting`/`GetSetting` under `HKCU\…\TPD_Addin\Preferences`), via `modPreferences`. Registry rather than a sheet inside the `.xlam` because Excel never saves an add-in on exit ([#84](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/84)). All keys are centralized in `modPreferences_KeyMap` (`PREF_*` constants) rather than used as loose string literals — follow that pattern for any new preference. A missing key falls back to the caller's default; the built-in column lists live in `modPreferences_Defaults`. The four `DefaultUser*` column keys are one shared set — `frmSetTPDDefaults` (the Set TPD Defaults dialog, `modMain_SetTPDDefaults`) sets them, the per-run pickers pre-fill from them and write the user's choice back.
- **Helpers** — shared utilities used by more than one feature area (sheet/workbook ops, header lookup, column filtering, string sanitizing, layout/formatting, checkbox-grid building).
- **Core** — add-in lifecycle/infra: `modPerformance.WithPerformance` (screen updating / events / calc mode wrapper around a routine), `modStartup` (init sequencing), `modExport_VBAModules` (the export macro `ExportAllVBAModules`).
- **Document** — code-behind for `ThisWorkbook`/`Sheet1`-`3`. These are document modules: unlike standard modules they can't be removed and re-imported normally — the build macro clears and re-pastes their code text instead of a plain `VBComponents.Import`.

## Design intent vs. bugs (don't "fix" these without asking)

- **Global preferences are shared across workbooks, not per-workbook** — "set once, applies everywhere" is intentional. They persist per Windows user (registry), not per machine or per file.
- **Re-running a command stacks its output** (e.g. running "Create Customer EQ List" twice adds a second sheet) — intentional.

## Known open work

Tracked as GitHub Issues — no open `bug`-labelled issues.

- **Set TPD Defaults dialog ([#25](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/25))** — **partially delivered.** `frmSetTPDDefaults` is now a working modal (ribbon: TPD → Settings → Set TPD Defaults) with EQ List / Schedule / Split Sheets / Logo tabs. Each column tab is a plain comma-separated MultiLine TextBox (paste a heading row; tabs/newlines → commas), normalized on paste + on exit; per-tab Restore defaults; OK → `SaveAllDefaults` + status-bar confirm; Cancel/Esc/X discard. It writes the shared `DefaultUser*` keys the pickers pre-fill from — the pickers stay as the in-the-moment override (single-key model, no `_CUSTOM` keys, no "save as default" checkbox). Logo tab is a read-only preview of the embedded `_Resources` logo. Open follow-ups: user-settable logo (Browse → embed); headless one-click feature buttons; wiring the EQ List behaviour toggles sketched on the form; moving the pickers' inline default arrays onto `modPreferences_Defaults`; the `txtImgPath` / `cmdBrowse` vestige on `CustEQListColumnPickerForm` from [#82](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/82).

### Regressions to guard against (all fixed — don't undo them)

- **`modPerformance.WithPerformance`** ([#7](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/7) / [#36](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/36)): the `CleanUp` handler must **always** restore `ScreenUpdating` / `EnableEvents` / `Calculation`, or they stay broken for the Excel session. Callers pass a `PERF_*` constant, not a bare string — a new flow needs a constant, a `Case` in the `Select`, and the `WithPerformance PERF_x` call in its callback.
- **Never call a `Sub` as `MySub (obj)`** ([#29](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/29)) — the parens force by-value evaluation (the default property), so a `Worksheet`/`Workbook` argument raises run-time 438. Use `MySub obj` or `Call MySub(obj)`. The `Split Sheet by Column` loop ([#32](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/32)) now reports per-value failures in its summary, so a regression there is at least visible.

## File-format conventions (`.gitattributes`)

- `.bas` / `.cls` / `.frm` / `.xml` are text, forced to CRLF (matches what the VBE always writes on export — keeps diffs to real code changes).
- `.frx` (UserForm binary resources) and `.xlam`/`.xlsm`/`.xlsx`/`.png`/`.ico`/`.bmp` are binary — never diff/merge/line-ending-convert these.

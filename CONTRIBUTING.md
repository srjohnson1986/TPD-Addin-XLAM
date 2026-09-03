# Contributing to TPD-Addin-XLAM

This document covers the developer workflow: how the VBA source is organized, how to make and export changes, and how to build and release a testable `.xlam`. For install/usage/troubleshooting help as an end user of the add-in, see the [TPD Add-in documentation](https://docs.google.com/document/d/1OaWXSklF3Ry4mUrTqPOSVHrjU-0SlxQ_mZ4MlydHqO8/edit?usp=sharing) instead.

## Prerequisites

- **Windows + Excel.** This add-in targets locally hosted files on Windows desktop Excel only. Mac and cloud-hosted (OneDrive/SharePoint web) paths are out of scope.
- **Trust access to the VBA project object model** enabled: File → Options → Trust Center → Trust Center Settings → Macro Settings. Required for exporting/importing code and for the build macro.
- **[Rubberduck](https://rubberduckvba.com/)** installed, matching your Office bitness (32-bit vs 64-bit — check File → Account → About Excel).
- **GitHub Desktop** (or another git client) connected to this repo.

## Repo layout

```
/src        VBA source - the source of truth (.bas / .cls / .frm / .frx)
/customUI   customUI14.xml (ribbon definition) + ribbon icons
/build      Local build output - gitignored, never committed
/docs       Developer documentation (this file, ARCHITECTURE.md, etc.)
/tests      Sample EQ List fixtures + test plan
```

The compiled `.xlam` is a **build artifact**, not source. It's never committed to the repo history — it's published as a downloadable asset on [GitHub Releases](../../releases) instead. See "Cutting a release" below.

## How modules are organized

Every module starts with a Rubberduck folder annotation, written as a plain VBA comment (the leading `'` matters — without it, VBA will throw a compile error on the `@`):

```vb
'@Folder("TPD_Addin.EQList")
Option Explicit
```

This groups modules in Rubberduck's Code Explorer to match the feature areas of the add-in, and the export tooling below uses the same tag to sort exported files into matching subfolders under `/src`. See **[ARCHITECTURE.md](ARCHITECTURE.md)** for the current, complete folder map and what each module does.

If you add a new module, give it a `@Folder` tag matching one of the existing groups (or propose a new one in `ARCHITECTURE.md` if it genuinely doesn't fit).

## Making a change

1. Edit code in the VBE as normal, and test interactively in the open workbook.
2. When you're happy with a change, export it to `/src`:
   - Run `ExportAllVBAModules` (in `modExport_VBAModules`) to export everything and refresh the whole `/src` tree, **or**
   - Select just the component(s) you changed in the VBE's Project Explorer → right-click → **Export File** → overwrite the matching file(s) in your local `/src`.
3. Check `export_log.txt` (written to `/src` by `ExportAllVBAModules`) to confirm what was exported and where.
4. In GitHub Desktop, review the diff for each changed file — this is your code review moment, even working solo.
5. Commit with a message describing the change (not "updated code"). For anything non-trivial, push to a feature branch and open a PR against `main` rather than committing straight to `main`.

## Rebuilding a testable `.xlam` from `/src`

1. Keep a known-good "base" file at `build/TPD_Addin_base.xlam` (gitignored — this supplies the worksheets, ribbon, styles, and embedded logo shape that live outside `/src`).
2. From a **separate driver workbook** (not the add-in itself), run the `BuildAddinFromSource` macro against `/src` and the base file. This produces `build/TPD_Addin.xlam`.
3. Load that file as an add-in (File → Options → Add-ins → Manage: Excel Add-ins → Browse) and run it against the sample EQ List fixtures in `/tests`.
4. If it checks out, this is your release candidate.

## Cutting a release

1. Bump the version (tracked via `PREF_VERSION` and/or `docProps`).
2. Tag the commit (`vX.Y.Z`) and publish a GitHub Release with the built `.xlam` attached as an asset — this is what the README's Download link points to.
3. Add an entry to `CHANGELOG.md` describing what changed.
4. That release's `.xlam` becomes the new `build/TPD_Addin_base.xlam` for next time.

## Known constraints to keep in mind

- **Document modules** (`ThisWorkbook`, `Sheet1`, `Sheet2`, `Sheet3`) can't be removed and re-imported like standard modules — the build macro clears and re-pastes their code text instead of doing a normal `VBComponents.Import`.
- **`customUI14.xml`** (the ribbon) and any static worksheet data (e.g. the embedded default logo shape on `_Resources`) live outside `/src` entirely. They're edited directly in the base file, or with the [Custom UI Editor for Microsoft Office](https://github.com/OfficeDev/office-custom-ui-editor) for ribbon XML specifically.
- **Global preferences are shared, not per-workbook** — "set once, applies everywhere" is intentional, not a bug.
- **Re-running a command stacks its output** (e.g. running "Create Customer EQ List" twice adds a second sheet) — also intentional.

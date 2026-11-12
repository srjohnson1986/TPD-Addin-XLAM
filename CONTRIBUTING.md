# Contributing to TPD-Addin-XLAM

This document covers the developer workflow: how the VBA source is organized, how to make and export changes, and how to build and release a testable `.xlam`. For install/usage/troubleshooting help as an end user of the add-in, see [docs/USER_GUIDE.md](docs/USER_GUIDE.md) (mirrored to the [GitHub wiki](https://github.com/srjohnson1986/TPD-Addin-XLAM/wiki)) instead.

## Prerequisites

- **Windows + Excel.** This add-in targets locally hosted files on Windows desktop Excel only. Mac and cloud-hosted (OneDrive/SharePoint web) paths are out of scope.
- **Trust access to the VBA project object model** enabled: File → Options → Trust Center → Trust Center Settings → Macro Settings. Required for exporting/importing code and for the build macro.
- **[Rubberduck](https://rubberduckvba.com/)** installed, matching your Office bitness (32-bit vs 64-bit — check File → Account → About Excel).
- **GitHub Desktop** (or another git client) connected to this repo.

## License and conduct

The source is released under the [MIT License](LICENSE); by contributing you agree that your contribution is licensed under it too. The TPD name and logo artwork are **not** covered by that license (see [assets/README.md](assets/README.md)) — don't add other third-party or customer material to the repo, and use only fabricated data in anything you commit (fixtures, screenshots, bug-report attachments). If you paste or adapt code from elsewhere, note where it came from and make sure its license allows it.

Everyone taking part is expected to follow the [Code of Conduct](CODE_OF_CONDUCT.md). Security problems go to the private channel in [SECURITY.md](SECURITY.md), not a public issue.

## Repo layout

```
/src        VBA source - the source of truth (.bas / .cls / .frm / .frx);
            every UserForm lives in /src/Forms
/customUI   customUI14.xml (ribbon definition) + ribbon icons
/assets     Source brand/logo images (TPD's - not under the MIT license)
/build      Local build output - gitignored, never committed
/docs       Developer documentation (this file, ARCHITECTURE.md, etc.)
/tests      Rubberduck unit tests (tests/unit) + synthetic EQ List / Schedule
            fixture workbooks (tests/fixtures) - see tests/README.md
/tools      Build / drift-check / static-check scripts + the TPD_Builder.xlsm driver
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

**UserForms group by kind, not by feature.** Every `.frm` is tagged `'@Folder("TPD_Addin.Forms")` and therefore exports to `/src/Forms`, no matter which feature area shows it — so all five dialogs sit together instead of being scattered across four folders. Tag any new form the same way; tag it with its feature area instead and the next export will move it out of `/src/Forms`.

## Making a change

1. Edit code in the VBE as normal, and test interactively in the open workbook.
2. When you're happy with a change, export it to `/src`:
   - Run `ExportAllVBAModules` (in `modExport_VBAModules`) to export everything and refresh the whole `/src` tree, **or**
   - Select just the component(s) you changed in the VBE's Project Explorer → right-click → **Export File** → overwrite the matching file(s) in your local `/src`.
   `ExportAllVBAModules` finds your clone's `/src` itself — it walks up from the folder the workbook is saved in until it finds a folder holding both `src` and `customUI` (a build in `build/` is found immediately). If the workbook lives somewhere else, it asks you to pick the `src` folder instead. Nothing needs editing.
3. Check `export_log.txt` (written to `/src` by `ExportAllVBAModules`) to confirm what was exported and where. It's gitignored — a local record only, never committed.
4. In GitHub Desktop, review the diff for each changed file — this is your code review moment, even working solo.
5. Commit with a message describing the change (not "updated code"), push to a short-lived branch and open a PR against `main`. `main` does not accept direct pushes — see "Merging" below.

## Merging

A ruleset on `main` ("main: pull request + static checks") requires that every change arrives through a pull request and that the `static-checks` job has passed on it. No approving review is required, so working solo you can merge your own PR as soon as the check is green. Squash is the only merge method, and the head branch is deleted on GitHub when the PR merges.

- **Auto-merge** is enabled. `gh pr merge --auto --squash` (or **Enable auto-merge** on the PR page) queues the PR to merge by itself the moment the check passes, so you don't have to wait and click.
- The PR branch does **not** have to be up to date with `main` to merge. If you want it current anyway, use **Update branch** on the PR page.
- New PRs start from `.github/pull_request_template.md`. The check only covers what can be seen in the text of the source; the template's checklist is where you record the Excel-side testing (compile, Test Explorer, fixtures) that no runner can do.
- The ruleset has no bypass list, so it applies to repository admins too. To change it: **Settings → Rules → Rulesets**.

## Working with UserForms (`.frm` / `.frx`)

A form's control layout is placed by hand in the VBE and lives in the binary
`.frx`; its code-behind is plain text at the bottom of the `.frm`. **Exporting a
form writes the whole `.frm` from whatever the VBE currently holds** — so if you
place controls in a VBE session that doesn't have the latest code-behind and then
export, the export silently overwrites that code-behind with the older version.

When the control work and the code-behind are done by different people (e.g. you
place controls, Claude writes the handlers), do them in this order, **per form,
per change**:

1. **Controls first** — place / rename / resize controls, set design-time
   properties, export `.frm` + `.frx` to `/src`.
2. **Code-behind second** — written on top of that exported `.frm`, along with any
   `.bas` modules.
3. **Build from `/src`** — `BuildAddin` imports everything fresh, so `/src` is
   authoritative regardless of what any VBE project holds.
4. **Don't re-export that form** after step 2 unless you first re-import the
   `/src` `.frm` into the VBE. Whoever touches a form last for a given change
   wins.

If step 2 turns up a control that needs adding, list it and go back to step 1 —
same order, just a loop. A **code-only** form change (no control touched) doesn't
need step 1. `.bas` modules have none of this hazard — edit them any time.

## Rebuilding a testable `.xlam` from `/src`

1. Keep a known-good **base** file at `build/_base/TPD_Addin_base.xlam` (gitignored — supplies the worksheets, ribbon, `_Resources` sheet + embedded logo, and styles that live outside `/src`). It must be **stripped of standard code modules and UserForms** — the builder replaces standard modules cleanly but chokes trying to re-import a form that already exists, so a full add-in `.xlam` is *not* a valid base. In practice the base almost never changes; update it only when the non-`/src` content does (see "Cutting a release").
2. Run the `BuildAddin` macro in **`tools/TPD_Builder.xlsm`** (a separate driver workbook, not the add-in itself). It copies the base file, imports every module from `/src`, and writes `build/TPD_Addin.xlam`, logging to `build/build.log`. A healthy build reports every `/src` component "imported/injected, 0 skipped".
   - Headless: `powershell -ExecutionPolicy Bypass -File tools\Build-TPDAddin.ps1` drives that macro via COM (paths default to this repo). Requires Trust Center → Macro Settings → "Trust access to the VBA project object model". A clean run means *a build exists*, not that it's good.
3. Load `build/TPD_Addin.xlam` as an add-in (File → Options → Add-ins → Manage: Excel Add-ins → Browse) and exercise the ribbon flows against the synthetic fixtures in `/tests/fixtures` — the manual smoke test and what to expect are in [tests/README.md](tests/README.md). Never use real customer workbooks in anything you commit.
4. `powershell -ExecutionPolicy Bypass -File tools\Test-SourceDrift.ps1` — confirms `/src` is exactly what got built into `build/TPD_Addin.xlam` (catches a VBE-only edit, or a control placed in a form, that never made it back into `/src`). Needs COM/Office like the build itself, so it's a local check, not part of CI. See "Automated checks" below.
5. `powershell -ExecutionPolicy Bypass -File tools\Test-Unit.ps1` — runs the unit tests in `/tests/unit` against that build, headlessly (add `-Build` to do step 2 first). Also a local check: it needs Excel. See [tests/README.md](tests/README.md#unit-tests).
6. If it checks out, this is your release candidate.

**All of the automated part in one command:** `powershell -ExecutionPolicy Bypass -File tools\Test-All.ps1` runs the static checks, steps 2, 4 and 5 above, and two checks of its own, in order, stopping at the first failure:

- nothing has the built add-in or the builder workbook open (the build can't overwrite a loaded add-in, so close Excel or unload the add-in first);
- the ribbon XML inside the base `.xlam`, and inside the build, is the same as `customUI/customUI14.xml` — an edit to the repo's XML that was never stamped into the base file fails here instead of shipping the old ribbon.

It ends with a short summary block to paste into the PR. Each step that drives Excel runs under a time limit; if one hangs, the script prints the text of the dialog Excel is showing, ends the Excel that step started (and no other), and stops. `-NoExcel` runs only the steps that need no Excel, for when Excel is busy — a partial result, and it says so.

It does not replace step 3, **Debug → Compile VBAProject**, or a Test Explorer run. Those still need a person.

## Automated checks

`.github/workflows/static-checks.yml` runs `tools/Test-StaticChecks.ps1` on every push/PR against `main`, on a normal GitHub-hosted runner (no Excel/COM needed — it's pure text/XML parsing). It covers `/src` and `/tests/unit`, and checks:

1. **Ribbon callbacks** — every `onAction`/`onLoad` named in `customUI/customUI14.xml` has a matching `Sub`/`Function` in `/src`.
2. **Option Explicit** — every module with code declares it (code-free document modules are exempt).
3. **@Folder tags** — every module has a well-formed `'@Folder("TPD_Addin....")` tag, and every `.frm` is tagged exactly `'@Folder("TPD_Addin.Forms")`.
4. **File shape** — CRLF line endings, no BOM, `Attribute VB_Name` matches the file name, a `.frm` header names its own form and `.frx`, the `.frx` exists, and no two modules share a name. (A hand-written `.frm` that gets this wrong is imported by the builder as a plain module.)
5. **Declarations first** — no module-level code after the first `Sub`/`Function`/`Property`. A `Const` or `Dim` down there doesn't resolve, and nothing says so until a flow reaches it at run time.
6. **Parenthesised calls** — no `MySub (arg)` statement for a `Sub`/`Function` declared in this repo; the parens force by-value evaluation, so an object argument raises run-time 438 ([#29](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/29)).
7. **String-named procedures** — every string literal passed to `Application.OnTime` / `Application.Run` names a `Sub`/`Function` in a standard module (today that is `ClearTPDDefaultsStatusBar`).
8. **WithPerformance wiring** — every `PERF_*` constant has a `Case` that calls the `Sub` its value names, that `Sub` exists, and some caller passes the constant.
9. **Constants declared** — every `UPPER_SNAKE` identifier used in code is declared in that module or `Public` in a standard module, so a mistyped `PREF_*` / `PERF_*` is caught without a compile.
10. **Version sync** — `ADDIN_VERSION`, the README's pinned download link and the newest `CHANGELOG.md` release heading all name the same version. A release PR has to change all three together.

Run it yourself any time with `powershell -ExecutionPolicy Bypass -File tools\Test-StaticChecks.ps1` — same script, same output, locally or in Actions.

These are text heuristics, not a compiler: they cover the mistakes this project has actually made, and **Debug → Compile VBAProject** in the VBE is still the only real compile check. The line-ending part of check 4 can only fail locally — a git checkout always writes CRLF for these files — so it is there for the file you haven't committed yet.

**Changing a check?** Run `tools\Test-StaticChecks.ps1 -SelfTest` afterwards (CI runs it too, before the real checks). It builds a small synthetic source tree, confirms it passes clean, then breaks it one way at a time and confirms each break is reported by exactly the expected check. Add a case there for any new check — a check that has never been seen to fail proves nothing.

It deliberately does **not** check that `/src` matches a built `.xlam`, and it does not run the unit tests — both need Excel/COM, which hosted runners don't have. Those are `tools/Test-SourceDrift.ps1` and `tools/Test-Unit.ps1`, steps 4 and 5 of "Rebuilding a testable `.xlam`" above: local gates, not CI jobs. (A self-hosted runner on a personal Windows machine could run it in Actions too, but that means GitHub Actions can execute workflow code — including from your own pushes — directly on that machine, so this repo keeps it local instead.)

## Cutting a release

1. Bump `ADDIN_VERSION` in `modStartup` (and the workbook's `docProps` if you keep those in sync). `modPreferences_Initializer` stamps `ADDIN_VERSION` into the registry as `PREF_VERSION` on the next run. The README's pinned download link tracks this automatically — a Claude Code `PostToolUse` hook (`.claude/settings.json`) runs `tools/sync-readme-version.sh` whenever `modStartup.bas` changes; if you edit the version another way, run that script yourself and commit the README change.
2. Add an entry to `CHANGELOG.md` describing what changed: a `## [X.Y.Z] - YYYY-MM-DD` section, with nothing left under `## [Unreleased]`. That section becomes the release notes word for word. Steps 1, 2 and the user-guide part of step 5 go in one release PR, since the static checks require the version, the README link and the newest CHANGELOG heading to agree.
3. Once that PR has merged, on an up-to-date `main` run `powershell -ExecutionPolicy Bypass -File tools\New-Release.ps1`. It checks the repo is ready, runs `tools\Test-All.ps1`, and creates a **draft** GitHub release `vX.Y.Z` with `build\TPD_Addin.xlam` attached and the CHANGELOG section (plus an install footer with the file's SHA-256) as its notes. Run it with `-DryRun` first to see every problem and exactly what it would create, without creating anything.
   - A draft has no tag and is not public. **Publishing it on GitHub is your step**: read the draft, do the manual checks (Debug → Compile VBAProject, the fixtures smoke test), then click Publish. That creates the `vX.Y.Z` tag.
   - The README's pinned link points at `releases/download/vX.Y.Z/TPD_Addin.xlam`, so the tag name and asset name must match exactly. The script always uses those.
4. Refresh the base **only if this release changed something outside `/src`** (ribbon XML, worksheets, styles, the `_Resources` logo shape). To do it, take a copy of the release `.xlam`, delete every standard module and UserForm from its VBA project (leaving the document modules), and save that as `build/_base/TPD_Addin_base.xlam`. If the release was `/src`-only, the existing base is still current — leave it.
5. **Update the user guide; the wiki follows by itself.** Bring [docs/USER_GUIDE.md](docs/USER_GUIDE.md) in line with anything user-visible in this release (ribbon, dialogs, behaviour), in the release PR. When you publish the release, `.github/workflows/wiki-sync.yml` regenerates the wiki's `Home.md` from the guide as it is at that release's tag, and pushes it. It runs on publish, not on every merge, so the wiki always describes the version people can download.
   - `Home.md` is built by `tools/Sync-WikiHome.ps1` as: the guide's `# TPD Add-in — User Guide` title, then the **"Do not edit this page in the wiki editor"** notice copied from the current `Home.md`, then the rest of the guide as written. If the current page has no such notice the script stops and writes nothing.
   - To fix a typo in the guide between releases, merge the fix and run the workflow by hand (**Actions → Sync wiki from user guide → Run workflow**); that syncs from `main`.
   - To do it without Actions: clone `https://github.com/srjohnson1986/TPD-Addin-XLAM.wiki.git`, run `tools\Sync-WikiHome.ps1 -GuidePath docs\USER_GUIDE.md -HomePath <clone>\Home.md`, commit and push. `-Check` reports whether the wiki is out of date without changing it.

   Only `Home.md` is generated this way. `_Sidebar.md` and any other wiki pages are hand-maintained — don't touch them. (The wiki repo exists only after its first page has been created via the GitHub web UI.)

## Known constraints to keep in mind

- **Document modules** (`ThisWorkbook`, `Sheet1`, `Sheet2`, `Sheet3`) can't be removed and re-imported like standard modules — the build macro clears and re-pastes their code text instead of doing a normal `VBComponents.Import`.
- **`customUI14.xml`** (the ribbon) and any static worksheet data (e.g. the embedded default logo shape on `_Resources`) live outside `/src` entirely. They're edited directly in the base file, or with the [Custom UI Editor for Microsoft Office](https://github.com/OfficeDev/office-custom-ui-editor) for ribbon XML specifically.
- **Global preferences are shared, not per-workbook** — "set once, applies everywhere" is intentional, not a bug.
- **Re-running a command stacks its output** (e.g. running "Create Customer EQ List" twice adds a second sheet) — also intentional.

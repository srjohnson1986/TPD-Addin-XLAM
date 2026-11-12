## What changed

<!-- What this PR does, in a sentence or two. Link the issue: Fixes #123 -->

## Why

## How it was tested

Tick what was done; write "n/a" beside anything that does not apply to this change. CI only runs the static checks - everything below that line needs Excel and a person.

- [ ] `tools\Test-StaticChecks.ps1` passes locally (CI runs it too)
- [ ] Built from this branch (`tools\Build-TPDAddin.ps1`): every `/src` component imported, 0 skipped
- [ ] **Debug > Compile VBAProject** run in the VBE on that build
- [ ] Rubberduck Test Explorer: Run All Tests is green (needed when pure logic or `tests/unit` changed)
- [ ] Ribbon flows this PR touches run against the fixtures in `tests/fixtures` (see `tests/README.md`)
- [ ] UserForm changed: controls were placed and exported first, code-behind written on top of that export

## Docs

- [ ] `CHANGELOG.md` `[Unreleased]` entry (user-visible changes)
- [ ] `docs/ARCHITECTURE.md` (module added, moved, or its job changed)
- [ ] `docs/USER_GUIDE.md` (ribbon, dialogs, or behaviour changed)

## Reviewer notes

<!-- Anything not verified, anything intentionally left out, anything that looks odd but is deliberate. -->

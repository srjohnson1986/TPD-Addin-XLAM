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

### Fixed

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

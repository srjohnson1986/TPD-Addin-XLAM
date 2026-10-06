# TPD-Addin-XLAM

A macro-enabled Excel add-in (`TPD_Addin.xlam`, written in VBA) that turns a raw equipment/schedule export into something ready to send to an external partner: customer EQ lists, customer schedules, per-value sheet splits, and one-file-per-sheet exports. It adds a **TPD** tab to the Excel ribbon.

Windows desktop Excel only. Works on locally saved workbooks.

## Download

<!-- download-link: version managed by tools/sync-readme-version.sh -->
**[Download `TPD_Addin.xlam` — v2.4.4](https://github.com/srjohnson1986/TPD-Addin-XLAM/releases/download/v2.4.4/TPD_Addin.xlam)**

Older builds are on the **[Releases page](https://github.com/srjohnson1986/TPD-Addin-XLAM/releases)**. Install the `.xlam` per the guide below.

Because the add-in contains macros, Excel will ask you to trust it. The full source of everything it does is in this repo (`/src`), so you can read it, or [build it yourself](CONTRIBUTING.md#rebuilding-a-testable-xlam-from-src) rather than trusting a download.

## Documentation

- **[User Guide](docs/USER_GUIDE.md)** — install, usage and troubleshooting (also on the [wiki](https://github.com/srjohnson1986/TPD-Addin-XLAM/wiki)).
- **[Architecture](docs/ARCHITECTURE.md)** — module-by-module map of the VBA source.
- **[Changelog](CHANGELOG.md)** — what changed in each release.

## Building and contributing

The repo tracks the VBA **source** in `/src`; the compiled `.xlam` is published as a versioned asset on [Releases](https://github.com/srjohnson1986/TPD-Addin-XLAM/releases), never committed. There is no CLI compiler: you edit in the Excel VBE and a builder macro assembles the `.xlam` from `/src`.

Contributions, bug reports and ideas are welcome. See **[CONTRIBUTING.md](CONTRIBUTING.md)** for prerequisites (Windows + Excel + [Rubberduck](https://rubberduckvba.com/)), the edit → export → build workflow and the release process. Please read the [Code of Conduct](CODE_OF_CONDUCT.md), and report security problems privately as described in [SECURITY.md](SECURITY.md).

## License

The source code is released under the [MIT License](LICENSE) — Copyright (c) 2026 Steve Johnson.

### Logos and trademarks

The TPD name and logo artwork (`assets/`, and the copies embedded in the dialogs' `.frx` files and the base workbook) are the property of TPD and are **not** covered by the MIT license. See [assets/README.md](assets/README.md). If you fork this project, replace them with your own branding.

# Security Policy

This add-in is a macro-enabled Excel workbook that runs on your machine with the
same permissions as Excel, so security reports are taken seriously.

## What the add-in does (and doesn't) touch

- It reads and writes **the workbooks you point it at** (creating output sheets,
  splitting sheets, saving `.xlsx` files to a folder you choose).
- It stores **your preferences** under
  `HKCU\Software\VB and VBA Program Settings\TPD_Addin\Preferences` in the
  Windows registry, and a copy of your chosen header logo in
  `%APPDATA%\TPD_Addin\`.
- It makes **no network connections** and sends no data anywhere. The only
  `Declare` statements (all in `modGdiPlus`) call Windows' own `gdiplus` and
  `oleaut32` libraries to load the logo preview image.
- Its dialogs link to this repo's GitHub pages (issues, wiki, releases) by
  opening them in your browser.

If you find behaviour that contradicts any of the above, that is a security bug.

## Supported versions

Only the **latest release** receives fixes. Please reproduce on the latest
[release](https://github.com/srjohnson1986/TPD-Addin-XLAM/releases) before
reporting.

## Reporting a vulnerability

**Please don't open a public issue for a security problem.** Instead, use GitHub's
private reporting:

1. Go to the repo's [Security tab](https://github.com/srjohnson1986/TPD-Addin-XLAM/security).
2. Choose **Report a vulnerability** and describe the problem.

Helpful things to include: the add-in version (shown in the dialogs' bottom
corner), your Excel version/bitness, what you expected vs. what happened, and
steps to reproduce.

You can expect an acknowledgement within a week. This is a small volunteer
project, so fix timelines depend on severity and availability, but you'll be kept
informed and credited in the release notes if you'd like.

## Verifying a download

Each release attaches a single `TPD_Addin.xlam`. If you'd rather not trust a
prebuilt binary, build it yourself from the tagged source — see
[CONTRIBUTING.md](CONTRIBUTING.md#rebuilding-a-testable-xlam-from-src). The
`.xlam` is unsigned, so Excel will show its standard macro warning.

# Assets

Source brand images kept in version control so the original art isn't lost. The
build does not read these files: a UserForm's `Picture` property embeds the
bytes straight into its `.frx` when set in the VBE, and the default header logo
is a shape on the `_Resources` sheet of the base `.xlam`.

| File | What it is |
| ---- | ---------- |
| `tpdAddinLogo.png` / `.jpg` | The wide brand-bar logo shown at the top of each dialog |
| `tpdHeaderLogo.jpg` | The square-ish header logo (matches `_Resources/DefaultLogo`), used for the built-in-logo preview in the Set Defaults dialog |

## These images are not open source

The TPD name and logo are the property of TPD. They are included here only so
this project can build and run as TPD uses it. **They are excluded from the
repository's [MIT license](../LICENSE)**: no permission is granted to use,
copy, modify or redistribute them, or to use the TPD name or logo in a way that
suggests your fork is TPD's or is endorsed by TPD.

The same applies to copies of this artwork embedded elsewhere in the project:

- the dialog `.frx` files in `src/Forms/` (brand band picture), and
- the `DefaultLogo` shape on the `_Resources` sheet inside the base `.xlam`
  (which is not tracked in the repo).

## Forking or redistributing

If you fork this project for your own use, replace the TPD artwork with your own:

1. Swap the images in this folder.
2. In the VBE, set the brand-band `Picture` on each dialog to your image and
   re-export the form (`.frm` + `.frx`) to `/src/Forms`.
3. Replace the `DefaultLogo` shape on the `_Resources` sheet in your base file.

Users of the add-in can also set their own header logo at any time under
**TPD > One-click > Set Defaults > Logo**.

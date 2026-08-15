# VistaType LP artwork

The two files here are the **source of truth** for how the installer looks. Jerry supplied
them on 8/15/2026. Nothing in the build edits them.

| File | What it is | Size as supplied |
|---|---|---|
| `vistatype-icon.png` | The mark on its own — a page with a magnifier over "LP" | 481 x 515, RGBA |
| `vistatype-wordmark.png` | "VistaType" set beside the same mark | 1573 x 515, RGBA |

Both have a transparent background, and that is deliberate: the wizard images are laid onto
whatever background the installer is drawing, so they suit a light or a dark Windows theme
without a second set of files.

## What is made from them

`tools/lib/build_branding.py` (`make branding`) generates **`installer/branding/`**, which is a
build output — regenerate it, never hand-edit it:

- **`vistatype.ico`** — nine sizes from 16 to 256 pixels, from the icon. Windows picks whichever
  fits where it is drawing; an icon carrying only one size gets scaled by Windows and looks it.
  This is the installer's own icon (`SetupIconFile`) and the one Programs and Features shows
  beside the entry (`UninstallDisplayIcon`, which is why a copy is also *installed*, into
  `%AppData%\VistaType LP\`).
- **`wizard-*.png`** — four sizes of the tall panel on the welcome page: the mark above, the
  name below, as one centered group.
- **`wizard-small-*.png`** — seven sizes of the small square that sits top-right on every page
  after the welcome one. The mark alone.

Several sizes of each, so a high-DPI screen gets a sharp image: Inno Setup picks the closest
match and does not have to stretch one.

To move the two elements about on the welcome panel, change `ICON_WIDTH`, `WORDMARK_WIDTH`,
`GAP` and `GROUP_CENTRE` at the top of `build_branding.py` and run `make branding` — the
proportions are fractions of the panel, so all four sizes follow together.

## Why the installer has an icon at all

Until 8/15/2026 it had none, and no version numbers either, so the built `Setup.exe` was a
generic Inno Setup stub whose properties read **FileVersion 0.0.0.0** with the file-version and
original-filename strings blank.

That is not only cosmetic. "No version information" sits beside "unsigned" as one of the two
heavily-weighted factors when the same kind of machine-learning detection hit one of Microsoft's
own tools — and something of that kind had been deleting this installer since 8/13/2026. A file
that looks like a product is less likely to be judged as though it were not one. The whole
story is in `docs/Code-Signing.md`.

It also means a transcriber who right-clicks the download and checks Properties before running
it — which is exactly what a careful one does — now sees who made it.

## Where these travel

`installer/` is **wiped on the build box before each copy**, so `installer/branding/` goes with
it and a file deleted here really does leave the next `Setup.exe`. That is the opposite of
`assets/`, which is never wiped there — which is why the generated files live under
`installer/` and not beside the artwork.

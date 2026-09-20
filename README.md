# VistaType LP

**[vistatypelp.org](https://vistatypelp.org/)**

Microsoft Word tools for producing **large-print books** for readers with low vision, and
**braille source files** for the [Duxbury Braille Translator](https://www.duxburysystems.com/)
(DBT) via its BANA template.

VistaType LP is written for transcribers — the people who take an ordinary book and turn it
into an edition someone can actually read. It is free software, and it has been in real
production use since 2015.

- **Author:** Jerry Whittaker · jerry@vistatypelp.org
- **License:** [GNU General Public License v3.0](LICENSE) · © 2015–2026 Jerry Whittaker

---

## What it does

VistaType LP loads with Word and adds two ribbon tabs — **VistaType LP** and
**Braille Macros** — alongside your own. They hold the jobs a transcriber does over and
over, as single buttons.

**Large print.** Page size, margins and fonts set to the size a reader needs; colored
callout boxes; table of contents formatting; image resizing and recoloring; fill-in lines;
multi-column layouts. The rule the whole large-print side is built on: **every character
appears at the size the reader asked for** — nothing is shrunk to make it fit.

**Braille and DBT.** Preparing a file so Duxbury translates it correctly — Nemeth math,
UEB and EBAE, BANA bullets, and attaching the right translation template.

**Reference page numbers.** The print book's page numbers, cited in the large-print or
braille edition so a reader can follow a class or a citation. Tag them automatically or by
hand, check them, and embed or lift them back out.

**File cleanup.** Stripping the stray formatting out of a raw source document and
normalizing its styles, before any of the above is worth attempting.

**DAISY, NIMAS and text sources.** Importing and repairing text-based source formats.

---

## Installing

Transcribers install with a single `Setup.exe`. It installs for one user only and needs no
administrator rights.

It places the add-in and the large-print template, adds the two ribbon tabs (merging with
your own ribbon rather than replacing it), stocks the Quick Access Toolbar, and sets Word's
trust settings so the macros are allowed to run.

- **[Installation Guide](docs/Installation-Guide.md)** — installing, updating, uninstalling,
  and what to do when something is wrong.
- **[Download](https://vistatypelp.org/#download)** — the current release.
- **[Why antivirus sometimes eats the installer](docs/Code-Signing.md)** — and what to do
  about it.

---

## What gets installed

| Piece | Role |
|---|---|
| `LPandBRL.dotm` | The add-in itself — the macros and the two ribbon tabs. Word loads it at startup, so the tools are there for every document. |
| `LargePrintTemplate.dotx` | The large-print styles and page setup, attached to each large-print document. VistaType LP treats a document as large print when this template is attached. |
| VistaTypeLP Sans | A typeface chosen for low vision and bundled with the add-in. See **[the coverage notes](docs/VistaTypeLP-Sans.md)** for what it does and does not set. |

---

## The source

Roughly 16,500 lines of VBA, written by one person over twenty years, kept here as text
rather than locked inside the `.dotm`. Word is the only thing that can compile VBA, so
builds run on a Windows machine with Word on it, driven from Linux.

If you want to build it yourself, that is all in **[DEVELOPMENT.md](DEVELOPMENT.md)**; the
installer has its own **[notes](installer/README.md)**.

---

## License

VistaType LP is free software under the **GNU General Public License v3.0** — see
[`LICENSE`](LICENSE). You may use, study, share and modify it; if you pass on a modified
version, its source has to go with it under the same license.

© 2015–2026 Jerry Whittaker. The software is provided "as is", without warranty of any kind.

---

*Built for large-print and braille transcribers.*

# VistaType LP + Braille Macros

Microsoft Word tools for producing **large-print documents** for visually impaired
readers, and **braille source files** for the [Duxbury Braille Translator](https://www.duxburysystems.com/)
(DBT) via its BANA template.

VistaType LP ships as a Word add-in (a macro-enabled template loaded at startup) plus a
large-print styles template. It adds two ribbon tabs — **VistaType LP** and **Braille
Macros** — with one-click tools for cleaning up source files, formatting reference page
numbers, converting lists and tables, handling images, tagging DAISY/NIMAS text, and more.

- **Version:** 3.0
- **Author:** Jerry Whittaker · jerry@vistatypelp.org
- **License:** [GNU General Public License v3.0](LICENSE) · © 2015–2026 Jerry Whittaker

---

## What's in the box

The product is three Office artifacts that work together:

| Artifact | Role |
|---|---|
| `LPandBRL.dotm` (built from `src/` onto the tracked `.dotm` shell) | **The engine.** The whole VBA project — ~208 subs/functions in `LPandBrlMacros` plus 45+ UserForms — and the embedded ribbon. Loaded from Word's `STARTUP` folder. |
| `LargePrintTemplate.dotx` | **The styles.** The large-print paragraph/character styles and page setup, attached to each large-print document. |
| Embedded ribbon (`customUI14.xml`) | **The UI.** The two ribbon tabs, embedded in the `.dotm` so they *merge* with the user's ribbon instead of replacing it. |

## Feature areas

- **Large print** — custom page size / margins / fonts, colored callout boxes, TOC
  formatting, image resize & recolor, fill-in lines, multi-column layouts.
- **Braille / DBT** — prep for the Duxbury Braille Translator: Nemeth math, UEB/EBAE,
  BANA bullets, and the target `.dxt` translation template.
- **Reference page numbers** (`$pg` tags) — auto-tag, manual-tag, validate, embed /
  un-embed, and format the print-book page numbers cited in the output.
- **File cleanup / normalization** — strip stray formatting and normalize raw source docs.
- **DAISY / NIMAS / text tools** — import and fix text-based source formats.

---

## How this repo is built

The **text source of truth** lives under `src/`; the binary `.dotm` is a build output.
VBA can only be compiled by Word itself, so builds run on a **remote Windows + Word box
driven over SSH** — you edit text and run `make` from Linux and never open Word by hand.

```
edit src/vba/*.bas ──▶ make build ──▶ [Windows+Word: import + compile]
                                  └─▶ embed ribbon (Linux) ──▶ dist/LPandBRL.dotm
```

```
src/vba/        VBA modules — *.bas (standard), *.cls (class/document)
src/forms/      UserForms — *.frm + *.frx
src/ribbon/     customUI14.xml (embedded ribbon) + legacy Word.officeUI
installer/      Inno Setup installer + QAT-merge scripts
tools/          extract / build / ribbon helpers (Python + PowerShell)
docs/           end-user installation guide
LPandBRL.dotm   the .dotm shell (tracked; project references + non-VBA parts; the build base)
LargePrintTemplate.dotx   the large-print styles template
```

Run `make help` for the full target list (`build`, `pull`, `ribbon`, `qat`, `read`,
`deploy`, `installer`).

---

## Documentation

- **[Installing (for users)](docs/Installation-Guide.md)** — the one-click installer,
  updating, uninstalling, and troubleshooting.
- **[Developing](DEVELOPMENT.md)** — the Linux-first source/build workflow, the remote
  Word build, the embedded ribbon, and the QAT merge.
- **[Building the installer](installer/README.md)** — what the installer automates.
- **[CLAUDE.md](CLAUDE.md)** — architecture and conventions overview.

---

## Installing

End users install with a single per-user `Setup.exe` (no admin required) — see the
[Installation Guide](docs/Installation-Guide.md). It places the add-in and template,
embeds the ribbon (merging non-destructively with the user's own), pre-stocks the
Quick Access Toolbar icons, and configures Word's trust settings so the macros run.

---

## License

VistaType LP is free software licensed under the **GNU General Public License v3.0** —
see [`LICENSE`](LICENSE). You may use, study, share, and modify it; if you distribute a
modified version, you must also make your source available under the GPL.

© 2015–2026 Jerry Whittaker. The Software is provided "as is", without warranty of any kind.

---

*Built for large-print and braille transcribers.*

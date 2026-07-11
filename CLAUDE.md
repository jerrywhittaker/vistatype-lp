# VistaType LP

A Microsoft Word add-in that helps transcribers produce **large-print** documents for
visually impaired readers, and **braille** source files for the Duxbury Braille Translator
(DBT) via its BANA template. Authored by Jerry Whittaker (jerry@thewhittakers.org),
copyright 2015–2026.

The repo keeps the **text source of truth** under `src/` and treats the binary Office
artifacts as build outputs. All real logic lives in the VBA project embedded in
`Normal.dotm`. Because VBA can only be compiled by Word itself, builds run on a remote
Windows+Word box driven over SSH — see **`DEVELOPMENT.md`** for the full workflow.
`src/` is authoritative; **never hand-edit `Normal.dotm`.**

## The three files and how they work together

| File | What it is | Role |
|------|-----------|------|
| `Normal.dotm` | Word global macro template (macro-enabled) | **The code.** Holds the entire VBA project: ~208 subs/functions in `LPandBrlMacros`, plus 45+ UserForms. Loaded by Word at startup, so its macros are available to every document. This is where all behavior lives. |
| `Word.officeUI` | Ribbon customization XML | **The UI.** Defines the custom ribbon tabs **"VistaType LP"** (large print) and **"Braille Macros"** (DBT/BANA), plus a QAT icon group. Each button's `onAction` names a VBA sub in `Normal.dotm`. |
| `LargePrintTemplate.dotx` | Word document template | **The style set.** The template *attached to a user's large-print document* (vs. `Normal.dotm` which is the global template). Supplies paragraph/character styles and page setup. The VBA references it by name in 7+ places; `AutoOpen` treats a document as "large print" when this template is attached. |

Flow: User opens/creates a doc → `Normal.dotm`'s `AutoOpen`/`AutoNew` run → detects
whether the attached template is `LargePrintTemplate.dotx` (large print), a BANA braille
template, or neither → configures Word accordingly → the ribbon buttons from `Word.officeUI`
invoke VBA subs to clean up, format, and tag the document.

## Deployment locations (on a transcriber's Windows PC)

- `Normal.dotm` → `%AppData%\Microsoft\Templates\` (Word's global template)
- `LargePrintTemplate.dotx` → Word user templates folder; attached to each LP document
- `Word.officeUI` → `%AppData%\Microsoft\Office\` (Word reads ribbon customizations here)

## Repo layout

```
src/vba/        canonical VBA text: *.bas (std modules), *.cls (class/document modules)
src/forms/      canonical UserForms: *.frm + *.frx (binary layout)
src/ribbon/     customUI14.xml — embedded ribbon (source of truth); Word.officeUI (legacy)
Normal.dotm     shell/base + current deployable build (project references + non-VBA parts)
LargePrintTemplate.dotx   the attached large-print template (styles/page setup)
Word.officeUI   legacy global ribbon (no longer shipped; kept for reference)
tools/windows/  Export-Vba.ps1 / Import-Vba.ps1 — run in Word on the build box
tools/lib/      decompress_vba.py (reader); officeui_to_customui.py + inject_customui.py (ribbon); extract_qat.py (QAT list)
installer/      Inno Setup installer (vistatype.iss) + scripts/ (QAT merge/remove) — replaces manual file-copy install
docs/           Installation-Guide.md (end-user install); Build-VM-Setup.md (Hyper-V build/test box); Software-Agreement.md (GPLv3 About-dialog text)
reference/      generated read aids (gitignored mirror + interim form-code dump)
Makefile        pull / build / ribbon / qat / read / deploy / stage / installer  (see DEVELOPMENT.md)
```

The ribbon is **embedded** in `LPandBRL.dotm` (`src/ribbon/customUI14.xml`), so it merges
with each user's ribbon instead of overwriting it — the old `Word.officeUI` file is no
longer shipped. Every ribbon button routes through one VBA dispatcher, `RibbonAction`
(`src/vba/RibbonCallbacks.bas`), which runs the macro named in the control's `tag`. The
QAT can't be set from a template, so the **installer** merges VistaType's 6 quick-access
icons into each user's own `Word.officeUI` non-destructively (`installer/scripts/`,
list in `installer/qat-controls.xml`).

## Build & edit workflow (short version)

Full detail in **`DEVELOPMENT.md`**. In brief: edit text under `src/`, then
`make build` ships it to the remote Windows+Word box over SSH, which imports the
source into `dist/Normal.dotm` (Word regenerates p-code) and copies it back; smoke-test
in Word, then `make deploy`. Never edit `Normal.dotm` by hand. `make pull` refreshes
`src/` from the `.dotm` (canonical export, needed to (re)seed valid `.frx`); `make read`
dumps readable source with the Linux decompressor without needing Windows.

**Keep this file in sync.** When you change the build pipeline or architecture —
`Makefile`, `DEVELOPMENT.md`, anything under `tools/`, `src/ribbon/`, or `installer/`,
or the build config — update this file's **Repo layout** and **Build & edit workflow**
sections in the *same* change. A project `PostToolUse` hook
(`.claude/hooks/claude-md-sync.py`, wired in `.claude/settings.json`) reminds Claude Code
to do this automatically after editing any of those files.

Why not compile on Linux: `vbaProject.bin` stores compiled **p-code** alongside source,
and only Word's VBA engine can regenerate it correctly — so the compile step always
runs in Word (this drove the remote-build design; see DEVELOPMENT.md).

### Notable modules

- **`LPandBrlMacros`** — the engine (~16,500 lines). Contains `AutoOpen`, `AutoNew`,
  `AutoClose`, all ribbon handlers, and orchestrators like `Lp_File_Fix_Sequence`,
  `Dx_File_Fix_Sequence`, `Lp_Get_Doc_Setup_Params`, `Lp_Is_The_Attached_Template_LP`,
  and `MS_Set_Word_Config_For_New_Install`.
- **`LpExportImportSelectedText` / `DxExportImportSelectedText`** — round-tripping selected
  text to/from separate files.
- **`ShNonModalMessage`** — shared non-modal status messaging.
- **~45 UserForms** — dialogs, prefixed by domain (see below).

### Naming convention (prefixes)

Everything is namespaced by a short prefix — grep by it to find a feature area:

- `Lp_` — Large Print features (~98 subs) — the "VistaType LP" ribbon tab.
- `Dx_` — Duxbury/braille features (68 subs) — the "Braille Macros" ribbon tab (DBT + BANA).
- `Sh_` — Shared helpers used by both (28 subs), e.g. `Sh_Doc_Info`, title-case, keep-together.
- `DN_` — DAISY / NIMAS / text-file tools.
- `MS_` — Microsoft Word configuration/normalization (`MS_Set_Word_Config_For_New_Install`).

The `Word.officeUI` `onAction` values map 1:1 to these subs (e.g. button
`Lp_File_Fix_Sequence` → `Sub Lp_File_Fix_Sequence`).

### Editing convention

Every sub is versioned inline via a comment block (Version/Date/Author). The module header
of `LPandBrlMacros` keeps a running dated changelog. Current version string in the About
forms: **VistaType LP (150)** / v2.2.3. When changing behavior, follow the existing pattern:
bump the per-sub version comment and add a dated line to the header changelog.

## Domain concepts

- **Large print**: reformatting for low-vision readers — big base fonts, custom page
  size/margins/gutter/orientation (public vars `PPH/PPW/PTM/PBM/PLM/PRM/PPG/PPO/DM`),
  colored "Box" styles, colored/format TOC, image resize & recolor, fill-in lines.
- **Braille / DBT**: prepping Word docs for the Duxbury Braille Translator. The target DBT
  translation template (e.g. `English (UEB) - BANA with Nemeth.dxt`) is stored in the
  document's `docProps/custom.xml`. Includes Nemeth math, UEB/EBAE, BANA bullet creation.
- **Reference page numbers ("$pg" tags)**: print-book page numbers embedded/tagged in the
  text so braille/LP output can cite the original pagination. Auto-tag, manual-tag, validate,
  embed/un-embed, and format tools exist for both LP and Dx.
- **File cleanup / normalization**: `*_File_Fix_Sequence` and selection cleanup subs strip
  stray formatting, fix paragraph marks, and normalize a raw source doc.

## Environment / dependencies

- Windows + Microsoft Word (built against Office 16 / VBA7.1). No cross-platform support.
- VBA references: MSO.DLL, VBE7, MSWORD.OLB, FM20.DLL (MSForms), `scrrun.dll` (Scripting
  Runtime), stdole2. Assumes Duxbury DBT installed for the braille side.
- **No build system, tests, or package manager** — these are hand-authored Office files
  edited in the Word/VBA IDE on Windows. There is nothing to compile here.

## Practical notes for changes

- To inspect or diff logic, extract `vbaProject.bin` as above; do **not** try to read
  `Normal.dotm` directly as text.
- Re-packaging edited VBA back into a `.dotm` cannot be done reliably by hand on this
  (Linux) box — real edits are made in Word's VBA editor on Windows and the `.dotm` re-saved.
- `Word.officeUI` and `LargePrintTemplate.dotx` are plain zip/XML and can be edited as XML,
  but keep `onAction` names in sync with the VBA subs, and keep style IDs stable (documents
  in the field reference them).

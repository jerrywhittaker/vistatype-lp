# VistaType LP

A Microsoft Word add-in that helps transcribers produce **large-print** documents for
visually impaired readers, and **braille** source files for the Duxbury Braille Translator
(DBT) via its BANA template. Authored by Jerry Whittaker (jerry@thewhittakers.org),
copyright 2015–2026.

The repo keeps the **text source of truth** under `src/` and treats the binary Office
artifacts as build outputs. The real logic is authored as text under `src/vba` and
`src/forms`; `make build` compiles it (in Word) into the shipping add-in `LPandBRL.dotm`.
Because VBA can only be compiled by Word itself, builds run on a remote Windows+Word box
driven over SSH — see **`DEVELOPMENT.md`** for the full workflow. `src/` is authoritative;
**never hand-edit `LPandBRL.dotm`** — it is a build output.

## The pieces and how they work together

The source of truth is `src/`. The items below are the *built* artifacts (and their
sources) that ship or that Word loads.

| Piece | What it is | Role |
|------|-----------|------|
| `LPandBRL.dotm` | Word add-in template (macro-enabled), built from `src/` | **The code.** The entire compiled VBA project — ~208 subs/functions in `LPandBrlMacros`, plus 45+ UserForms — and the embedded ribbon (below). Loaded from Word's `STARTUP` folder, so its macros are available to every document. Behavior is *authored* under `src/vba`/`src/forms`; this is where it *runs*. |
| Embedded ribbon (`src/ribbon/customUI14.xml`) | Ribbon customization XML, embedded into `LPandBRL.dotm` at build time | **The UI.** Defines the custom ribbon tabs **"VistaType LP"** (large print) and **"Braille Macros"** (DBT/BANA). Every button's `tag` names a VBA sub, dispatched through one `RibbonAction` handler. Because it is *embedded* (not the old global `Word.officeUI`), it **merges** with the user's ribbon instead of replacing it. |
| `LargePrintTemplate.dotx` | Word document template | **The style set.** The template *attached to a user's large-print document* (vs. `LPandBRL.dotm`, the global add-in loaded for every document). Supplies paragraph/character styles and page setup. The VBA references it by name in 7+ places, and treats a document as "large print" when this template is attached. |

Flow: User opens/creates a doc → `LPandBRL.dotm`'s Word **application events** fire (the
`VtEvents` sink, hooked by `AutoExec`) → detect whether the attached template is
`LargePrintTemplate.dotx` (large print), a BANA braille template, or neither → configure
Word accordingly → the embedded-ribbon buttons invoke VBA subs (via `RibbonAction`) to
clean up, format, and tag the document.

## Deployment locations (on a transcriber's Windows PC)

Installed by the Inno Setup installer (`installer/vistatype.iss`):

- `LPandBRL.dotm` → `%AppData%\Microsoft\Word\STARTUP\` (Word auto-loads it as a global add-in)
- `LargePrintTemplate.dotx` → `%AppData%\Microsoft\Templates\`; attached to each LP document
- VistaType's **standard QAT toolbar** (`installer/qat-template.officeUI`) is installed into the user's own `Word.officeUI` (both Roaming and Local), preserving their ribbon and appending their own QAT icons; the embedded ribbon supplies the tabs

## Repo layout

```
src/vba/        canonical VBA text: *.bas (std modules), *.cls (class/document modules)
src/forms/      canonical UserForms: *.frm + *.frx (binary layout)
src/ribbon/     customUI14.xml — embedded ribbon (source of truth); Word.officeUI (legacy)
LPandBRL.dotm   the .dotm shell/base (tracked): project references + non-VBA parts; build base
LargePrintTemplate.dotx   the attached large-print template (styles/page setup)
Word.officeUI   legacy global ribbon (no longer shipped; kept for reference)
tools/windows/  Export-Vba.ps1 / Import-Vba.ps1 — run in Word on the build box
tools/lib/      decompress_vba.py (reader); officeui_to_customui.py + inject_customui.py (ribbon); extract_qat.py (obsolete/reference — QAT is now qat-template.officeUI)
installer/      Inno Setup installer (vistatype.iss) + scripts/ (QAT merge/remove) + qat-template.officeUI (the standard QAT) — replaces manual file-copy install
docs/           Installation-Guide.md (end-user install); Build-VM-Setup.md (Hyper-V build/test box); Software-Agreement.md (GPLv3 About-dialog text)
reference/      generated read aids (gitignored mirror + interim form-code dump)
Makefile        pull / build / ribbon / qat / read / deploy / stage / installer  (see DEVELOPMENT.md)
```

The ribbon is **embedded** in `LPandBRL.dotm` (`src/ribbon/customUI14.xml`), so it merges
with each user's ribbon instead of overwriting it — the old `Word.officeUI` file is no
longer shipped. Every ribbon button routes through one VBA dispatcher, `RibbonAction`
(`src/vba/RibbonCallbacks.bas`), which runs the macro named in the control's `tag`. The
QAT can't be set from a template's customUI, so the **installer** imposes VistaType's standard
QAT — the full hand-maintained toolbar in `installer/qat-template.officeUI` — via
`installer/scripts/Merge-Qat.ps1` (uninstall reverses it with `Remove-Qat.ps1`). It is
non-destructive: only the QAT `sharedControls` are replaced (the user's ribbon customizations are
kept), any QAT icons the user added themselves are re-appended to the **right** of the VistaType
block (deduped), and the user's original `Word.officeUI` is backed up (`.vtqatbak`) so uninstall
restores it (or deletes the file if we created it). Three subtleties that make it actually work:
- **Location:** Word reads `Word.officeUI` from `%APPDATA%` (Roaming) on most machines but from
  `%LOCALAPPDATA%` (Local) when the profile roams/redirects or Office can't roam, so the merge
  writes **both** (Word honors whichever it uses; the other is ignored).
- **Format:** each VistaType QAT entry is a **reference to the add-in's own ribbon control** —
  `<mso:control idQ="x1:btn_<macro>">`, where the `x1` namespace is the installed
  `LPandBRL.dotm`'s full path — the exact shape Word itself writes when a user adds one of our
  ribbon buttons to the QAT by hand. (Standalone `onAction` macro buttons did **not** display.)
- **Template:** `qat-template.officeUI` is hand-edited (`__VT_DOTM_PATH__` is substituted with the
  install path at merge time); its `btn_<macro>` ids must match `src/ribbon/customUI14.xml`.

## Build & edit workflow (short version)

Full detail in **`DEVELOPMENT.md`**. In brief: edit text under `src/`, then
`make build` ships it to the remote Windows+Word box over SSH, which imports the
source into `dist/LPandBRL.dotm` (Word regenerates p-code) and copies it back; smoke-test
in Word, then `make deploy`. Never edit `LPandBRL.dotm` by hand. `make pull` refreshes
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

- **`LPandBrlMacros`** — the engine (~16,500 lines). Contains `AutoExec`, the document-event
  handlers (`Sh_HandleDocumentOpened`/`New`/`Closing`), all ribbon handlers, and orchestrators
  like `Lp_File_Fix_Sequence`, `Dx_File_Fix_Sequence`, `Lp_Get_Doc_Setup_Params`,
  `Lp_Is_The_Attached_Template_LP`, and `MS_Set_Word_Config_For_New_Install`.
- **`VtEvents`** (class) — Word application-event sink (`WithEvents Application`). Hooked once
  by `AutoExec`, it drives document-type detection via `DocumentOpen`/`NewDocument`/
  `DocumentBeforeClose` so detection fires even when the add-in is loaded from the STARTUP
  folder (where `AutoOpen` would not). The `AutoOpen`/`AutoNew`/`AutoClose` macros remain as
  thin back-compat stubs used only if the add-in is instead loaded as `Normal.dotm`.
- **`LpExportImportSelectedText` / `DxExportImportSelectedText`** — round-tripping selected
  text to/from separate files.
- **`ShNonModalMessage`** — shared non-modal status messaging.
- **~45 UserForms** — dialogs, prefixed by domain (see below).

The built add-in's **VBA project is named `LPandBRL`** (not `Normal`): it ships in Word's
STARTUP folder loaded alongside the user's own `Normal.dotm`, and two loaded projects can't
share the name `Normal`. `Import-Vba.ps1` sets this name at build time (`-ProjectName`,
`PROJNAME` in the Makefile); nothing in the code references the project name.

### Naming convention (prefixes)

Everything is namespaced by a short prefix — grep by it to find a feature area:

- `Lp_` — Large Print features (~98 subs) — the "VistaType LP" ribbon tab.
- `Dx_` — Duxbury/braille features (68 subs) — the "Braille Macros" ribbon tab (DBT + BANA).
- `Sh_` — Shared helpers used by both (28 subs), e.g. `Sh_Doc_Info`, title-case, keep-together.
- `DN_` — DAISY / NIMAS / text-file tools.
- `MS_` — Microsoft Word configuration/normalization (`MS_Set_Word_Config_For_New_Install`).

Each embedded-ribbon button's `tag` names one of these subs (e.g. `tag="Lp_File_Fix_Sequence"`
→ `Sub Lp_File_Fix_Sequence`), run by the shared `RibbonAction` dispatcher.

### Editing convention

Every sub is versioned inline via a comment block (Version/Date/Author). The module header
of `LPandBrlMacros` keeps a running dated changelog. Current version: the shipped
package/installer is **3.0.2** (`APPVER` in the Makefile, `AppVer` in
`installer/vistatype.iss`, which drives the `VistaType-LP-Setup-<ver>.exe` name); the About
form's `VersionLabel` caption reads **v3.0.2**. (The `VistaType LP (NNN)` numbers in MsgBox
titles are per-dialog IDs, *not* version numbers.) When changing behavior, follow the existing
pattern: bump the per-sub version comment and add a dated line to the header changelog.

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
  `LPandBRL.dotm` directly as text.
- Re-packaging edited VBA back into a `.dotm` cannot be done reliably by hand on this
  (Linux) box — real edits are made in Word's VBA editor on Windows and the `.dotm` re-saved.
- `src/ribbon/customUI14.xml` and `LargePrintTemplate.dotx` are plain XML / zip+XML and can
  be edited directly, but keep control `tag` names in sync with the VBA subs, and keep style
  IDs stable (documents in the field reference them).

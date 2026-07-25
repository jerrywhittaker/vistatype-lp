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

## How to talk to Jerry

**Plain English. Always.** This governs every reply, not just the git parts.

Jerry is a capable programmer — he wrote this entire ~16,500-line system himself over
roughly twenty years. **Speak Word and VBA to him freely and without translation:** subs,
modules, UserForms, `.frx`, ranges, styles, `ScreenUpdating`, attached templates, AutoCorrect
entries, wildcards, section breaks, DBT translation tables. That is his working vocabulary
and he knows it better than you do.

What he does *not* use is the vocabulary of professional software engineering — version
control, build pipelines, and release process. He never needed it: he was one person editing
a `.dotm` in the VBA editor. So **the jargon to strip is process jargon, not programming
jargon.**

**Do not use these without plainly saying what they mean** (or better, avoid them entirely):
branch, merge, fast-forward, rebase, cherry-pick, HEAD, upstream, remote, origin, checkout,
staging, working tree, diff, CI/CD, pipeline, artifact, idempotent, atomic, regression,
refactor, blocking, upstream/downstream.

Say the **effect**, not the mechanism:

| Instead of | Say |
|---|---|
| "I'll cherry-pick this onto `master` and rebase `dev`." | "I'll put this on the released version, then copy it into your working copy." |
| "`master` is an ancestor of `dev`, so it'll fast-forward." | "Your working copy already contains everything that's released, so releasing is just moving a marker." |
| "This commit is a no-op." | "This one changes nothing — it's already in there." |
| "Resolve the conflict in the `.frx`." | "Two edits touched the same dialog's layout; I need you to look at it." |
| "I'll stage and commit the diff." | "I'll save these changes into the history." |

Other rules:
- **Lead with what happened or what he should do**, then the reason. Not the reverse.
- When a term is unavoidable because he'll *see* it — in git's own output, or on a GitHub
  page — name it once, explain it in a clause, and point him at
  `docs/Daily-Workflow-and-Releases.md`, which has a glossary.
- **Never say "just"** ("just rebase onto…"). It makes unfamiliar things sound obvious.
- **Don't over-explain programming itself.** Simplify the *process* vocabulary; do not
  simplify the code, the Word behavior, or the reasoning. Explaining what a loop is would be
  as wrong as saying "rebase" unexplained. He is experienced, not a beginner.
- Short paragraphs. One idea each. He is reading this in a terminal between Word sessions.

## The pieces and how they work together

The source of truth is `src/`. The items below are the *built* artifacts (and their
sources) that ship or that Word loads.

| Piece | What it is | Role |
|------|-----------|------|
| `LPandBRL.dotm` | Word add-in template (macro-enabled), built from `src/` | **The code.** The entire compiled VBA project — ~208 subs/functions in `LPandBrlMacros`, plus 43 UserForms — and the embedded ribbon (below). Loaded from Word's `STARTUP` folder, so its macros are available to every document. Behavior is *authored* under `src/vba`/`src/forms`; this is where it *runs*. |
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
docs/           Daily-Workflow-and-Releases.md (Jerry's plain-language guide to dev/master, building, releasing, and what to ask Claude); Installation-Guide.md (end-user install); Build-VM-Setup.md (Hyper-V build/test box); Software-Agreement.md (GPLv3 About-dialog text)
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

Full detail in **`DEVELOPMENT.md`**. In brief: **on the `dev` branch** (never `master` — see
*Git workflow and releases* below) edit text under `src/`, then `make build` ships it to the
remote Windows+Word box over SSH, which imports the source into `dist/LPandBRL.dotm` (Word
regenerates p-code) and copies it back; smoke-test in Word, then `make deploy`. Never edit
`LPandBRL.dotm` by hand. `make pull` refreshes `src/` from the `.dotm` (canonical export,
needed to (re)seed valid `.frx`); `make read` dumps readable source with the Linux
decompressor without needing Windows. Releases go out from `master` via `make installer`.

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
- **43 UserForms** — dialogs, prefixed by domain (see below).

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
of `LPandBrlMacros` keeps a running dated changelog. The version number lives in **four**
places that must always agree — `APPVER` (Makefile), `AppVer` (`installer/vistatype.iss`,
which drives the `VistaType-LP-Setup-<ver>.exe` name), the LP **and** Braille About dialogs'
`VersionLabel` caption (stored in the binary `.frx`), and this file. **What actually shipped
is the newest `v*` tag on `master`** — not whatever these files say. Those four normally read
a working `3.0.X` number that has never been released; see *Version numbering* below. (The
`VistaType LP (NNN)` numbers in MsgBox titles
are per-dialog IDs, *not* version numbers.) When changing behavior, follow the existing
pattern: bump the per-sub version comment and add a dated line to the header changelog.

## Git workflow and releases

This add-in ships to working transcribers, so a release must stay **stable and revertable**.
Two branches, fast-forward only, one tag per release.

> Jerry's plain-language version of this section is **`docs/Daily-Workflow-and-Releases.md`**.
> Keep the two in step, and point him there when he asks how any of this works.

- **`master` = the last released version.** Always shippable. **Never commit to it directly**
  and never run `make deploy` while sitting on it.
- **`dev` = all day-to-day work.** Small, focused commits, exactly as before.
- **A release is `dev` fast-forwarded into `master`, then tagged.**

Work only ever moves `dev` → `master`, never the reverse. Hold to that and the fast-forward
always succeeds. **One sanctioned exception: documentation** — see immediately below.

### Version numbering — the third number is internal, not a release

Jerry's rule, set 2026-07-25: **the third number is a private build counter, not a shipping
version.** Do not confuse a version bump with a release.

- **Third number — `3.0.6`, `3.0.7`, `3.0.8` …** — day-to-day builds on `dev`. Bump freely,
  build the installer, have Jerry install and test it, commit. **Nothing is tagged, nothing
  is pushed to `master`, nothing is published.** These numbers exist so Jerry can tell one
  test build from the next in the About box. After `3.1` ships, the counter carries on as
  `3.1.1`, `3.1.2`, … — still internal.
- **`X.Y` — `3.1`, `3.2`, `3.3` …** — actual releases. A release takes the next `X.Y`, never
  a third number. Several `3.0.X` builds' worth of work rolls up into one `3.1`.
- **Fourth number — `3.1.0.1`, `3.1.0.2` …** — a **hotfix** to a released `X.Y`. Read it as
  "3.1, repair 1". Note the `.0`: released `3.1` is `3.1.0`, so hotfixes hang off that and
  **can never collide** with the internal `3.1.1`, `3.1.2` builds on `dev`. That separation
  is the whole reason for the fourth number.

**Only Jerry starts a release, and he starts it by name** — "let's release 3.1". Until he
says that, `master` does not move, no `v*` tag is created, and nothing is published, no
matter how many builds have accumulated on `dev`. Do not propose a release just because a
build tested clean; a clean build is the *normal* end of a day's work here.

When he does say it: renumber all four places from the working `3.0.X` to the release
`X.Y`, rebuild (the number is compiled into the About dialogs' `.frx`), let him test *that*
installer, and only then run the release checklist below.

The two lanes never interfere, so — unlike the old scheme — a hotfix never forces `dev`'s
working number to move.

### Documentation-only changes go straight to `master`

Jerry reads the guide from a bookmark that points at `master`
(`github.com/jerrywhittaker/vistatype-lp/blob/master/docs/Daily-Workflow-and-Releases.md`).
Holding a doc fix until the next release means he reads stale instructions in the meantime,
which is backwards — documentation ships no code and cannot break a release.

**Qualifies only if the change touches nothing outside these paths:**
`docs/**`, `CLAUDE.md`, `DEVELOPMENT.md`, `README*`.
One file under `src/`, the `Makefile`, `installer/`, or any `.dotm`/`.dotx`/`.frx` disqualifies
it — that is a code change and goes through `dev` like everything else.

**Verify before committing, every time** — do not eyeball it:

```bash
git diff --name-only HEAD          # and/or --cached; must list only the paths above
```

**Procedure — write it on `master`, then carry it into `dev`:**

```bash
git checkout master
# make the edits, or cherry-pick them if they were already written on dev
git commit ...
git push origin master             # push freely: docs only. Code still waits for "push".
git checkout dev && git merge master
```

That last line is not optional. It keeps `master` an ancestor of `dev` so the next release
still fast-forwards. **Authoring on `master` and merging into `dev` is strictly better than
committing on `dev` and cherry-picking to `master`** — cherry-picking leaves the two branches
permanently diverged and the next release fails.

The "never commit on `master`" rule below still holds for everything else.

### Why fast-forward only (do not "just merge")

`LPandBRL.dotm`, `LargePrintTemplate.dotx`, and every form `.frx` are tracked **binaries**.
Git cannot merge them — a real merge conflicts, and a wrongly-resolved `.frx` silently
corrupts a dialog's layout. `--ff-only` never runs the merge algorithm, so binaries can never
conflict, and history stays linear. If a fast-forward is ever *refused*, something committed
onto `master` directly — stop and ask Jerry rather than forcing a merge.

### Cutting a release (the checklist — walk Jerry through it, one step at a time)

0. **Jerry has said "let's release 3.1"** (or the equivalent). Without that, stop — there is
   no release. See *Version numbering* above.
1. **Confirm `dev` is ready** — `git status` clean, the change-set smoke-tested in Word.
2. **Renumber from the working `3.0.X` to the release `X.Y` in all four places** above (the
   `.frx` captions are edited headless over SSH via the VBA object model, *not* the form
   designer — see `DEVELOPMENT.md`).
3. **`make installer`** — rebuilds the `.dotm` from `src/` so the new `.frx`/ribbon/VBA compile
   in, compiles `Setup.exe`, copies it to `dist/` and the VM Desktop. Verify the built `.dotm`'s
   About caption reads the new version.
4. **Jerry installs and tests from the Setup.exe.** **Word must be fully closed first** — Word
   locks the STARTUP `.dotm` and the install silently no-ops otherwise (symptom: About still
   shows the old version).
5. **Commit the bump and the rebuilt `.dotm` on `dev`.**
6. **Confirm the installer exists before going near git** — `ls -l dist/VistaType-LP-Setup-<ver>.exe`.
   No `.exe`, no release. Go back to step 3.
7. **Only when Jerry says the build is good**, move `master` and tag:

   ```bash
   git checkout master
   git merge --ff-only dev
   git tag -a v3.1 -m "VistaType LP 3.1"
   git checkout dev            # go straight back to dev; never linger on master
   ```

8. **Publish — only when Jerry explicitly says "push"** (see below). The push and the
   `.exe`-bearing release are **one step**; never do the first without immediately doing the
   second:

   ```bash
   git push origin master dev --follow-tags
   gh release create v3.1 dist/VistaType-LP-Setup-3.1.exe \
       --title "VistaType LP 3.1" --notes "<what changed, in transcriber-facing terms>"
   gh release view v3.1 --json assets      # VERIFY: must list the .exe, not []
   ```

9. Delete the superseded `VistaType-LP-Setup-*.exe` from `dist/` and the VM Desktop (keep old
   real releases).

### The `.exe` IS the release — non-negotiable

**A GitHub release without the `VistaType-LP-Setup-<ver>.exe` attached is not a release.**
GitHub auto-attaches a "Source code (zip)" to every release; that is a tarball of VBA text
files and **cannot be installed by a transcriber**. A release with no asset therefore looks
official and delivers nothing.

This also matters because the build is **not byte-reproducible** — Word regenerates p-code on
every compile, so checking out an old tag and rebuilding does *not* give back the binary that
shipped. The uploaded `.exe` is the only true copy of a release, and the only real rollback.

Hard rules:
- Never run `gh release create` without an `.exe` path as an argument in the **same** command.
- Never push a `v*` tag unless the release-with-asset follows immediately in the same sitting.
  A pushed tag already renders as a downloadable "release" in GitHub's UI.
- Always verify with `gh release view <tag> --json assets` and confirm it is not `[]`.
- **Known gap:** the existing **v3.0** release has no asset (`assets: []`) and the 3.0
  installer is not in `dist/`; it cannot be regenerated byte-identically. If Jerry still has
  `VistaType-LP-Setup-3.0.exe` archived anywhere, upload it with
  `gh release upload v3.0 <path>`. `v2.2.3` has an installer in `dist/` but no GitHub release.

### Recovering from a bad release

- **For a transcriber, right now:** have them reinstall the previous `Setup.exe` from its
  GitHub release. Fastest fix; no rebuild involved.
- **For the source:** `git checkout v3.1 && make build`.

### Hotfix — an emergency fix for the *released* version

Use when someone in the field needs a fix now and `dev` holds half-finished work that must
not ship. Jerry's plain-language version is in `docs/Daily-Workflow-and-Releases.md`.

**Part 1 — ship the fix from the released line.** Branch from the **tag**, never from `dev`:

A hotfix takes the **fourth** number off the released one: `3.1` → `3.1.0.1` → `3.1.0.2`.
Never a third number — that lane belongs to `dev`'s internal builds.

```bash
git checkout -b hotfix/3.1.0.1 v3.1     # the TAG — an exact copy of what shipped
# fix, bump the version, make installer, Jerry installs and tests
git checkout master && git merge --ff-only hotfix/3.1.0.1
git tag -a v3.1.0.1 -m "VistaType LP 3.1.0.1"
git push origin master --follow-tags    # only when Jerry says "push"
gh release create v3.1.0.1 dist/VistaType-LP-Setup-3.1.0.1.exe --title "VistaType LP 3.1.0.1"
git branch -d hotfix/3.1.0.1            # merged; the tag is the permanent record
```

**Part 2 — fold the fix back into `dev`. Do not skip this.** The fix exists only on the
released line; without this step the next release from `dev` silently reintroduces the bug.

```bash
git checkout dev && git merge master   # dev gains the fix; master stays an ancestor of dev
```

**Merge, not rebase, here.** Rebasing `dev` rewrites commits Jerry may already have pushed as
his backup, which would demand a force-push — forbidden below. The merge commit on `dev` is
cosmetic and costs nothing: `master` remains an ancestor of `dev`, so the next release still
fast-forwards. (Rebase is fine *only* if `dev` has never been pushed.)

Two things that bite:
- **Version numbers do *not* collide** under this scheme — the hotfix uses the fourth slot
  (`3.1.0.1`) and `dev`'s working builds use the third (`3.1.4`). Leave `dev`'s number alone
  after the merge; it still becomes the next `X.Y` when Jerry calls the release. Only the
  four version locations *on the hotfix branch* get touched.
- **Binary conflicts.** `LPandBRL.dotm` is a build output — never hand-resolve it; take either
  side and `make build` to regenerate. A conflicting `.frx` is *source* (form layout) and does
  need real attention — surface it to Jerry rather than guessing.

**Claude's standing duty:** after a hotfix is released, verify `dev` contains it
(`git branch --contains <hotfix-commit>` or `git merge-base --is-ancestor master dev`). If it
does not, tell Jerry before starting other work — an unfolded hotfix is a bug that comes back.

### Rules for Claude

- **Never start a release Jerry did not ask for.** Bumping `3.0.6` → `3.0.7` and building an
  installer is ordinary work; moving `master`, tagging, and publishing happen *only* after he
  says "let's release 3.1". A build that tests clean is not a cue to release it.
- **Never push code** — not `master`, not `dev`, not tags — until Jerry says the word "push".
  He reviews and installs the Setup.exe first. Honor this every time. *Documentation-only
  changes are exempt* (see above): push those freely so his bookmark stays current.
- **Never commit code on `master`.** Check `git branch --show-current` before committing; if
  it says `master`, switch to `dev` first. The lone exception is a documentation-only change,
  verified with `git diff --name-only`.
- **Never force-push, never rewrite a published tag**, and never delete a `v*` tag without
  being asked — the tags are the recovery points.
- **Never publish a release without its `.exe`** — see the non-negotiable section above.
- This workflow is new to Jerry. **Guide him through the release steps explicitly** — say which
  command comes next and what it will do, run the git steps for him, and confirm each stage
  landed before moving on. Don't assume he knows the branch he is on; tell him.
- **Say it in plain English** — see *How to talk to Jerry* at the top of this file. Process
  jargon is the thing to strip; Word and VBA terms are fine.

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
- **On this Linux box:** `make` + `python3` (the `tools/lib` helpers), `ssh`/`scp` to the
  Windows build box, and the **`gh` CLI** — required to publish a release with its installer
  attached (authenticated as `jerrywhittaker`).
- **On the Windows build box:** Word (compiles the VBA — the one step that cannot run here)
  and Inno Setup 6 (`ISCC.exe`) for `make installer`.
- **No automated tests.** Verification is a manual Word smoke-test plus an install-and-run of
  the built `Setup.exe`.

## Practical notes for changes

- To inspect or diff logic, extract `vbaProject.bin` as above; do **not** try to read
  `LPandBRL.dotm` directly as text.
- Re-packaging edited VBA back into a `.dotm` cannot be done reliably by hand on this
  (Linux) box — real edits are made in Word's VBA editor on Windows and the `.dotm` re-saved.
- `src/ribbon/customUI14.xml` and `LargePrintTemplate.dotx` are plain XML / zip+XML and can
  be edited directly, but keep control `tag` names in sync with the VBA subs, and keep style
  IDs stable (documents in the field reference them).

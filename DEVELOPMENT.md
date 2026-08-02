# Developing VistaType LP

The whole workflow runs from this Linux terminal. Word can't compile VBA on Linux,
so a remote Windows+Word box acts as an invisible build server driven over SSH. You
edit text, run `make`, and get a finished `.dotm` back — you never open Word by hand.

## Source of truth

`src/` is authoritative. **Never hand-edit `LPandBRL.dotm`** — it is a build output.

```
src/vba/        *.bas standard modules, *.cls class/document modules   (canonical text)
src/forms/      *.frm UserForms + *.frx binary layout                  (canonical)
src/ribbon/     customUI14.xml — embedded ribbon (source of truth)
                Word.officeUI  — legacy global ribbon (kept for reference)
LPandBRL.dotm   the shell/base .dotm: project references + non-VBA parts (tracked)
LargePrintTemplate.dotx   the attached large-print template (styles/page setup; tracked)
dist/           build outputs (gitignored)
reference/      read-only aids (gitignored mirror + interim form-code dump)
```

## Embedded ribbon

The ribbon is defined **inside `LPandBRL.dotm`** as an embedded `customUI14.xml` part.
It is *deployed* in one of two ways, and the difference is the whole reason for the
machinery below.

**Two homes for the tabs (3.0.34 onwards).** The embedded definition is always the source
of truth for what the buttons *are*. Where the visible tabs *live* depends on the user:

| | Embedded (the fallback) | In the user's `Word.officeUI` (the default) |
|---|---|---|
| Set up by | nothing — it ships in the `.dotm` | the installer, `Merge-Qat.ps1 -Tabs Install` |
| In *Customize the Ribbon*? | **no** | **yes** |
| Hide / reorder / rename? | **no** | **yes** |
| Travels with the add-in? | yes | **no** — see *orphan tabs* below |
| When it applies | hand-copied install, or the option declined | ticked at install (the default) |

Why the second exists: Word does not list add-in customUI tabs in *File → Options →
Customize the Ribbon*, so tabs defined in the `.dotm` cannot be hidden, reordered or
renamed. Jerry, 7/29/2026: *"Even I want to hide the braille macro tab when I'm producing
large print, and it is not unusual that I will rearrange the tabs. Not being able to do
these things is not acceptable."* Tabs written into the user's own file are ordinary custom
tabs, so Word lists them and all three become possible.

`getVisible="VtTabVisible"` on the two visible embedded tabs makes them go dark when
`HKCU\Software\VistaType LP\UserRibbonTabs` is `1`, so the two never appear at once. Absent
(hand-copied install, or declined) means the embedded tabs show exactly as they always did —
**nobody can end up with no tabs at all.** `tab_LP_and_BRL_QAT_Icons` keeps a hard
`visible="false"` and deliberately gets **no** callback: a bug in the callback must not be
able to expose a tab of duplicated stock Word buttons.

**Orphan tabs — the price, accepted knowingly.** Tabs in the user's file do not travel with
the add-in. A proper uninstall removes them, but if Word disables the add-in after a crash,
or someone deletes the `.dotm` by hand, the tabs stay and their buttons stop working.
`docs/Installation-Guide.md` carries the troubleshooting entry. This is not fixable while
the tabs live where Word will let the user edit them; it is the trade.

- Source of truth: `src/ribbon/customUI14.xml`. Edit it directly (plain XML).
- **Never rename or remove a `btn_*` id.** Every ribbon tab and toolbar already installed in
  the field references buttons by id; a renamed one renders blank on the user's machine with
  no error, and their file is not ours to migrate. `installer/ribbon-button-ids.txt` is an
  append-only record of every id that has shipped, and `make build` stops if one disappears.
- `tools/lib/build_ribbon_tabs.py` generates `installer/ribbon-tabs.officeUI` (run by
  `make build`). Each button is emitted as a **reference** —
  `<mso:control idQ="x1:btn_...">` — not as a legacy `<mso:button onAction="...">`. Both
  render and fire (tested 7/29/2026), but the reference form keeps clicks going through
  `RibbonAction`, and Word takes the label and icon from the add-in so neither is duplicated.
  Emitting `onAction` directly would skip the `Sh_Pos_Depth` reset for 43 of the 50 buttons.
- `make build` injects it into the built `.dotm` with `tools/lib/inject_customui.py`
  (pure Linux zip/XML; runs after Word has compiled the code).
- Every button calls one dispatcher, `RibbonAction` (in `src/vba/RibbonCallbacks.bas`),
  which runs the macro named in the control's `tag`. That's why the existing
  parameterless entry-point Subs needed **no** changes.
- `make ribbon` regenerates `customUI14.xml` from the legacy `Word.officeUI` via
  `tools/lib/officeui_to_customui.py` — a one-off migration aid; normally you just
  hand-edit the XML.
- **QAT (quick-access toolbar):** a template's customUI *can't* populate the app QAT
  (Word only honors `<qat>` from a document, not an add-in). Tested 7/28/2026 on the build
  box: putting `<qat><documentControls>` in `LargePrintTemplate.dotx` — the *attached*
  template, which the note above does not cover — produces nothing either. So the QAT is
  set up by the **installer**: `installer/scripts/Merge-Qat.ps1 -Mode Mine|Vista|Restore|None`,
  chosen by the user via `[Tasks]` in `vistatype.iss`; `Remove-Qat.ps1` takes ours back off
  at uninstall. See CLAUDE.md for the modes and the reasoning behind dropping the old merge.
  `tools/lib/build_qat.py` generates `installer/qat-icons-only.officeUI` from the hidden
  `tab_LP_and_BRL_QAT_Icons` ribbon tab and validates `qat-template.officeUI` against it;
  `make build` fails if they disagree. (`qat-controls.xml` / `make qat`'s old behavior and
  `tools/lib/extract_qat.py` are obsolete.)
  Users restore an accidentally-removed icon by **re-running the installer** and choosing
  "keep my toolbar" — the "LP and BRL QAT Icons" tab is `visible="false"` and cannot be
  reached, so the old advice to right-click it was impossible to follow.

- **Toolbar facts worth not re-deriving** (measured on the build box, 7/28/2026):
  Word does **not** rewrite `Word.officeUI` on exit when the user changed nothing (same
  bytes, same timestamp), and does **not** lock it while running — a write made with Word
  open survives its exit and takes effect at the next start. That is what makes a
  "Switch Toolbar" button viable as plain VBA plus a restart prompt, with no helper process.
  The one caveat: if the user edits their toolbar through Word's own dialog in the same
  session, Word will write its version on exit and overwrite ours.

## One-time setup

1. On the **Windows box**: enable OpenSSH Server, set up key login so `ssh WIN_HOST`
   works with no password. In Word: *Options → Trust Center → Macro Settings →
   ☑ Trust access to the VBA project object model*.
   - Building the Windows box as a clean Hyper-V VM? Follow
     [`docs/Build-VM-Setup.md`](docs/Build-VM-Setup.md) — recommended, so the pipeline
     never touches your daily-driver Word and you can snapshot a pristine Word for
     installer testing.
2. On **Linux**: define a host alias in `~/.ssh/config` (keeps the hostname, user, and
   key path out of the repo), then point `build.config` at it:
   ```
   # ~/.ssh/config
   Host vistatype-build
       HostName windows-box.local      # or the VM's IP
       User builduser
       IdentityFile ~/.ssh/id_vistatype_build
   ```
   Then `cp build.config.example build.config` and set `WIN_HOST = vistatype-build`
   plus `WIN_DIR`. (You can also put the full `user@host` in `WIN_HOST` directly and
   skip the alias — but the alias keeps machine-specific details in `~/.ssh`, not `src/`.)
3. **Seed canonical source** (the current `src/` was bootstrapped by the Linux reader,
   which cannot produce valid `.frx`): run `make pull` once. This exports IDE-native
   `.bas/.cls/.frm/.frx` from `LPandBRL.dotm` into `src/`. Review with `git diff`, commit.

## Secrets / credentials

There are **no secrets in this repo, and none should ever be added.** The build pipeline
authenticates to the Windows box with an SSH **key pair**, and the private key lives in
`~/.ssh/` — outside the repo, shared with the rest of your SSH usage. The Makefile only
ever runs `ssh $(WIN_HOST)`; your SSH agent/config resolves the key.

- `build.config` (gitignored, seeded from `build.config.example`) holds connection
  *coordinates* only — `WIN_HOST`, `WIN_DIR`, paths. Nothing here is sensitive.
- **Never** put a private key, password, or token in the repo tree, even gitignored — a
  stray `git add -f` or a `.gitignore` typo would leak it. Keys stay in `~/.ssh/`, where
  their permissions and agent integration work correctly.
- No separate `.env` file is needed: `build.config` already is the per-machine,
  gitignored settings file.

## Everyday loop

**Work on `dev`, never on `master`.** `master` holds the last *released* version and only ever
moves forward by a fast-forward at release time. Check with `git branch --show-current`.

```
git checkout dev          # if you aren't already there
edit src/vba/*.bas        # or src/forms, src/ribbon — on Linux, in git
make build                # Word imports src/ -> dist/LPandBRL.dotm, copied back here
# smoke-test dist/LPandBRL.dotm in Word (see below), then:
make deploy               # promote dist/ artifacts to repo root
git add -A && git commit
```

- `make read` — regenerate `reference/vba-src/` from `LPandBRL.dotm` with the Linux-only
  decompressor. Handy for diffing what's actually compiled into the binary; needs no Windows.
- `make pull` — pull canonical VBA source back into `src/` after anyone edits in the VBE.

## Cutting a release

`master` = last released, `dev` = work in progress, one annotated `vX.Y.Z` tag per release,
merges are **`--ff-only`** (the tracked `.dotm`/`.dotx`/`.frx` binaries cannot be merged by
git — a fast-forward never tries). The full step-by-step checklist, the hotfix procedure, and
how to revert a bad release live in **`CLAUDE.md` → "Git workflow and releases"**, with a
plain-language walkthrough in **`docs/Daily-Workflow-and-Releases.md`**. Short form:

```
# on dev: bump the version in all four places, make installer, install & test, commit
git checkout master
git merge --ff-only dev
git tag -a v3.0.7 -m "VistaType LP 3.0.7"
git checkout dev
git push origin master dev --follow-tags      # only when you're ready to publish
gh release create v3.0.7 "dist/VistaType LP and Braille Macros Setup 3.0.7.exe" --title "VistaType LP 3.0.7"
gh release view v3.0.7 --json assets          # verify: must list the .exe, not []
```

**The `.exe` is the release.** A GitHub release without `"VistaType LP and Braille Macros Setup <ver>.exe"`
attached ships nothing — the "Source code (zip)" GitHub adds automatically is VBA text files,
which no transcriber can install. Push the tag and create the release-with-asset in the same
sitting; a bare tag already shows up as a "release" in GitHub's UI.

It is also the only true copy: the VBA build is not byte-reproducible (Word regenerates
p-code), so checking out an old tag and rebuilding does **not** give back the binary that
shipped. The uploaded installer is the real rollback.

## Always smoke-test a build

`make build` produces a `.dotm` but does not prove it runs. Before `make deploy`,
open `dist/LPandBRL.dotm` in Word once and confirm:
- the VBA project compiles (VBE → *Debug → Compile*);
- the two visible tabs appear (**VistaType LP**, **Braille Macros**) and merge with —
  don't replace — your normal ribbon. The third tab, **LP and BRL QAT Icons**, is
  `visible="false"` by design and must *not* show: it exists only so the Quick Access
  Toolbar's `idQ` references have ribbon controls to resolve against;
- a couple of buttons actually run their macro (the `RibbonAction` dispatcher);
- **which set of tabs you are looking at.** Opening the built `.dotm` by hand shows the
  *embedded* tabs, which will not be listed in *Customize the Ribbon* — correct for that
  path. To test the user-installed tabs you have to run the Setup.exe with the **Ribbon
  tabs** task ticked, then look for them in that dialog. Testing the wrong one and
  concluding the feature is broken is an easy half-hour to lose.

Enable VBE → *Tools → Options → General → Show ToolTips* and turn on ribbon load
errors (`File → Options → Advanced → General → Show add-in user interface errors`)
to catch a bad callback or imageMso. This is the one manual Word step; the rest is
automated.

## Gotchas baked into the tooling

- **`.frm` files MUST keep CRLF line endings.** A UserForm `.frm` starts with a designer
  header block (`VERSION 5.00` / `Begin {GUID} FormName … End`) that Word's importer parses
  *before* the code. That parser requires **CRLF**. Rewrite a `.frm` with LF endings — easy to
  do from Linux, e.g. any Python `open(path,"w")` — and Word silently fails to recognise the
  header, treats those lines as VBA source, and drops `VERSION 5#` / `Begin {…}` into the
  form's **code module**. The build still succeeds; the failure only shows on the user's
  machine as **"Compile error in hidden module: &lt;FormName&gt;"**. Cost a debugging round on
  2026-07-26. Check with `python3 -c "d=open(p,'rb').read(); print(d.count(b'\n')-d.count(b'\r\n'))"`
  — any bare LF is a bug. Fix by rewriting the file in binary: read bytes, `replace(b"\n", b"\r\n")`,
  write bytes. **Editing a `.frm` from Linux at all is only safe for the code below the header;
  never let a text-mode write touch one.** (`.bas` is *not* affected —
  `src/vba/LPandBrlMacros.bas` has been LF for its whole life and imports fine — so the two
  file types genuinely differ here.)
- **Stale Word lock files hang the build forever — the build now clears them.** Word writes a
  hidden `~$<name>` "owner file" next to every open document and deletes it on a clean close.
  A Word that was **killed** leaves one behind. The next build opens or saves that document,
  Word reads the leftover owner file, decides someone else has it locked, and raises the modal
  **"File In Use"** dialog. Headless over SSH nobody can answer it, so the build hangs
  indefinitely — **no error, nothing in the Windows event log, Word alive and "responding" at
  near-zero CPU.** The usual recovery (kill the hung Word) leaves a *fresh* owner file, so
  every following build hangs identically. This cost a full day on 2026-07-26: a
  `~$andBRL.dotm` from 2026-07-22 had been blocking every build in between.
  `Import-Vba.ps1` now deletes `~$*` from the shell and output directories before starting
  Word, and warns if a `WINWORD.EXE` is already running. **These files are hidden**, so plain
  `del` skips them and reports "Could Not Find" — which reads like they were already gone.
  To clear by hand: `ssh vistabuild 'del /a /q /s "C:\Users\jerry\vistatype-build\~$*"'`.
  Diagnosing a hang: check whether the import log stops (`Documents.Open` produces no output
  at all; a stall after the last `import module` line means `$doc.Save()`), then
  `Get-Process WINWORD | Select Id,CPU` — flat CPU over minutes means it is waiting on a
  dialog, not working. **Word's CPU is a poor progress signal; the import log is the reliable
  one.**
- **The VBA project must not be locked.** A password-locked project ("Lock project for
  viewing" in the VBE) enumerates as **0 components** through Word's object model, and
  there is no API to unlock it — so `make pull`/`make build` would silently produce an
  empty result. Both scripts now hard-fail on a 0-component project instead. If you see
  that error, open the shell in Word's VBE (Alt+F11 → Tools → *&lt;project&gt; Properties*
  → **Protection** → uncheck "Lock project for viewing", clear the password), save, and
  re-seed the shell. Keep the shipped project unlocked (the source is GPLv3 and public
  anyway).
- **Encoding: repo source is UTF-8; Word uses ANSI (Windows-1252).** Word exports/imports
  `.bas`/`.cls` in the system ANSI codepage, but the repo keeps them UTF-8. `Export-Vba.ps1`
  transcodes ANSI→UTF-8 after export and `Import-Vba.ps1` transcodes UTF-8→ANSI before
  import, so characters like the smart quotes / ellipsis / bullet / en-dash / fractions
  used in the find-and-replace macros round-trip losslessly. Forms (`.frm`) are ASCII and
  the `.frx` is binary, so both are left exactly as Word writes them. (The Linux reader
  `decompress_vba.py` also decodes cp1252 — decoding as latin-1 turns those bytes into C1
  control codes.)
- **Build renames the VBA project to `LPandBRL`.** The add-in ships in Word's STARTUP folder
  loaded alongside the user's own `Normal.dotm`; two loaded projects can't both be named
  `Normal`, so `Import-Vba.ps1` renames the built project (`-ProjectName`, default `LPandBRL`,
  fed by `PROJNAME` in the Makefile). Because STARTUP globals don't fire `AutoOpen` per
  document, document-type detection runs through Word application events (the `VtEvents` class,
  hooked by `AutoExec`) — see CLAUDE.md "Notable modules". Nothing references the project name,
  so the rename is safe.
- **`ThisDocument`** is a Document module — it can't be Import-ed. `Import-Vba.ps1`
  clears its code module and refills it from `src/vba/ThisDocument.cls`.
- **Forms** must round-trip as `.frm` + `.frx` (the `.frx` holds images/binary layout).
  The Linux reader can't rebuild `.frx`; that's why `make pull` (Word export) is the
  canonical seeder for `src/forms/`.
- **Project references** (MSWORD.OLB, FM20.DLL/MSForms, scrrun.dll, stdole) and the
  attached-toolbar ribbon live in the shell `LPandBRL.dotm`, not in `src/`. The build
  imports code *into a copy of that shell*, so those are preserved automatically.
- **Version bumps**: follow the in-file convention — update the per-sub `' Version`
  comment and add a dated line to the `LPandBrlMacros` header changelog and the About
  forms' version string.

## Keeping CLAUDE.md in sync

`CLAUDE.md` documents the build pipeline and architecture. A project `PostToolUse` hook
(`.claude/hooks/claude-md-sync.py`, registered in `.claude/settings.json`) watches for
edits to pipeline-defining files — this file, the `Makefile`, and anything under
`tools/`, `src/ribbon/`, or `installer/` — and reminds Claude Code to reconcile
`CLAUDE.md` in the same change. It only nudges; the actual edit is Claude's. Disable or
review it via `/hooks`.

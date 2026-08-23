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

> **These rules are now system-wide as well, in `~/.claude/CLAUDE.md`** (Todd, 8/22/2026), so
> they apply in every project and in ones that do not exist yet. They were only here until then,
> which meant a session started outside this folder ran without them — and one did, on
> 8/22/2026. What follows is kept because it is more specific: the Word and VBA vocabulary below
> is this project's, and the examples are drawn from it. If the two ever disagree, this file wins
> for this project — but they should not, so change both.

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

## Converting a macro off the temporary document

`Lp_Copy_To_Temp_Doc` and `Dx_Copy_To_Temp_Doc` create the scratch document hidden and then
deliberately show, maximize and activate it — four statements, not an accident. That is the
screen flashing. They must, because their callers work through `Selection`, and `Selection` only
reaches the active document. Converting a macro means moving its passes onto a **range**, so the
window is never needed. The round trip itself stays: it is what gives the single undo and the
private workspace.

**Jerry's rule, 8/12/2026: whenever converting a pair leaves the `Lp_` and `Dx_` versions
identical, merge them into one `Sh_` macro.** That is not a coincidence when it happens —
everything that differed between a pair usually existed to serve the round trip, and goes with
it. Merging as we go makes the second half of the list shorter rather than longer.

Worked examples: `Lp_Horz_List_To_Vertical` (3.0.147, hidden scratch document) and
`Sh_Remove_Multi_Spaces` (3.0.149, no scratch document at all, and the first merge under the
rule). The full list, what to test, and the two paragraph-mark traps are in
`docs/Temp-Doc-Conversion-Checklist.md`.

## American spellings, everywhere

Jerry's rule, 8/8/2026: **this product is for American transcribers, so everything in it uses
American spellings** — dialog text, installer wording, documentation, and code comments alike.
license (not licence), color, behavior, recognize, normalize, gray, dialog. The whole repo was
swept on 8/8/2026; keep new writing in step.

**This covers what you SAY to Jerry, not only what you write into the product** — and usage as
well as spelling. He corrected "double full stop" to "double periods" on 8/8/2026. Period, not
full stop; parentheses, not brackets; quotation marks, not inverted commas.

Two files are **excluded and must stay untouched**: `LICENSE` (the GPL's own text) and
`docs/Software-Agreement.md`. A third, `assets/fonts/atkinson-hyperlegible/OFL.txt`, was on this
list until the bundled typeface was dropped on 8/20/2026 and `assets/fonts/` deleted.

No style name, ribbon id or control id contains any of these words, which is why the sweep was
safe. Check that again before extending it: renaming an id breaks toolbars already installed in
the field (see the `btn_*` rule further down).

## The pieces and how they work together

The source of truth is `src/`. The items below are the *built* artifacts (and their
sources) that ship or that Word loads.

| Piece | What it is | Role |
|------|-----------|------|
| `LPandBRL.dotm` | Word add-in template (macro-enabled), built from `src/` | **The code.** The entire compiled VBA project — 236 subs/functions in `LPandBrlMacros`, plus 49 UserForms — and the embedded ribbon (below). Loaded from Word's `STARTUP` folder, so its macros are available to every document. Behavior is *authored* under `src/vba`/`src/forms`; this is where it *runs*. |
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
- **REMOVED, not installed:** the four `VistaTypeLPLegible-*.ttf` faces and their `OFL.txt`. The
  bundled typeface shipped from 3.0.101 (8/8/2026) to 3.0.196 (8/20/2026) and was **dropped** —
  its character set is Latin, so Latin proper, mathematics, the IPA and the Greek that runs
  through medical transcription all fall back to another face at another size, silently, which is
  the one thing large print must never do. Jerry's call, 8/20/2026.
  `RemoveLegacyLegibleFont` in `[Code]` now takes it back off a machine that has it: every build
  from 3.0.101 on installed it `uninsneveruninstall`, so not shipping it would have left it on
  every machine for ever. The **registry** half is what retires the face (Windows will not
  release a per-user font mid-session, so the `.ttf` often cannot be deleted until a later run);
  nothing in that procedure is fatal or reported to the transcriber. It runs at **both**
  `ssPostInstall` and `usUninstall` — the first pass usually leaves the locked files behind, and
  without the second an uninstall would strand them for ever. It deletes the `VistaType LP Fonts`
  license folder only when no `.ttf` of ours is left AND something was actually unregistered:
  counting this run's successes instead would read a hand-installed machine-wide copy as "all
  gone" and strip the license off a font still on the machine. The per-user Fonts folder is named
  literally (`#define UserFontsDir`) rather than via `{autofonts}`, so a removal cannot raise on a
  Windows too old to have per-user fonts. This is safe only because
  `Lp_Attach_The_Template` embeds the face, unsubsetted, in every book set in it — those books
  carry their own copy. Do not add a bundled font back without testing character COVERAGE first:
  Greek, the IPA, the math operators.
- `%AppData%\VistaType LP Settings\VistaType.ini` — **written by the add-in, never by the installer,
  and deliberately NOT in `[UninstallDelete]`.** Holds the transcriber's own Word settings that
  VistaType puts back for her: from 8/18/2026 the five spelling/grammar settings braille and large
  print switch off (`CheckGrammarAsYouType`, `IgnoreMixedDigits`, `ContextualSpeller`,
  `LabelSmartTags`, `IgnoreUppercase`), and whatever the saved-settings library grows to hold. A
  *separate folder* from `%AppData%\VistaType LP` on purpose — that one is wiped whole at uninstall,
  and a reinstall must not cost her her settings; same reasoning as the font license folder. Read and
  written with Word's own `System.PrivateProfileString` via `Sh_Settings_File` /
  `Sh_Setting_Read` / `Sh_Setting_Write` in `LPandBrlMacros`. **A file, not module variables:** VBA's
  `End` statement resets every module-level variable, and this project runs `End` on ordinary paths
  (34 times in `LPandBrlMacros` alone, plus dialog Cancel buttons and `Sh_Is_Doc_Open`), so anything
  held in memory would be lost mid-session — and the next save would record VistaType's own value as
  the transcriber's choice. See `docs/User-Settings-And-Word-Configuration.md`.
- The **Quick Access Toolbar** is set up per the user's choice on the install wizard (append VistaType's icons / install VistaType's toolbar whole / restore their pre-VistaType one / leave it alone), written to `Word.officeUI` in **both** Roaming and Local; their ribbon is never touched and the embedded ribbon supplies the tabs

## Repo layout

```
src/vba/        canonical VBA text: *.bas (std modules), *.cls (class/document modules)
src/forms/      canonical UserForms: *.frm + *.frx (binary layout)
                *.frm MUST stay CRLF: Word parses the designer header at the top of the file,
                and LF endings make it dump that header into the form's CODE module - which
                fails only on the user's machine, as "Compile error in hidden module".
                `make build` now refuses to run if any .frm has bare LF. See DEVELOPMENT.md.
                A file DELETED from src/ is now really gone from the build: `push-src` wipes the
                build box's src/ and tools/ before copying, because scp only adds and overwrites.
                Without that, Import-Vba.ps1 kept re-importing removed forms from the box's stale
                copy and they shipped in every later .dotm - three of them were, one since
                8/3/2026 - and nothing said so, since the build log lists what it REMOVES from the
                .dotm, not what it puts back.
src/ribbon/     customUI14.xml — embedded ribbon (source of truth); Word.officeUI (legacy)
                From 3.0.34 the two VISIBLE tabs are also written into the user's own
                Word.officeUI by the installer (generated: installer/ribbon-tabs.officeUI),
                because Word does not list add-in tabs in Customize the Ribbon and they could
                therefore not be hidden, reordered or renamed. ALWAYS written from 3.0.118 -
                the installer's "Ribbon tabs:" task is gone (Jerry, 8/9/2026: declining it was
                the option that made Word behave abnormally, so it should not have been one).
                The embedded copies stay as the
                fallback, hidden by getVisible when the user has their own. See DEVELOPMENT.md
                "Embedded ribbon" for the two-homes table and the orphan-tab trade-off.
src/keymap/     lp-template-keymap.xml — keyboard shortcuts, injected into LargePrintTemplate.dotx
                at build time. Source of truth for the ~60 key assignments (53 style shortcuts,
                8 VistaType macros, 1 MathType). Macro targets MUST read
                LPANDBRL.LPANDBRLMACROS.<SUB>: they said NORMAL.NEWMACROS.<SUB> until 7/28/2026
                and had been silently dead ever since the code left Normal.dotm.
LPandBRL.dotm   the .dotm shell/base (tracked): project references + non-VBA parts; build base
LargePrintTemplate.dotx   the attached large-print template (styles/page setup)
Word.officeUI   legacy global ribbon (no longer shipped; kept for reference)
tools/windows/  Export-Vba.ps1 / Import-Vba.ps1 / New-UserForm.ps1 — run in Word on the build box
                (New-UserForm.ps1 creates a BRAND-NEW UserForm and exports a valid .frm/.frx
                 pair: the binary .frx cannot be written by hand from Linux and `make pull`
                 must never be used to seed one. Builds it in a throwaway blank document, so
                 nothing tracked is touched. Layout only — append the form's code to the .frm
                 as text afterwards, then run tools/lib/check_frm_eol.py. `-InfoDialog` lays in
                 the read-only note shape used by Sh_Prodnote_Info_Form; `-MessageDialog` lays in
                 the shared message shape used by Sh_Message_Form — the same text box plus an
                 Okay and a Cancel button, captioned and accelerated per *How VistaType LP says
                 things* below. It applies Tahoma 10, the ControlTipText on every button and the
                 O/C accelerators itself, so a form built with it is already right)
                (Import-Vba.ps1 clears stale hidden `~$*` Word lock files first — a leftover
                 one makes Word raise an invisible "File In Use" dialog and the build hangs
                 forever with no error; see DEVELOPMENT.md "Gotchas baked into the tooling")
                Remove-FormControls.ps1 — takes named controls OFF a UserForm and rewrites the
                 .frm/.frx pair, headlessly, in a throwaway blank document. A control lives in the
                 BINARY .frx, so deleting its lines from the .frm leaves it on the form, still
                 drawn, with no code behind it. REMOVE THE CODE FIRST: a handler naming a control
                 that is gone fails nowhere in this pipeline and reaches the transcriber as
                 run-time error 424 when the dialog opens (these forms have no Option Explicit,
                 so a name that is no longer a control is an empty Variant). -Controls is ONE
                 comma-separated string, not an array — PowerShell's -File never parses an
                 argument as an expression. It reports each control's position before removing it,
                 which is the only record of the hole left behind. Added 8/20/2026 to take the
                 Typeface choice off LP_Attach_An_Lp_Template_Form
                Setup-SigningTools.ps1 — one-time box preparation for CODE SIGNING, added
                 8/17/2026 when Jerry bought the Certum card. Installs only the Windows SDK's
                 SigningTools feature (which lands in the usual C:\Program Files (x86)\Windows
                 Kits\10) and Microsoft's Office Subject Interface Packages (msosip/msosipx —
                 what lets signtool sign the VBA project inside a .dotm; they do NOT ship with
                 Office) into C:\vt-signing, which is OUTSIDE the build folder because src/,
                 tools/ and installer/ are wiped there before every build. Those two libraries
                 are registered MACHINE-WIDE and point at that folder, so run the script with
                 -Unregister before ever deleting or moving it: a dangling registration breaks
                 signature checking on Word files for every process on the box, silently. Checks
                 every download with Get-AuthenticodeSignature and refuses to register a library
                 that is not validly signed. Needs no card, and nothing in the daily build calls
                 it. Two findings from the rehearsal that will save a day: the **x64** signtool
                 signs a .dotm and the x86 one always fails with SignerSign() 0x800403f4, which
                 is the OPPOSITE of Microsoft's own instructions (Word here is 64-bit and the
                 library must match); and the ribbon injection does NOT break a macro signature,
                 so signing can sit on either side of it. Full write-up, including what to do the
                 day the card arrives, is docs/Code-Signing.md)
tools/lib/      check_style_guards.py (refuses to build when an ActiveDocument.Styles("name") lookup is
                not behind Sh_Style_Exists / Sh_Style_In_Use. Styles(name) raises run-time error 5941 on a
                document that does not carry the style, and documents legitimately do not: RefPageNemeth
                comes from the Nemeth braille templates, Print Pg Num from the LP one. The macro dies
                where it stands, usually before doing anything the user can see. All 158 were guarded on
                8/5/2026 after one cost Jerry an AutoTag run);
                check_form_calls.py (refuses to build when a FORM calls a macro that exists nowhere
                under src/vba — nothing in the build compiles VBA, so that ships and fails on the
                user's machine as "Compile error in hidden module: <form name>", naming the form
                and not the missing macro; it also WARNS, without failing, about Application.Run
                "name" string calls, which compile and fail only when the button is pressed);
                check_vba_structure.py (refuses to build on the structural mistakes that COMPILE nowhere in
                this pipeline and so ship silently: a module-level Const/Dim/Type written beside the
                procedure that uses it instead of in the declarations section at the top of the module,
                a procedure name declared twice in one module, or a With with no End With. All three
                reach the transcriber as "Compile error in hidden module: <name>" with no line number
                and nothing naming the cause. Cost a build on 8/18/2026 - three Private Const lines for
                the settings store. Covers .bas, .cls and the code section of .frm);
                decompress_vba.py (reader); officeui_to_customui.py + inject_customui.py (ribbon);
                inject_keymap.py (keyboard shortcuts -> .dotx/.dotm); build_ribbon_tabs.py
                (ribbon tabs -> installer/ribbon-tabs.officeUI, plus the shipped-button-id
                guard); trim_frm_blanks.py (collapses the blank lines Word adds to a form's code
                section on every export — the two About forms are re-exported by the version
                stamper on EVERY build, so the runs grow without limit; runs as part of `make build`);
                build_branding.py (`make branding` — turns the two artwork PNGs in
                assets/branding/ into installer/branding/: vistatype.ico at nine sizes, four
                sizes of the welcome-page panel and seven of the small corner image, so a
                high-DPI screen gets a sharp one instead of a stretched one. Needs Pillow.
                `make installer` runs --check-only and refuses to build if they are missing,
                but does NOT regenerate — a machine without Pillow can still build everything
                else. Regenerate, never hand-edit. Added 8/15/2026 with the version block:
                until then the Setup.exe was a generic Inno stub reporting FileVersion 0.0.0.0,
                and "no version information" sits beside "unsigned" as one of the two
                high-weight factors behind the machine-learning deletions — docs/Code-Signing.md);
                scan_virustotal.py (`make scan` — uploads the built Setup.exe to VirusTotal and
                reports which of its ~70 engines flag it. Exists because Defender began deleting
                every build on 8/13/2026 as Trojan:Win32/Bearfoos.B!ml, a machine-learning GUESS,
                and two days of theorizing produced four wrong answers. Fails the run if
                MICROSOFT flags it — that engine is what a transcriber has — or if more than
                --max (3) do, and exits 2 - never 0 - on a report with no engine results in it,
                because a gate that fails open is worse than none. Deliberately NOT part of
                `make installer`: the upload is PUBLIC and PERMANENT and cannot be withdrawn, so
                it must be a decision, never a side effect. For the same reason it REFUSES
                anything that is not a .exe directly inside dist/, with no override - the
                likeliest thing anyone would ever name here is a transcriber's document that
                Defender ate. VirusTotal runs Defender WITHOUT its cloud and !ml
                verdicts are cloud verdicts, so a clean line there is not proof — confirm with
                MpCmdRun -DisableRemediation on the build box. Reads VT_API_KEY from the
                gitignored build.config. Background: docs/Code-Signing.md);
                extract_qat.py (obsolete/reference — QAT is now qat-template.officeUI)
                (rescale_font.py stood here until 8/20/2026, when the bundled typeface was
                dropped — see assets/fonts/ below)
assets/branding/  vistatype-icon.png + vistatype-wordmark.png — Jerry's artwork, 8/15/2026, the
                source of truth for how the installer looks. Never edited by the build.
                installer/branding/ is DERIVED from these by tools/lib/build_branding.py and is a
                build output. The generated files live under installer/ and not beside the
                artwork ON PURPOSE: installer/ is wiped on the build box before each copy, so a
                deleted one really goes, whereas assets/ is never wiped there. README.md records
                what is made, the four constants that move the welcome panel about, and why the
                installer has an icon at all.
installer/      Like src/ and tools/, this folder is WIPED on the build box before each copy, so a
                file deleted here is really gone from the next Setup.exe. scp only adds and
                overwrites; see the note on push-src above for what that cost when it bit.
                Inno Setup installer (vistatype.iss) + scripts/ (QAT merge/remove) + qat-template.officeUI
                (the standard QAT — HAND-maintained: its order, separators and visible="false"
                entries suppressing Word's default buttons are deliberate and cannot be derived)
                + qat-icons-only.officeUI (GENERATED by tools/lib/build_qat.py — VistaType's 8
                icons and nothing else, safe to append to a user's own toolbar). `make build`
                validates every x1:btn_* in the curated toolbar against the hidden ribbon tab
                and stops if they disagree; a stale reference renders as a blank button on the
                user's machine with no warning.
docs/           Daily-Workflow-and-Releases.md (Jerry's plain-language guide to dev/master, building, releasing, and what to ask Claude); Installation-Guide.md (end-user install); Build-VM-Setup.md (Hyper-V build/test box); Software-Agreement.md (GPLv3 About-dialog text); Code-Signing.md (why Defender deleted the unsigned Setup.exe on 8/13-14/2026, what in the installer scores against it, how to test with `make scan` and MpCmdRun, and the signing options — note Azure Artifact Signing does NOT sign VBA projects, and EV no longer skips SmartScreen. Jerry bought the Certum open source card on 8/17/2026; the last section is the step-by-step for the day it arrives, what the build box already has, and the five things the card-free rehearsal proved)
reference/      generated read aids (gitignored mirror + interim form-code dump)
                (assets/fonts/ was here. It held the bundled typeface, VistaTypeLP Legible,
                shipped from 3.0.101 to 3.0.196 and DROPPED on 8/20/2026 along with
                tools/lib/rescale_font.py, tools/windows/Check-Font.ps1 and `make fonts`. The
                face has no Greek, no IPA and almost no mathematics, and Word substitutes a
                missing character silently and at the wrong size — see Deployment locations and
                Domain concepts above. The pristine Atkinson Hyperlegible download, the OFL text
                and the whole licensing write-up are in git history if a bundled face is ever
                wanted again; the Braille Institute's license PDF is still in
                ~/reference/vistatype-lp/ and must never be committed.)
.claude/agents/ four review helpers Claude hands work to, each with its own reading space so
                the ~16,500-line engine does not crowd out the job in hand. vba-review (the
                failure modes that ship silently — run before `make build`); braille-lp-review
                (what the READER gets: LP formatting, UEB/Nemeth/BANA output, $pg tags);
                ship-safety (what a change does to a transcriber's own toolbar, ribbon and
                %AppData% — run before any release); release-check (verifies version numbers,
                that the Setup.exe exists, that a published release carries its .exe, and that
                a hotfix reached dev). All four are READ-ONLY and report; none edits or pushes.
                They complement tools/lib's guards rather than repeat them — each file says
                what the guards already cover.
Makefile        pull / build / ribbon / qat / read / deploy / branding / stage / installer / scan
                (`make fonts` went with the bundled typeface on 8/20/2026)
                (see DEVELOPMENT.md)
```

The ribbon is **embedded** in `LPandBRL.dotm` (`src/ribbon/customUI14.xml`), so it merges
with each user's ribbon instead of overwriting it — the old `Word.officeUI` file is no
longer shipped. Every ribbon button routes through one VBA dispatcher, `RibbonAction`
(`src/vba/RibbonCallbacks.bas`), which runs the macro named in the control's `tag`. The
QAT can't be set from a template's customUI, so the **installer** sets it up —
`installer/scripts/Merge-Qat.ps1`, undone by `Remove-Qat.ps1`.

**The user chooses, and the merge is gone (3.0.33).** Up to 3.0.32 the installer merged
VistaType's curated toolbar into whatever the user had. That is what "caused no end of
problems": it pushed their icons right, discarded their separators, and hid ten of Word's own
buttons — **Undo, Redo and AutoSave among them**. Those `visible="false"` entries are not a
bug; suppressing Word's defaults is the only way to make a curated toolbar look clean. The
bug was doing it *by merge*, on top of someone else's toolbar. Jerry's call, 7/28/2026:
**replace, don't merge — save theirs, install ours whole, give them a way back.**

`Merge-Qat.ps1 -Mode` now takes:

| Mode | Behavior | Installer task |
|---|---|---|
| `Mine` | Keep their toolbar untouched; append VistaType's icons. Nothing hidden, moved or reordered. | `qat\mine` (default) |
| `Vista` | Write the curated toolbar whole. No merging at all. | `qat\vista` |
| `Restore` | Put back `.vtqatbak` (their pre-VistaType toolbar), then append our icons. | `qat\restore`, shown only when `HasQatBackup` |
| `None` | Touch nothing. | parent task unchecked |

`.vtqatbak` is written **once** and never overwritten, so it is still the user's true
original years later — that is what makes `Restore` a real rescue. **Never delete it**, not
even at uninstall. Merge also writes a `.vtqatmanifest` (exactly what it laid down, so
uninstall removes precisely that and nothing else) and a `.vtqatprev` per-run snapshot;
both are logged to `%AppData%\VistaType LP\qat.log`.

`Remove-Qat.ps1` no longer restores the whole file over the top — that silently destroyed
every toolbar *and ribbon* change made since installing. It removes the manifest's entries,
falls back to identifying ours by namespace when there is no manifest, and only reaches for
`.vtqatbak` when our removal would leave the toolbar empty (i.e. `Vista` mode) — re-appending
anything the user added on top of ours.

Three subtleties that make it actually work:
- **Location:** Word reads `Word.officeUI` from `%APPDATA%` (Roaming) on most machines but from
  `%LOCALAPPDATA%` (Local) when the profile roams/redirects or Office can't roam, so the merge
  writes **both** (Word honors whichever it uses; the other is ignored).
- **Every machine that ran VistaType before 3.0 still carries our OLD ribbon tabs**, and they
  must be recognized and replaced on upgrade. Before 3.0 the install was a manual file copy
  that included the global `Word.officeUI` still tracked at the repo root, and it REPLACED the
  user's ribbon. Those tabs are `mso_c1.F9211` "Braille Macros", `mso_c1.56E551D` "VistaType
  LP" and `mso_c1.4EA2EBA` "LP and BRL QAT Icons" — Word-generated ids, and controls that no
  longer resolve to the add-in. Nothing identified them, so the installer added the current
  pair beside them: two tabs of each name, the older missing every button added since, and the
  transcriber clicking whichever came first. Jerry hit it on 8/6/2026 and it took eight builds
  and four wrong theories to find, because it only appears on an upgrade and a clean install
  was always perfect.

  A tab or group of ours is now recognized three ways, and `Merge-Qat.ps1` needs **all three**:
  the `vt_*` id we write; the add-in controls it carries; or **the same id as one of the
  template's groups with the decoration stripped** — `vt_grp_mso_c1_18B5F8FF` and
  `mso_c1.18B5F8FF` are the same group, ours prefixed with dots turned into underscores. That
  last test is the only thing that sees a pre-3.0 group.

  Apply it in the REFRESH path as well as the duplicate path. Putting it in only one (3.0.96)
  left the legacy groups in place and appended ours beside them — a tab with fourteen groups,
  half of them drawing as empty placeholders because their controls point nowhere.
- **Format:** each VistaType QAT entry is a **reference to the add-in's own ribbon control** —
  `<mso:control idQ="x1:btn_<macro>">`, where the `x1` namespace is the installed
  `LPandBRL.dotm`'s full path — the exact shape Word itself writes when a user adds one of our
  ribbon buttons to the QAT by hand. (Standalone `onAction` macro buttons did **not** display.)
- **Template:** `qat-template.officeUI` is hand-edited (`__VT_DOTM_PATH__` is substituted with the
  install path at merge time); its `btn_<macro>` ids must match `src/ribbon/customUI14.xml`.
  `make build` now *enforces* that via `tools/lib/build_qat.py` — they had already drifted
  (the toolbar carries `ParagraphMarks`, which is not on the tab). It also generates
  `qat-icons-only.officeUI`. **Never rename the eight `btn_*` ids on the hidden tab**: every
  toolbar already installed in the field references them by `idQ`.
- **XML comments may not contain `--`.** Word and PowerShell both refuse to parse the file,
  and the symptom is an install that silently leaves the toolbar alone. `build_qat.py`
  parse-checks both toolbar files for exactly this reason. **Inno's Pascal has the same trap
  in a different costume:** `{ }` are its comment delimiters, so an inline `{code:...}`
  reference written inside a comment ends it early and the rest of the sentence is compiled.
  Both cost a build on 7/28–29/2026.
- **Never rename or remove a `btn_*` id** in `customUI14.xml`. Every ribbon tab and toolbar
  already installed in the field references buttons by id; a renamed one renders blank on the
  user's machine with no error, and their `Word.officeUI` is not ours to migrate.
  `installer/ribbon-button-ids.txt` is an append-only record of every id that has shipped and
  `make build` stops if one disappears. This supersedes the narrower warning about the six
  toolbar ids: it now covers all 50.

## Build & edit workflow (short version)

Full detail in **`DEVELOPMENT.md`**. In brief: **on the `dev` branch** (never `master` — see
*Git workflow and releases* below) edit text under `src/`, then `make build` ships it to the
remote Windows+Word box over SSH, which imports the source into `dist/LPandBRL.dotm` (Word
regenerates p-code) and copies it back; the ribbon (`inject_customui.py`) and the keyboard
shortcuts (`inject_keymap.py`) are then injected on this side by plain zip surgery, no Word
needed; smoke-test in Word, then `make deploy`. Never edit
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
- **49 UserForms** — dialogs, prefixed by domain (see below). `Sh_Message_Form` is the shared
  message dialog every `Sh_Say` / `Sh_Ask` goes through — see *How VistaType LP says things*.

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

### How VistaType LP says things — dialogs and messages

Jerry's rule, 8/23/2026. It governs every new dialog and every message:

- **The message is at least 10 point Tahoma.** That is the floor, not the target — a bigger
  message is fine, a smaller one is not. The people using this software work on large print
  and braille, and several of them have low vision themselves.
- **The OK button is captioned `Okay`, never `OK`.**
- **`Okay` has an accelerator of `O`; `Cancel` has `C`** — Alt+O and Alt+C press them.
- Every CommandButton still gets a `ControlTipText` of its caption plus " Button", and a form
  is still Tahoma 10 throughout (the 8/7/2026 rules; both stand).

**A `MsgBox` can do none of the three**, so a message is a UserForm here. Its font is whichever
one Windows draws a message box in, its button says `OK` and cannot be changed, and it has no
accelerators. `Sh_Message_Form` is the one dialog every message goes through, reached by two
subs in `LPandBrlMacros` and never touched directly:

```vba
Sh_Say "text", "VistaType LP (nnn)"             ' tell her. One Okay button.
If Sh_Ask("text", "VistaType LP (nnn)") Then    ' ask. Okay and Cancel; True means Okay.
```

The second argument is the title bar, and that is where the dialog's own number lives. Those
numbers identify one dialog out of all of them when a transcriber says what she saw, so a new
message takes the next unused number and an existing one keeps the number it has always had.
(Written `nnn` above on purpose — a real number in an example is one more hit to wade through
when hunting for the next unused one.)

Paragraph breaks are written `vbCr`, as everywhere else in this project. `Sh_Message_Form`
turns them into `vbCrLf` on the way into the text box, because an MSForms text box is not a
`MsgBox` and the one other form in this project that fills one uses `vbCrLf` throughout.

What is lost against `MsgBox`: its information / warning / question **icon**, and the sound
the warning one made. A UserForm has neither.

**Not retrofitted.** The two newest toolbar buttons — Reset Word Configuration and Styles Pane:
Recommended — were converted on 8/23/2026, five messages. Counted the same day, what is left:
**85 `MsgBox` calls in `LPandBrlMacros`, 48 inside the forms, 20 in the three smaller modules.**
They convert a feature at a time. Nothing new uses `MsgBox`.

**Hover text cannot be made 10 point Tahoma, on a ribbon button or on a form.** Word draws a
ribbon/QAT screentip and supertip itself in the Office UI font; customUI has no font attribute
of any kind, and a `getSupertip` callback supplies the words and nothing else. `ControlTipText`
on a UserForm control is a Windows tooltip and is the same. The only thing that changes either
is the reader's own Windows text size (Settings → Accessibility → Text size), which changes it
everywhere. Where hover text needs to be readable at 10 point or more, the answer is to put the
words in the dialog, which VistaType LP does control.

### Editing convention

Every sub is versioned inline via a comment block (Version/Date/Author). The module header
of `LPandBrlMacros` keeps a running dated changelog.

**The version bumps itself on every `make installer`** (Jerry's rule, 2026-07-25) so that no
two test installers ever carry the same number — three builds shared `3.0.6` on 2026-07-26
and there was no way to tell from the About box whether an install had taken. `make bump`
increments the third digit in `APPVER` (Makefile) and `AppVer` (`installer/vistatype.iss`),
then `Import-Vba.ps1` stamps the LP **and** Braille About dialogs during the build and exports
those two forms back into `src/forms/` (their captions live in the binary `.frx`, so they can
only be written from inside Word). The two dialogs use **different control names** for that
caption — `VersionLabel` in LP, `Label5` in Braille — so the stamper matches on the caption
text (`"This computer is running Version…"`), not the control name. Nothing else needs editing —
this file describes the scheme but stores no live number. The bump is left **uncommitted**;
commit it with the work it belongs to. **What actually shipped
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
6. **Confirm the installer exists before going near git** — `ls -l dist/"VistaType LP and Braille Macros Setup <ver>.exe"`.
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
   gh release create v3.1 "dist/VistaType LP and Braille Macros Setup 3.1.exe" \
       --title "VistaType LP 3.1" --notes "<what changed, in transcriber-facing terms>"
   gh release view v3.1 --json assets      # VERIFY: must list the .exe, not []
   ```

9. Delete the superseded `"VistaType LP and Braille Macros Setup *.exe"` from `dist/` and the VM Desktop (keep old
   real releases).
10. **Remind Jerry to update the website — and do not touch it yourself.** See *The website is a
    separate project* immediately below. Give him the paste-able prompt from
    `docs/Daily-Workflow-and-Releases.md` → *Step 7*.

### The website is a separate project — never merge them

**`~/projects/vistatypelp-org`** (live at **vistatypelp.org**) is the public website. It is its
own repository with its own `CLAUDE.md`, and **the separation is deliberate and permanent**.

**Jerry's rule, 8/15/2026: never merge the two, and never suggest merging them** — not as a
monorepo, not as a subfolder, not "while we're in here". Do not propose it as a tidy-up, and do
not edit the website from this project even when the change looks trivial. If something there
needs doing, say so and stop.

They are related in exactly one way: **when a real release goes out (an `X.Y`, not a `3.0.X`
build), the website should be looked at.** Nothing breaks if it is not — the site's download
button points at "the newest release" on purpose, so releasing never *requires* a website change.
What goes stale is the writing: what the add-in does, what is new, and the install page.

So at step 10, tell Jerry to **open Claude in a second session in that folder**, and hand him the
ready-made prompt in `docs/Daily-Workflow-and-Releases.md` → *Step 7* with the version number
filled in. That session has the website's own guidance and its own "never publish without being
asked" rule. This one does not do the work.

### The `.exe` IS the release — non-negotiable

**A GitHub release without the `"VistaType LP and Braille Macros Setup <ver>.exe"` attached is not a release.**
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
gh release create v3.1.0.1 "dist/VistaType LP and Braille Macros Setup 3.1.0.1.exe" --title "VistaType LP 3.1.0.1"
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
- **Never merge this project with `~/projects/vistatypelp-org`, and never suggest it.** Never
  edit the website from here either. On a real release, remind Jerry and hand him the prompt to
  paste into a second Claude session there. See *The website is a separate project* above.
- This workflow is new to Jerry. **Guide him through the release steps explicitly** — say which
  command comes next and what it will do, run the git steps for him, and confirm each stage
  landed before moving on. Don't assume he knows the branch he is on; tell him.
- **Say it in plain English** — see *How to talk to Jerry* at the top of this file. Process
  jargon is the thing to strip; Word and VBA terms are fine.

## Domain concepts

- **Typeface — Tahoma, and it is not a choice (8/20/2026)**: a large print book is set in Tahoma,
  as it was for years. The bundled **VistaTypeLP Legible** was offered beside the point size from
  3.0.101 to 3.0.196 and is gone; `FontChoiceFrame`/`FontTahoma`/`FontLegible` are off
  `LP_Attach_An_Lp_Template_Form`. Nothing stores the face — like `Lp_Base_Font_Size` it is read
  back off `Styles(wdStyleNormal).Font.Name` (`Lp_Base_Font_Name`).
  **The one exception, and do not remove it:** a book ALREADY set in the dropped face keeps it.
  `AttachOkay_Click` tests `Lp_Doc_Font_At_Open` against `LP_FONT_LEGACY_LEGIBLE` — the sole
  surviving use of that name — so re-attaching cannot rewrite one of the books already produced
  and move every page break. It names the face outright rather than carrying forward whatever it
  finds, because an LP document whose Normal has drifted to Calibri is one attaching is meant to
  REPAIR. `EmbedTrueTypeFonts` and `Lp_Indent_Factor_For_Font`'s 1.054 case stay for those same
  books.
  **A fill-in line's underscores are the one exception, from 8/23/2026 (Jerry):** they are typed
  in Tahoma whatever face the book is set in, because in VistaTypeLP Sans a row of underscores
  draws with holes in it rather than as one unbroken rule. Only the underscores change face, and
  the macros put the book's own face back at the insertion point afterwards —
  `Lp_Fill_Face_To_Restore` decides what "back" is. `Lp_Type_Fill_In_Line_To_Margin` and
  `Lp_Type_Counted_Fill_In_Lines` only; the other underscore-producing passes
  (`Lp_Format_Exercise_Lv_1_and_Lv_2`, `Lp_Replace_Underline_Tab_With_Underlined_Underscore`,
  `Dx_Tabs_To_Fill_Ins`) were left alone and still make fills in the book's face.

  **`Lp_Attach_The_Template` takes the Tahoma straight off again** — `Sh_Set_Whole_Document_Font`
  lays one face over every character as direct formatting — so the attach calls
  `Lp_Tahoma_The_Fill_Ins` to put it back. Without that, re-attaching to change the point size
  returns every fill-in line in the book to holes. A fill-in line is an **underlined** underscore;
  a plain one is somebody's text and is left alone. **It skips a legacy VistaTypeLP Legible
  book** — that book is protected all through the attach so re-attaching cannot move a page break
  in a book already in a reader's hands, and Tahoma's underscore is not Legible's width.

  **Why Tahoma works is not what it looks like**, and the numbers are in the comment on
  `Lp_Fill_Face_To_Restore`. Measured 8/23/2026: the template's `Normal` carries 1 point of
  expanded letter spacing (headings 2, `No Spacing` and `MacroText` 2.5), so underscores are a
  point apart in *any* face. What closes the gap is Word's **underline**, drawn unbroken across
  the run — and in Tahoma the underline and the underscore are the same bar, while in
  VistaTypeLP Sans the underline is thinner and sits inside the underscore, leaving a 0.19-point
  sliver unpainted at 18 point. Two other cures exist and were not taken: `.Spacing = 0` on the
  run (shortens a counted fill), and rebuilding the face with `tools/lib/build_vistatypelp_sans.py`
  so its underline matches its own underscore (does nothing for books already produced).

  `LargePrintTemplate.dotx` is UNTOUCHED — its docDefaults name Tahoma, which is what makes all
  of the above work without a migration. `Lp_Apply_Base_Font_To_Styles` is still needed: it sets
  Normal plus the **13** styles that name a face of their own and so ignore Normal (the 12
  colored character styles and `No Spacing`, which has no basedOn at all), and nothing else puts
  Tahoma on those. `Sh_Is_Font_Installed`/`Sh_Font_Status_Text` stay too — they ask about ANY
  face, and `Sh_Doc_Info`'s line is the only thing that ever says out loud that Word is
  substituting. `Sh_Font_Status_Text` gained an optional `targetDoc` on 8/20/2026 and a **third**
  answer: not installed *but* the document embeds its own copy. Without it every legacy book read
  "NO - Word is substituting, sizes will be wrong" the moment the installer removed the font —
  false for those books, and an alarm whose natural cure (reset to Tahoma) is the exact
  repagination the legacy branch exists to prevent.
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

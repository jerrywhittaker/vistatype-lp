# Repo layout, and how the pieces fit together

Moved out of `CLAUDE.md` on 20 September 2026, which had grown to 1,382 lines and was being
read in full at the start of every session. Nothing here is new; it is the same text.

---

## The pieces and how they work together

The source of truth is `src/`. The items below are the *built* artifacts (and their
sources) that ship or that Word loads.

| Piece | What it is | Role |
|------|-----------|------|
| `LPandBRL.dotm` | Word add-in template (macro-enabled), built from `src/` | **The code.** The entire compiled VBA project — 236 subs/functions in `LPandBrlMacros`, plus 47 UserForms — and the embedded ribbon (below). Loaded from Word's `STARTUP` folder, so its macros are available to every document. Behavior is *authored* under `src/vba`/`src/forms`; this is where it *runs*. |
| Embedded ribbon (`src/ribbon/customUI14.xml`) | Ribbon customization XML, embedded into `LPandBRL.dotm` at build time | **The UI.** Defines the custom ribbon tabs **"VistaType LP"** (large print) and **"Braille Macros"** (DBT/BANA). Every button's `tag` names a VBA sub, dispatched through one `RibbonAction` handler. Because it is *embedded* (not the old global `Word.officeUI`), it **merges** with the user's ribbon instead of replacing it. |
| `LargePrintTemplate.dotx` | Word document template | **The style set.** The template *attached to a user's large-print document* (vs. `LPandBRL.dotm`, the global add-in loaded for every document). Supplies paragraph/character styles and page setup. The VBA references it by name in 7+ places, and treats a document as "large print" when this template is attached. |

Flow: User opens/creates a doc → `LPandBRL.dotm`'s Word **application events** fire (the
`VtEvents` sink, hooked by `AutoExec`) → detect whether the attached template is
`LargePrintTemplate.dotx` (large print), a BANA braille template, or neither → configure
Word accordingly → the embedded-ribbon buttons invoke VBA subs (via `RibbonAction`) to
clean up, format, and tag the document.


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
- **47 UserForms** — dialogs, prefixed by domain (see below). `Sh_Message_Form` is the shared
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
- `%AppData%\VistaType LP Settings\AutoCorrect-DEF.txt`, `-LP.txt`, `-BRL.txt` — **the three
  AutoCorrect tables, one per kind of document.** Also written by the add-in, never by the
  installer, and in the same folder the uninstaller leaves alone. **Word keeps ONE AutoCorrect list
  for the whole application** and nothing in VBA can make a second, so until 9/19/2026 the nineteen
  compact fractions large print and braille delete came off the transcriber's own letters too, and
  for good — reported that day by a transcriber on Office 365 who worked the mechanism out herself.
  Jerry's call: *"build a firewall between the autocorrect tables."* `Sh_AutoCorrect_Switch` swaps
  them at the top of the three configuration subs; `MS_Set_Word_Config_For_New_Install` deletes
  nothing now, and the two books still delete the nineteen, from their own table. Plain Unicode text,
  one entry per line, `name<TAB>value`, read and written with `Scripting.FileSystemObject` — **not**
  `System.PrivateProfileString`, because there are ~900 entries and one API call each would take
  minutes. **Rich text entries are never touched** (their content lives in `Normal.dotm` and cannot
  be read back as a string), so a formatted entry of hers stays shared between the three — the same
  reasoning that already keeps VistaType LP away from the four AutoCorrect exception lists. Measured
  on the build box 9/19/2026: 926 entries, 913 plain and 13 rich; a blind rebuild costs 4.6s, so the
  load applies only the differences and a switch costs about a second. Which table is loaded is
  **recorded** in `VistaType.ini`, never inferred from `Sh_ConfiguredAs`: Word writes its list to the
  `.acl` when it closes, so a session begun after quitting inside a book starts with that book's
  table already standing in Word.
  **Document Settings says which table is in use** — `Sh_AutoCorrect_Line`, on the line below the
  configuration. Every failure path in `Sh_AutoCorrect_Switch` is a silent exit, deliberately, so
  without that line a machine where the settings folder cannot be reached would go on stripping
  the transcriber's letters with nothing to say so. It reads `not recorded` there.
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
tools/windows/  Export-Vba.ps1 / Import-Vba.ps1 / New-UserForm.ps1 / Run-VbaTests.ps1 — run in
                Word on the build box
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
                Add-FormControl.ps1 — puts ONE control ON to an existing UserForm and rewrites the
                 .frm/.frx pair, headlessly. The counterpart to Remove-FormControls.ps1 and the only
                 way to write a control into a binary .frx from Linux. ADD THE CONTROL FIRST, THEN
                 THE CODE THAT NAMES IT - the opposite order from removal, and for the same reason:
                 code naming a control that is not there reaches the transcriber as run-time error
                 424 when the dialog opens. Applies Jerry's form conventions to what it adds
                 (Tahoma 10, a ControlTipText and the O/C accelerators on a button). Two traps it
                 shares with New-UserForm.ps1: the Designer does not expose the FORM's own Width and
                 Height to PowerShell at all - they live on the component, through
                 Properties.Item, and must be handed over as STRINGS - and a control's font can only
                 be set from VBA. Added 8/23/2026 for the KeyHelp lines on the $pg validation menus
                Move-FormControl.ps1 — moves and/or resizes controls ALREADY ON a form, headlessly,
                 and rewrites the .frm/.frx pair. The fourth form tool, added 9/6/2026 while putting a
                 spinner beside the progress bar: Add-FormControl could add it, and nothing could shift
                 the four controls already there. A control's POSITION and SIZE live in the binary .frx,
                 so they cannot be edited from Linux, and the .frm says nothing about where anything
                 sits. Every bound is optional - pass only what changes. -Controls takes several at once
                 as ONE comma-separated string, "Name:L,T,W,H;Name:...", with an empty slot meaning
                 "leave that bound alone" (one string, not an array: PowerShell's -File never parses an
                 argument as an expression - the same trap as Remove-FormControls.ps1). IT PRINTS THE
                 OLD BOUNDS BEFORE CHANGING THEM, which is the only record of where a control used to be
                 and the way back if the new position is wrong; and it REFUSES on a name that matches
                 nothing, printing every control with its bounds. Shares the other tools' two traps: the
                 form's own Width and Height live on the COMPONENT via Properties.Item and must be
                 handed over as STRINGS, and a control's FONT can only be set from VBA - set fonts in
                 the form's own code, as Sh_Convert_Progress_Form does for its spinner
                Set-FormCaption.ps1 — changes the CAPTION of one control on an existing UserForm and
                 rewrites the .frm/.frx pair, headlessly. A caption set in the designer lives in the
                 BINARY .frx, so it cannot be edited from Linux — this is the technique Import-Vba.ps1
                 uses to stamp the two About dialogs' version captions, made general. Matches on the
                 control's CURRENT CAPTION by default (-OldCaption), the way the version stamper does,
                 or on -Name when the words are not unique. REFUSES when it matches nothing, and prints
                 every caption on the form so the near-miss is obvious; refuses on more than one match
                 unless -All. Added 8/29/2026 to reword the Encode Fractions instruction on
                 Dx_Type_Dashes_Form. Afterwards run tools/lib/check_frm_eol.py and
                 tools/lib/trim_frm_blanks.py, as with the other two form tools
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
.githooks/      pre-push - refuses a push that would publish a secret. Version controlled
                so a fresh clone gets it; `make hooks` points git at this directory, since
                .git/hooks cannot be tracked. Two passes: gitleaks over the whole history
                (307 commits in ~1.6s), and patterns gitleaks cannot know about. The real
                exposure in this project is NOT a credential shape - it is the smart card
                PIN, which the planned unattended signing command puts on a command line
                where it reaches process listings, shell history and any build log. Also
                blocks a tracked .pfx/.p12/.pem/.key and build.config, which holds
                VT_API_KEY. SKIP_SECRETS_CHECK=1 overrides, for a verified false positive
                only. Tested both ways 9/20/2026: passes clean, blocks a staged PIN and a
                tracked .pfx.
tools/lib/      vm_up.py (starts the VirtualBox VM when the box is not answering - WSL can run
                Windows programs and the VM is on the same host, so VBoxManage.exe is reachable.
                Hooked into push-src, so every target that touches the box gets it. Does NOTHING
                when SSH already answers, and says so plainly when there is no VirtualBox here
                rather than failing. `make vm-up GUI=1` starts it WITH A WINDOW - headless is
                right for a build, a window is what anything that has to SEE Word needs);
                build_vba_test_bundle.py (gathers the VBA procedures tests/vba names, plus what they
                call and the declarations those need, into one importable module - `make vba-test`);
                check_style_guards.py (refuses to build when an ActiveDocument.Styles("name") lookup is
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
                build_ribbon_dispatch.py (`make build` runs it FIRST, because it WRITES
                src/vba/RibbonDispatch.bas - a generated Select Case calling all 47 ribbon macros
                by name, built from src/ribbon/customUI14.xml so the two can never drift. It exists
                because WORD DOES NOT PASS AN ERROR BACK OUT OF Application.Run: it takes the error
                itself and shows its own Run-time error dialog - the one offering Debug, which opens
                this source on the user's machine - and the caller's On Error handler is NEVER
                entered. Measured on the build box 8/26/2026: four real button presses wrote a marker
                set immediately before `Application.Run control.Tag` and never the one immediately
                after it, nor the one first thing in the handler. A DIRECT call propagates normally,
                which is the whole reason for the table. It refuses to generate if a button's macro
                stops being a plain no-argument Sub, since the Select Case would then not compile -
                and nothing here compiles VBA, so that would ship. NEVER hand-edit RibbonDispatch.bas);
                check_startup_unpulled.py (refuses to run `make try` when the add-in in Word's STARTUP
                folder on the build box is not the one this repo last built. That is what a UserForm
                LAYOUT edit looks like - Jerry makes those in the VBA editor against that one file,
                because a .frx cannot be written from Linux - and the copy at the end of `make try`
                goes straight over it: the build SUCCEEDS, says nothing, and the dialog quietly goes
                back to how it looked. Compares by CONTENT, not by date, because installing a real
                Setup.exe rewrites that file and its date with no edit involved and a date test would
                cry wolf every time. Runs BEFORE the bump so a stop costs no version number, and it
                also refuses while Word is open, which `make try` needs anyway - two minutes earlier
                than the check that used to catch it. `make installer` runs it --warn-only: it does
                not care whether Word is running, so a locked file is reported rather than fatal.
                ALLOW_STARTUP_OVERWRITE=1 goes ahead anyway. Added 8/26/2026 at Jerry's request);
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
                check_try_scope.py (what `make try` cannot test — prints, after a try, which
                changed files need a real installer instead. Ignores the version bump in
                installer/vistatype.iss, which `make try` itself makes: a warning that fires
                every time is a warning nobody reads);
                check_dotm_unsigned.py (refuses to build from, or promote, a .dotm whose VBA project
                is SIGNED. Wired into `push-src` (checks the repo-root base) and into `make deploy`
                (checks the dist/ file being promoted, which is the single place a signed one can
                get in). The root .dotm is the base every build starts from, so a signed file
                landing there is not a one-build problem. Measured 9/20/2026 by signing a real copy
                of the base with a throwaway certificate: the signature is NOT inside
                vbaProject.bin — that part came back byte-identical, all 302 OLE streams the same
                size — it is its own package part, word/vbaProjectSignature{,Agile,V3}.bin. The
                first version of this guard searched the OLE streams, the way the old binary .doc
                format stored it, and passed a genuinely signed file. And the consequence is NOT a
                stale signature: Word RE-SIGNS on import with whatever code-signing certificate the
                build user holds — one signature part went in, three came out, over completely
                re-imported code, and signtool's only complaint was the untrusted self-signed root.
                See docs/Code-Signing.md);
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
installer/      vistatype.iss is the product installer. From 8/24/2026 its [Code] carries a SECOND
                font sweep beside RemoveLegacyLegibleFont: RemoveSupersededSansFaces deletes any
                VistaTypeLPSans-*.ttf in the per-user Fonts folder that is not one of the four
                canonical filenames. A stray one still says "VistaTypeLP Sans" inside, so
                right-clicking it > Install repoints the family at it - measured on the build box,
                one held 3,094 characters against the shipping 6,450, three of fourteen math
                operators, no arrows, no italic. It runs at ssPostInstall ONLY, and the order is the
                thing that can go wrong: the four FontInstall entries must have rewritten the
                registrations to the canonical paths first, or deleting a stray the family currently
                points at would break the typeface. Unregister first, delete second; match registry
                values by DATA not name; the four filenames are the whitelist and the guard. NOT
                called at uninstall - the faces stay there by design. The hand-install path in the
                font .zip is the likeliest source of a stray; that mechanism was NOT reproduced
                (the shell Install verb does nothing over SSH with no desktop), so it is inferred. vistatype-font-only.iss beside it is a
                STANDALONE installer for the typeface and nothing else, built by hand with
                `make font-installer` when a tester needs the font without the add-in (added
                8/23/2026). It needs NO administrative rights - PrivilegesRequired=lowest, so
                {autofonts} is the user's own Fonts folder - and it has its own AppId, which must
                stay different from the product's or installing either would look like an upgrade
                of the other. Unlike the product installer, ITS uninstaller does take the font
                away. `make font-installer` also builds a .zip of the four faces, the three OFL
                texts and an instruction sheet (tools/lib/build_font_zip.py): an unsigned .exe can
                never promise not to be eaten by antivirus, and a font installed by right-click
                has no program in it to judge.
                Like src/ and tools/, this folder is WIPED on the build box before each copy, so a
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
docs/           Reported-Errors.md (the register of reported faults and what fixed them - READ IT
                BEFORE INVESTIGATING ANY REPORTED ERROR; see the rule below);
                Daily-Workflow-and-Releases.md (Jerry's plain-language guide to dev/main, building, releasing, and what to ask Claude); Installation-Guide.md (end-user install); Build-VM-Setup.md (the Windows build/test box; VirtualBox since 9/20/2026); Software-Agreement.md (GPLv3 About-dialog text); Code-Signing.md (why Defender deleted the unsigned Setup.exe on 8/13-14/2026, what in the installer scores against it, how to test with `make scan` and MpCmdRun, and the signing options — note Azure Artifact Signing does NOT sign VBA projects, and EV no longer skips SmartScreen. Jerry bought the Certum open source card on 8/17/2026; the last section is the step-by-step for the day it arrives, what the build box already has, and the five things the card-free rehearsal proved); VistaTypeLP-Sans.md (what the bundled typeface covers and does not - languages, mathematics, science, medicine - measured face by face against Tahoma, plus the width and pagination trade-off. The Insert Symbol subset-list defect is FIXED (9/6/2026): Word builds that list from the OS/2 unicode-range flags, not from the character map, and 16 were unset - 11 are now claimed and 5 REFUSED on purpose, Arabic among them because the font has 62 of its characters and no arab shaping. Browsable went 45.7% -> 88.2%; the 695 characters still out of reach sit at code points no OS/2 block covers, so no flag can reach them. A block that is neither claimed nor refused now STOPS the font build)
reference/      generated read aids (gitignored mirror + interim form-code dump)
tests/          the Python test suite over tools/lib - the build's own checkers and the files it
                generates into a transcriber's Word.officeUI. Run with `make test`; needs pytest
                (`sudo apt install python3-pytest`), takes about half a second, and never opens
                Word or reaches the build box. Added 9/20/2026, because until then the eight
                checks that gate every build were themselves untested. Each test file ends with a
                canary over the REAL source, so a failure there means `make build` is already
                refusing. Two rules for anything added: a test must never write into the
                repository (build_ribbon_tabs.py appends to the shipped-ids ledger even with
                --check-only, so canaries copy first), and fixtures are CRLF, because
                check_vba_structure.py splits on CRLF and a bare-LF .frm is itself a defect.
                tests/README.md has the table of what each file protects.
                WIRED INTO THE BUILD 9/20/2026 (Jerry): `make build` runs `make test` first, so
                every try and every installer does too, and a failure stops it before the box is
                touched. `make installer` also runs `make vba-test`. That one is on the installer
                and not on every build because it starts Word and `make try` is the fast loop -
                say so if that trade should change.
tests/vba/      the VBA unit tests - `make vba-test`. They do NOT reference the built add-in:
                tools/lib/build_vba_test_bundle.py lifts the procedures each test names out of
                src/vba as they stand, follows what they call, and the runner imports that into
                a throwaway invisible document on the build box. So they test the SOURCE, and
                they can reach the Private helpers, which a referencing project could not call.
                lib/ is vba-test (Tim Hall, MIT), with TestCase.cls modified in one marked place
                to late-bind its Dictionary - early binding hangs, see tests/vba/README.md -
                plus VtFileReporter.cls, ours, because vba-test reports to the Immediate Window
                and nothing can read that out of a Word started over SSH. Rubberduck was
                considered and cannot be used: archived March 2026, and it only ever ran tests
                from its own window inside the VBA editor.
                ONLY PURE HELPERS belong here. Anything that reaches a Document, a UserForm or a
                MsgBox hangs an invisible Word; the table of what cannot be tested this way, all
                of it measured, is in tests/vba/README.md.
assets/fonts/vistatypelp-sans/   the bundled typeface, VistaTypeLP Sans — four tracked .ttf faces
                plus the THREE OFL texts it needs (Noto Sans, Noto Sans Math, Noto Sans Symbols).
                Built by `make fonts` from the current Noto releases; the built faces are what
                ship, so commit them when they change. Coverage, and what it will NOT set, is
                docs/VistaTypeLP-Sans.md.
                (This folder held a DIFFERENT face until 8/20/2026: VistaTypeLP Legible, shipped
                from 3.0.101 to 3.0.196 and dropped because it has no Greek, no IPA and almost no
                mathematics, and Word substitutes a missing character silently and at the wrong
                size. VistaTypeLP Sans replaced it on 8/22/2026 and was checked for exactly that
                first. tools/lib/rescale_font.py and tools/windows/Check-Font.ps1 went with the old
                face and did not come back. The pristine Atkinson Hyperlegible download and the
                licensing write-up are in git history; the Braille Institute's license PDF is still
                in ~/reference/vistatype-lp/ and must never be committed.)
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
Makefile        pull / build / build-dispatch / ribbon / qat / read / fonts / try / test /
                vba-test / vm-up / deploy /
                branding / stage /
                installer / font-installer / scan
                (`make fonts` rebuilds VistaTypeLP Sans from the current Noto Sans, Noto Sans Math
                 and Noto Sans Symbols releases. It REACHES THE NETWORK, so it is deliberately not
                 part of `make installer`; twelve checks run at the end and it refuses to claim
                 success if any fail. The version that went with the dropped typeface on 8/20/2026
                 came back on 8/22/2026 building the new one)
                (`make try` bumps, builds, and puts the new .dotm straight into Word's STARTUP
                 folder on the build box - no installer, no wizard. Added 8/23/2026 at Jerry's
                 request to shorten the code-test loop; see *Testing a change* below. It stops
                 first on an unpulled form edit - tools/lib/check_startup_unpulled.py above)
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
- **To take a button OFF the ribbon, move it to the hidden `tab_VT_Retired_Buttons`** (from 9/15/2026,
  Compress Linear Math first) rather than deleting its id. **Never park it on `tab_LP_and_BRL_QAT_Icons`**:
  `build_qat.py` makes the icons-only toolbar out of that tab, and the installer's default `Mine` mode
  would add the retired button to every transcriber's toolbar. `build_ribbon_tabs.py` skips the retired
  tab as it skips the toolbar one.


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

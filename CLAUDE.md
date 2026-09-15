# VistaType LP

A Microsoft Word add-in that helps transcribers produce **large-print** documents for
visually impaired readers, and **braille** source files for the Duxbury Braille Translator
(DBT) via its BANA template. Authored by Jerry Whittaker (jerry@vistatypelp.org),
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

**FINISHED 9/4/2026 (3.0.366). No macro in VistaType LP puts a document on the screen any more.**
Change Picture Color was the last one, and `Lp_Copy_To_Temp_Doc` was removed with it — along with
`Sh_Is_End_Paragraph_Mark_Included`, whose only caller it was, and the two public variables only
that macro wrote. Dialogs 294, 301 and 305 are retired. **Five macros still make a scratch
document and all five keep it hidden**: `Lp_Table_Convert_Hidden`, `Lp_Exercise_Levels_Hidden`,
`Dx_Exercise_Levels_Hidden`, `Lp_Horz_To_Vert_Hidden` and — from 3.0.402 —
`Lp_TOC_CleanAndFormat_TOC`.

The rest of this section is kept because the reasoning is what to copy the next time a macro has
to move onto a range — the traps are not specific to the scratch document.

`Lp_Copy_To_Temp_Doc` created the scratch document hidden and then deliberately showed, maximized
and activated it — four statements, not an accident. That was the screen flashing. It had to,
because its callers worked through `Selection`, and `Selection` only reaches the active document.
Converting a macro means moving its passes onto a **range**, so the window is never needed. The
round trip itself stays where it earns its place: it is what gives the single undo and the
private workspace.

**The braille half of this is finished** (3.0.318, 8/31/2026). `Dx_Copy_To_Temp_Doc`,
`Dx_Copy_From_Temp_Doc` and `Dx_Attach_Same_BANA_Template` were removed once Exercise Levels
1 & 2 — their last caller — moved onto `Dx_Exercise_Levels_Hidden` at 3.0.309, and the
`Sh_Copy_To_Temp_Doc` / `Sh_Copy_From_Temp_Doc` route pickers went with them. No braille macro
shows a scratch document any more; three still make one, hidden.
`Lp_TOC_CleanAndFormat_TOC` came off it at 3.0.338 — and **went back on, hidden, at 3.0.402** (9/7/2026), which is the second time this project has learned the same lesson. Doing its passes in the book cost **39 Ctrl+Z presses**, measured on Jerry's own file; Jerry, 9/7/2026: *"multiple (i mean more then 3 or 4) is not an acceptable undo requirment."* The passes stayed on ranges, so the document is made hidden and never shown, and **one** assignment comes home — one press, not the two Convert Table to List needs, because the landing is paragraphs rather than a table.

**The large-print half finished at 3.0.366** (9/4/2026) with Change Picture Color, and that was a
deletion too — the fourth. Recoloring pictures is "walk some images and set one property", so it
needs a `Range`, not a document. Five faults went with the round trip and none had ever been
reported: the transcriber's clipboard, emptied by the `Copy` and `Paste` that carried the text
home; cells the user had not selected, because `Lp_Copy_To_Temp_Doc` widened any selection made
inside a table to the whole table; the selection replaced by a paste rather than edited where it
stood; the demand that the selection end with a paragraph mark, which existed only to make the
text travel whole; and, on a machine without `LargePrintTemplate.dotx`, **the user's own book
closed without saving** — the missing-template `Exit Sub` that `Application.Run` never hands
back, followed by `Sh_Is_End_Paragraph_Mark_Included` doing
`ActiveDocument.Close SaveChanges:=False` on that book. **Equations are out of reach now, and that is Jerry's call** — *"I don't want equations
(from MathType) touched."* A MathType equation is an embedded OLE object and so an inline shape
like any other, and all three of the old loops tested nothing; the recolor takes
`wdInlineShapePicture` and `wdInlineShapeLinkedPicture` only. What Word does when asked to gray an
embedded object could **not** be measured — embedding one into an invisible Word fails outright
over SSH, the same family as `Tables.Add` hanging — and the type test means it never has to be
answered.

A **sixth** thing was about forms rather than scratch documents: the Okay button read its radio
buttons *after* `Unload Me`, and touching any member of a UserForm's default instance is what
creates the form, so those lines might have been reading a brand new form's design-time values.
**They were not** — put to Jerry, 9/4/2026: choosing grayscale has always given him grayscale.
Read the choices before unloading anyway; it is the right order, not a repair. **Ask him before
writing a fault into the record**: a UserForm cannot be exercised headlessly, so the only
measurement available was his own Word, and one sentence from him settled it.

**Table Tools came off it at 3.0.345** (9/3/2026), reported by Jerry testing 3.0.339: flashes of
the table's yellow rows during Convert to List and Rotate Table, and a full-screen blink at the
end of a rotation. The blink was deliberate — the rotation finished by minimizing and maximizing
the Word window "to force Windows OS to repaint the pixels". Ten macros now take the table they
are to work on (`Lp_Table_Cleanup_For_Roation_And_List`, `Lp_Table_Transpose_Table` — now a
Function returning the new table — `Lp_Table_Fill_Empty_Cells`, `Lp_Table_Row_Column_Header_Setup`,
the two `Lp_Table_Apply_Character_Case_To_*_Headers`, `Lp_Table_Style_InCell_Para_And_Image`,
`Lp_Table_Mark_Keep_With_Next`, `Lp_Table_Is_R1C1_Empty` and the three `Lp_Table_Convert_*_To_List`),
and `Lp_Table_Note_Above` opens the transcriber note in a new paragraph above the result instead of
at `Selection.HomeKey wdStory`. The clipboard round trip went with it, so the transcriber's own
clipboard is no longer emptied.

**And then the round trip came back, hidden — which is the lesson.** 3.0.345 removed the scratch
document altogether and did the work in the transcriber's book. That made Ctrl+Z **30 to 50
presses** (Jerry: *"not tollerable for anyone"*), and a custom undo record to cure that
**crashed Word outright** at 3.0.345 and again at 3.0.347 — the defect already recorded for
Format TOC: **a `Find` with `Replace:=wdReplaceAll` inside an open `StartCustomRecord` kills
Word**, access violation in `wwlib.dll`, nothing raised, nothing logged.

**The scratch document does two jobs and only one of them was ever the fault.** It had to be
*shown* because the passes reached their table through `Selection` — that is the flashing, and
converting the passes onto ranges is what fixes it. But it is also **what keeps the undo short**:
edits made in another document are not in this book's undo stack at all, so only the one
assignment home is. That is what the rule at the top of this section means by *"the round trip
itself stays"*. `Lp_Table_Convert_Hidden` (3.0.348) is the shape to copy — hidden
`Documents.Add`, passes on ranges, result home by `FormattedText`, no undo record. **Two presses,
not one, and that is as good as it gets when the original is a table:** assigning `FormattedText`
to a range that *is* a table fills that table's cells instead of replacing it (3.0.349 shipped
the list back inside its own table), so the table has to be deleted first — one step for the
delete, one for the result. Jerry's own words when it went wrong: *"why not save the original table and restore it
with one undo?"*

Two things follow. **Convert the passes, do not delete the round trip** unless the macro provably
needs no private workspace (Resize Images, Replace Section Breaks and Format TOC did not).
And **read `docs/Reported-Errors.md` for the mechanism you are about to introduce, not only for
the feature you were asked about** — the undo-record row was already there and was walked past.
**Three faults could not have survived the move**, and finding them is the argument for reading
what a macro does rather than only how it gets there: two `Find` passes used `wdFindContinue`,
which against a real book is every colon in the chapter rather than one table;
`Lp_Table_Mark_Keep_With_Next` read a table variable the half above it had never set (run-time
error 91 on any row-and-column list run without *keep list groups on same page*); and its
paragraph-mark tidy-up began at `Selection.HomeKey wdStory`, the top of the book.

**Check first whether the round trip is needed at all — and what the macro actually does.** Of the
three settled on 9/1/2026 only one was a conversion. `Lp_Resize_Images` (3.0.321) only walks
images and sets their scale, and an `InlineShapes` collection comes off a `Range` as readily as
off a document. `Lp_TOC_CleanAndFormat_TOC` (3.0.338) was the third deletion **and was reversed at 3.0.402 — the
passes were right to convert, the round trip was wrong to delete**: every pass is Paragraphs or
Characters off a Range, it never used the clipboard, and it turned out that on a machine which
cannot find `LargePrintTemplate.dotx` the round trip ran all nineteen of its replaces across the
transcriber's whole book and then closed it without saving — because `Lp_Copy_To_Temp_Doc`
reports a missing template with a `MsgBox` and a plain `Exit Sub`, and `Application.Run` never
hands that back to the caller. **Every `Find` in a converted macro must be `wdFindStop`**; the
scratch-document versions use `wdFindContinue`, which on a range means the whole book.
`Lp_Section_Brk_Caution` (3.0.326) turned out to be **corrupting books**: replacing
every section break with a page break collapses the document to one section and discards every
section's page setup, which had been true since long before the conversion work. It was cured by
changing each section's `SectionStart` property instead of deleting anything. Reading what a macro
does to the scratch document, rather than only how it gets there, is what found that — see
`docs/Reported-Errors.md`.

`Lp_Format_Exercise_Lv_1_and_Lv_2` was the sixth and was converted at 3.0.319 (9/1/2026), onto
`Lp_Exercise_Levels_Hidden` / `Lp_Ex_Passes` / `Lp_Ex_Repl`. **It was NOT merged with the braille
twin**, and that is worth recording because the merge rule below made it look likely: the two
apply different styles and produce different output — `List` and `List 2` with underlined Tahoma
fill-in lines here, `Exercise1`/`Exercise2`/`Ex2Nemeth2` with `[[*kps*]]` and `[[*kpe*]]` markers
there. `Lp_Fix_Para_Space_Errors` and `Sh_Remove_Spaces_Before_Punctuation` took an optional range
so they could be run against the hidden document.

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
                Daily-Workflow-and-Releases.md (Jerry's plain-language guide to dev/master, building, releasing, and what to ask Claude); Installation-Guide.md (end-user install); Build-VM-Setup.md (Hyper-V build/test box); Software-Agreement.md (GPLv3 About-dialog text); Code-Signing.md (why Defender deleted the unsigned Setup.exe on 8/13-14/2026, what in the installer scores against it, how to test with `make scan` and MpCmdRun, and the signing options — note Azure Artifact Signing does NOT sign VBA projects, and EV no longer skips SmartScreen. Jerry bought the Certum open source card on 8/17/2026; the last section is the step-by-step for the day it arrives, what the build box already has, and the five things the card-free rehearsal proved); VistaTypeLP-Sans.md (what the bundled typeface covers and does not - languages, mathematics, science, medicine - measured face by face against Tahoma, plus the width and pagination trade-off. The Insert Symbol subset-list defect is FIXED (9/6/2026): Word builds that list from the OS/2 unicode-range flags, not from the character map, and 16 were unset - 11 are now claimed and 5 REFUSED on purpose, Arabic among them because the font has 62 of its characters and no arab shaping. Browsable went 45.7% -> 88.2%; the 695 characters still out of reach sit at code points no OS/2 block covers, so no flag can reach them. A block that is neither claimed nor refused now STOPS the font build)
reference/      generated read aids (gitignored mirror + interim form-code dump)
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
Makefile        pull / build / build-dispatch / ribbon / qat / read / fonts / try / deploy /
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

**These are requirements about how a message looks and reads. They are not a ban on anything** —
Jerry's correction, 8/26/2026, after this file's earlier wording was quoted back at him as a
reason his own `MsgBox` suggestion could not be done. He set how a message should appear; he did
not forbid a control or a mechanism.

What follows from that is a fact about `MsgBox`, not a rule of his: VBA's `MsgBox` draws in
whichever font Windows uses for message boxes, its button says `OK` and cannot be changed, and it
has no accelerators — so a plain `MsgBox` cannot meet the three. A UserForm can, which is why
messages here are UserForms. Where a `MsgBox` is the right answer anyway, say so and let Jerry
decide; do not cite the rule as settling it. `Sh_Message_Form` is the one dialog every message goes through, reached by two
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

**If the form refuses to appear, `Sh_Say` and `Sh_Ask` fall back to a `MsgBox`** (8/26/2026).
Until then the trap around `.Show` swallowed the failure and the message was simply **lost** —
intolerable in an error handler, which is where it came up. Smaller and saying `OK` beats never
appearing. In `Sh_Ask` it changes an answer too: a question that could not be shown used to come
back `False`, i.e. Cancel, and silence that quietly means "no" cannot be told apart from the user
having chosen "no".

**Not retrofitted.** The two newest toolbar buttons — Reset Word Configuration and Styles Pane:
Recommended — were converted on 8/23/2026, five messages. **The attach-a-large-print-template path
followed on 9/6/2026, sixteen messages** — the obsolete-template warning (123), the missing-template
stop (141), the cancelled-save question (306), the could-not-reopen notice (233), the
stabilized-and-saved notice (201), and the eleven page-size and margin checks on
`LP_Attach_An_Lp_Template_Form` (189–199). Every one kept the number it already had.
They convert a feature at a time. Nothing new uses `MsgBox`.

**What is left, recounted 9/6/2026: 65 in `LPandBrlMacros`, 30 across the forms, 18 in the three
smaller modules — 113 in all.** **Count with comment lines excluded**, or the number comes out
around 40 too high: the changelog and the version blocks say the word `MsgBox` constantly, and so
do commented-out calls. `grep -ah MsgBox <files> | sed 's/^[[:space:]]*//' | grep -v "^'"` is the
count that means something. The figures this paragraph carried until 9/6/2026 (85 / 48 / 20) were
made the unfiltered way on 8/23/2026 and were never right.

**Three things a conversion cannot carry, and they decide the ORDER of the remaining work:**
- **The icon and the sound.** Around 45 of the calls pass `vbExclamation`, `vbInformation`,
  `vbCritical` or `vbQuestion`, and the warning ones make a noise. A UserForm draws neither and
  plays neither. Weigh that per message rather than sweeping.
- **Yes/No has to be reworded.** `Sh_Ask` offers Okay and Cancel. Sixteen calls are `vbYesNo`;
  four are already `vbOKCancel` and are free. **None is three-button**, which is why this is
  possible at all.
- **`Sh_Message_Form` fixes which key presses which button, and it cannot be told otherwise.**
  `New-UserForm.ps1` built it with `Okay.Default = True` and `Cancel.Cancel = True`, so **Enter is
  always Okay and Esc is always Cancel**. Six `MsgBox` calls pass `vbDefaultButton2` — they are
  saying the SECOND button is the safe one, which this form cannot express. Word such a question
  so that Okay is the safe answer, or leave it a `MsgBox`.
- **A `vbYesNo` gains an escape hatch the moment it moves here.** A Yes/No `MsgBox` disables its own
  X and ignores Esc: the user has to answer it. `Sh_Message_Form` takes both as Cancel. On dialog
  306 that means Esc now carries on **without saving**, silently, where before it did nothing.
  Before converting any other Yes/No, ask what Cancel costs if it is pressed by reflex.

**Twenty of the calls read an answer; the rest only tell the user something.** The tell-only ones
convert almost mechanically. Do not attempt all of them in one sweep: nothing here compiles VBA,
there are no automated tests, and a scripted edit across this module has already deleted code out
of both book configurations with every guard passing and a green build.

**Hover text cannot be made 10 point Tahoma, on a ribbon button or on a form.** Word draws a
ribbon/QAT screentip and supertip itself in the Office UI font; customUI has no font attribute
of any kind, and a `getSupertip` callback supplies the words and nothing else. `ControlTipText`
on a UserForm control is a Windows tooltip and is the same. The only thing that changes either
is the reader's own Windows text size (Settings → Accessibility → Text size), which changes it
everywhere. Where hover text needs to be readable at 10 point or more, the answer is to put the
words in the dialog, which VistaType LP does control.

### How VistaType LP shows that it is working

Jerry's rule, 9/6/2026: **one progress indicator.** His words — *"the current state of progress
indicators makes VistaType LP look schizophrenic... I like to have just a progress bar but i do
like the text which indicates what is happening."*

There were four of VistaType LP's own: two near-identical spinner boxes
(`Sh_NonModalMessageForm`, `Sh_Please_Wait_Form`), the bar (`Sh_Convert_Progress_Form`), and a bar
drawn out of pipe characters in Word's status line by the two Export Selection macros. **Every
feature is on the bar now and both spinner boxes have no callers left** — they are still in the
project, unused, because deleting a UserForm is a deliberate job where the code naming it has to
go first.

**The shape.** A fill that only ever goes forwards, a percentage, a line of text naming the step,
and a spinner. Reached through four helpers in `ShNonModalMessage`, never by naming the form:

```vba
Sh_Progress_Open "Fixing common file errors"   ' modeless; starts the spinner
Sh_Progress_Say 40, "Removing square bullets"  ' moves it and says what is happening
Sh_Progress_Hide / Sh_Progress_Show            ' step aside for a modal dialog of Word's own
Sh_Progress_Close                              ' every path out, including the error handler
```

- **Never touch `Sh_Convert_Progress_Form` directly.** Every helper is gated on the module flag
  `Sh_Progress_Up`, never on the form's `.Visible`, because reading any property of an unloaded
  UserForm instantiates it and runs its `Initialize` — and this form's `Initialize` reads
  `Application.Left`, which hangs a Word driven over SSH with no desktop. That gating is also what
  makes a `Say` with no bar open cost nothing, so a macro can carry these calls and still be run
  from somewhere that shows no indicator.
- **The spinner is for the steps where nothing can move.** `Repaginate`, `InsertFile` and the Save
  As dialog are single calls into Word, and VBA is single-threaded. The fill says how far along;
  the spinner says still alive. Before it existed the only answer was a message apologizing for a
  frozen bar.
- **`Sh_Progress_Span lo, hi` lends a slice to a sequence that counts its own work 0–100**, so one
  counted job can run inside another. The attach hands 3–35 to File Cleanup and 74–90 to Normalize
  Styles. **Give the bar back** (`Sh_Progress_Span 0, 100`) the moment the inner sequence returns,
  or every later `Say` lands inside the slice and the bar stops part way along.
- **Count with a step helper, one per sequence** — `Dx_Ffc_Step`, `Lp_Ffc_Step`, `Lp_Ns_Step`. The
  total is **one more than the number of passes**, because each step announces itself *before* its
  pass runs and a bar must never read 100% while work is still going on. Where stages differ
  enormously in cost — the attach, the DAISY converter — hand-pick the percentages instead of
  counting equal steps. Jerry's choice, 9/6/2026: rough stages that always advance, never a bar
  that sits frozen looking like a hang.
- **Close it on every path**, including the error handler, and **take a copy of `Err.Number` and
  `Err.Description` first** — `Sh_Progress_Close` runs `Err.Clear` inside itself, so reporting
  after it reports 0.
- **Word's own indicator is not ours and cannot be removed.** Word draws a message and a bar in
  the status line while it saves, and on OneDrive while it uploads. VistaType LP writes nothing to
  `Application.StatusBar` any more. **Do not hide our bar to avoid the overlap** — that was tried
  at 3.0.384 and reversed at 3.0.385: an export to a LOCAL disk gets no indicator from Word at
  all, so hiding ours takes the indicator away from the only case that has nothing else, and on
  OneDrive it turned one steady box into three events. Jerry: *"we need to see the bar if the
  export is to a local disk."*

### How VistaType LP reports a macro that failed

Added 8/26/2026, at Jerry's request. Before it, of 175 `On Error` statements in the project,
**four** ever put an error in front of the user; the rest either swallowed it or let it reach
the user as **Word's own "Run-time error" dialog**, which offers **Debug** — and the VBA project
is not locked for viewing, so Debug opens the source on their machine.

- **`RibbonAction` is the single catch point** for all 47 ribbon and toolbar buttons. It calls
  each macro **directly**, through the generated `Sh_Dispatch` — *not* `Application.Run`, which
  never lets the error back out (see `build_ribbon_dispatch.py` above). `Application.Run` remains
  only as the fallback for a name the table lacks.
- **`Sh_Report_Error`** restores the screen first (`ScreenUpdating`, `DisplayAlerts`, the progress
  box and the please-wait box), then appends one line to the log, then shows dialog **240**
  through `Sh_Ask` — where Okay opens the folder holding the log. Nothing in it may raise: it
  runs on a path that has already gone wrong, and a second error there is Word's dialog again.
- **The log is `%AppData%\VistaType LP Settings\VistaType-Errors.log`**, beside `VistaType.ini`
  in the folder the uninstaller deliberately leaves alone — a log thrown away by the reinstall
  someone suggested to cure the fault is worth nothing. `Sh_Store_Folder` is shared by both.
- **`Sh_Last_Activity`** carries the "Where:" line. VBA gives a handler a number and a description
  and nothing else — no line number (`Erl` needs numbered lines, which this project does not use)
  and no call stack — so the last `SetActivityMessage` string is the only position marker there is.
  The form's `SetActivityMessage` records it as well as showing it.
- **Not total coverage.** The nine keyboard shortcuts in `src/keymap` and every UserForm button
  call their macros directly and are still unguarded.
- **`VT_VERSION` says which build produced a fault.** A placeholder in `src/vba` forever
  (`"unstamped"`); `Import-Vba.ps1` replaces it inside the **built** `.dotm`, the same pass that
  stamps the two About captions. Never put a real number in the source: the version bumps on
  every `make try` and it would be committed by accident. A log line reading `unstamped` means
  the build did not stamp it.
- **The log is newest-first and capped at 50 lines** (Jerry, 8/26/2026): the line anyone needs is
  the one they just caused. `Sh_Log_Newest_First` rewrites the whole file — Append can only add
  at the bottom. **It is not cleared when the user presses Cancel**, and that was considered and
  rejected: Cancel means "do not open the folder", not "throw this away", and the user who
  cancels is exactly the one whose repeats are the only record you would ever get.

### Before investigating a reported error, read `docs/Reported-Errors.md`

**Jerry's rule, 8/26/2026** — his words: *"I don't want to spend time and tokens on
re-investigating errors that have already been fixed."*

`docs/Reported-Errors.md` is the register of faults and what fixed them, keyed on the four things
a log line carries: **version, error number, macro, step**. Check it *first*, every time, before
reading any code. It records the one thing the repository cannot infer — that *this* reported
fault was cured by *that* change.

The repo is only a backstop, and a weak one: the per-sub `' Version:` blocks say what changed,
not which reported fault it cured; the macro in a log line is the *button's*, not necessarily
where the failure was, because VBA gives no call stack; and only `X.Y` releases are tagged, so a
report from an internal `3.0.X` build has nothing to diff against.

**Add the row in the same change as the fix, never afterwards** — a register that lags says "not
fixed" about something that is. Fill **Fixed in** with the build and **Shipped in** with the
release; they are different, and only the second is something a user can install.

### Nothing here compiles VBA, and a build-time compile gate is NOT possible

Established by measurement on 8/26/2026, after a Variant passed to a ByRef `String` shipped in
3.0.274 and reached Jerry as **"Compile error in hidden module"**. Every guard in `tools/lib`
passed it — structure, line length, style guards, form calls — because none of them type-check.

**Do not try to add a compile gate to the build. It cannot work.** Three things were tried and
measured, so nobody repeats them:

- **VBA compiles lazily, one procedure at a time.** A module holding a good `Harmless` sub and a
  broken one: running `Harmless` succeeded while the broken neighbour sat there uncompiled. So a
  probe macro compiles only itself and what it actually calls.
- **A probe that calls into the project does not help.** Versions calling nothing, calling
  `Sh_Dispatch` (which names all 47 ribbon macros), and calling into `LPandBrlMacros` itself all
  reported "compiles clean" with the fault deliberately reintroduced. A gate that always passes
  is worse than none, which is why it was removed rather than kept.
- **`Debug > Compile` is unreachable through automation.** The VBE's command bars enumerate no
  matching control, even after `VBE.MainWindow.Visible = True`.

**So the compile step is a HUMAN one**, and it is step 3a of the release checklist above. Say so
whenever a change touches VBA and matters.

The narrower lesson worth keeping: **give every helper `ByVal` parameters.** The fault was a
Variant handed to a ByRef `String`, which VBA rejects at compile time. `ByVal` also avoids the
other half of that bug — `Sh_IsValidRomanNumeral` reassigned its own parameter, so under ByRef it
was quietly trimming and upper-casing the *caller's* variable. The 39 implicit-ByRef parameters
left in the project are mostly UserForm event handlers (`Cancel As Integer`, `CloseMode As
Integer`), which must stay ByRef.

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

### Testing a change — `make try` first, an installer when it matters

Jerry's loop was: code, build the installer, run the wizard, test, repeat. From 8/23/2026 the
short version is **`make try`** — it bumps, builds, and drops the new `LPandBRL.dotm` straight
into Word's STARTUP folder on the build box. He closes Word, runs it, opens Word, tests. The
`.dotm` is the *same file the installer packages*, so this is not a lesser test of the code.

**It bumps, and that is not tidiness.** Every third number has meant "a build you can tell apart
in the About box" since three builds shared `3.0.6` on 7/26/2026. Swapping `.dotm` files without
bumping brings that straight back. So a `make try` takes a number too — that number simply never
gets a `Setup.exe`, which is what a private build counter is for.

**Word must be closed on the build box.** `make try` checks by process name and stops if it is
not — a stray automation Word from a headless test run counts and is invisible on the desktop.
That check now runs **before** the bump and the build (`tools/lib/check_startup_unpulled.py`), so
an open Word costs a second rather than two minutes and a version number.

**It also refuses to overwrite a form edit that has not been pulled into `src/` yet.** Jerry
edits a UserForm's LAYOUT in the VBA editor against the add-in in Word's STARTUP folder on the
box — a `.frx` is binary and cannot be written from Linux — and that edit lives in that ONE file
until it is exported back into `src/forms/`. The copy at the end of `make try` went straight over
it: the build succeeded, nothing was reported, and the dialog quietly went back to how it looked
before. Nothing detected that until 8/26/2026, when Jerry asked what stops it. The guard compares
by **content**, not by date, so installing a real `Setup.exe` does not read as an edit.
`ALLOW_STARTUP_OVERWRITE=1 make try` goes ahead anyway. **Claude's duty when it fires:** pull the
form out of the box's copy first — never `make pull`, which wipes `src/` from the repo-root
`.dotm`; see `docs`/memory on the export-and-cherry-pick round trip.

**What `make try` cannot test, and these need `make installer`:**

| Change | Why the `.dotm` does not carry it |
|---|---|
| `installer/**` | the toolbar, trusted locations, the license file, the artwork, uninstalling |
| `src/ribbon/**` | the tabs *are* embedded, but on a machine the installer has set up, the copy in the user's own `Word.officeUI` is what shows and the embedded one is hidden (`VtTabVisible`) — so a tab change can look like it did nothing |
| `src/keymap/**` | injected into `LargePrintTemplate.dotx`, which only the installer puts in place |
| `*.dotx` | the large-print template: styles and page setup |
| `assets/fonts/**` | installed and registered by the installer |

`tools/lib/check_try_scope.py` prints exactly this after every `make try`, naming the changed
files that need a real installer. **Claude's standing duty (Jerry, 8/23/2026): say so in the
reply too.** He should never have to work out for himself that what he is about to test cannot
be tested this way — and before anything is called done, and always before a release, he installs
a real `Setup.exe`.

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
3a. **DEBUG → COMPILE, BY HAND, ON THE BUILD BOX.** Open the STARTUP `.dotm` *directly*
   (right-click → Open in `%AppData%\Microsoft\Word\STARTUP`, not the loaded add-in), then
   **Alt+F11 → Debug → Compile LPandBRL**. Nothing happens if it is clean; it stops on the
   offending line if it is not. **This is the only full compile that exists** — see the section
   below — and it is the last chance to catch a "Compile error in hidden module" before a
   transcriber does. Ten seconds; it caught nothing on 8/26/2026 only because the fault had
   already been found the expensive way.
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

- **Typeface — VistaTypeLP Sans or Tahoma, and it IS a choice again (8/22/2026, 3.0.224)**: the
  choice is `FontChoiceFrame` on `LP_Attach_An_Lp_Template_Form`, carrying `FontSans` and
  `FontTahoma`. Two constants name the faces: `LP_FONT_SANS` ("VistaTypeLP Sans") and
  `LP_FONT_TAHOMA` ("Tahoma").
  `UserForm_Initialize` decides what is offered, in this order:
    - **VistaTypeLP Sans is the default** when `Sh_Is_Font_Installed(LP_FONT_SANS)` says it is
      there. When it is not, `FontSans` is DISABLED, Tahoma is selected, and the button's own
      caption gains "-- NOT INSTALLED (or Word needs restarting)". The words go on the control
      because hover text cannot be made 10 point — see *How VistaType LP says things*.
    - **An existing LP book's own face wins over the default.** `Lp_Doc_Font_At_Open` is read off
      `Styles(wdStyleNormal).Font.Name`, and Tahoma or Sans selects its own button.
  Nothing stores the face — like `Lp_Base_Font_Size` it is read back off
  `Styles(wdStyleNormal).Font.Name` (`Lp_Base_Font_Name`).
  The face is named outright rather than carried forward, because an LP document whose Normal has
  drifted to Calibri is one attaching is meant to REPAIR.
  **The legacy VistaTypeLP Legible protection is GONE (9/4/2026, 3.0.366+), and this replaces the
  rule that used to say it must never be removed.** Five pieces of code recognized a book set in
  that dropped face and shielded it from a re-attach — the constant, two tests on the form (gray
  both buttons, caption the frame "this book keeps VistaTypeLP Legible"), the 1.054 case in
  `Lp_Indent_Factor_For_Font`, and the guard in `Lp_Attach_The_Template` skipping
  `Lp_Tahoma_The_Fill_Ins`. **All of it protected a book that does not exist.** Jerry, 9/4/2026:
  *"the legible type face was short lived and only with the beta testers... there is no issure
  with it and no books were produced using it."* The old rule was written believing those books
  were in readers' hands. `EmbedTrueTypeFonts` stays on for a non-Tahoma book — a VistaTypeLP Sans
  book must carry its own copy of the face — and the **installer is untouched**:
  `RemoveLegacyLegibleFont` still takes that font off a machine that has it, and must, because
  every build from 3.0.101 installed it `uninsneveruninstall`.
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
  a plain one is somebody's text and is left alone. **It runs on every book from 9/4/2026** — a
  guard skipping it for a legacy VistaTypeLP Legible book went with the rest of that code, and
  both remaining faces want their fills in Tahoma.

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

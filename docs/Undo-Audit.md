# Undo audit — every macro on both tabs, 9/16/2026

**Status: CLOSED by Jerry, 9/16/2026. No undo work is wanted on any macro in VistaType LP.**
Read at build **3.0.432** (commit `70fc8d4`).

He closed it in three steps, reading this file:

1. *"there is no need to have an undo on the file cleanups, autotag ref pages and normalize styles.
   they do their job and the user should not undo them."*
2. On Format $pg Tags and Attach BANA Template: *"both are whole book jobs, no undo needed."*
3. On everything that remained — Selection Clean Up, the TOC color bars, the two picture dialogs,
   table border weights, title case, delete prodnotes, Embed/UnEmbed $pg, braille Remove Bullets,
   Spelling List, Compress Linear Math: **none of them need any undo either.**

**So every `UndoClear` stays where it is, and no macro is to be converted onto a hidden scratch
document for the sake of its undo.** Two clears were removed at his request before he closed the
question — `Sh_Color_Dollar_PG_Red` and `Sh_Para_Before_Dollar` — and that stands. No more.

**The press counts below are still accurate and worth keeping.** It is the verdicts that are void:
read "NEEDS WORK" as "costs this many presses, and that is accepted". The one part of this file
that still asks for anything is **section 4**, which is not about undo at all.

**Jerry's ruling, 9/16/2026, reads out a large part of this file.** His words: *"there is no need
to have an undo on the file cleanups, autotag ref pages and normalize styles. they do their job and
the user should not undo them."* So **both File Cleanups, both AutoTags, Normalize Styles and — by
extension — the template attach are SETTLED and off the list**, and their `UndoClear` lines stay
exactly where they are. They were the most expensive items here, at 75 to 104 passes each. He was
told plainly that `UndoClear` also erases hand edits made *before* the button was pressed; it did
not change his answer, and it should not be raised again for these.

Asked the same day where **Format $pg Tags** and **Attach BANA Template** sat, he said: *"both are
whole book jobs, no undo needed."* So `Lp_Format_Page_Numbers`, `Dx_Format_Tagged_Page_Numbers` and
`Dx_Attach_BANA_Template_Run` are settled too, clears and all.

`Sh_Fix_Ref_Pages_Before_and_After_Tables` comes off with them — the two AutoTags are its only
callers.

**Nine of the eighteen clears are therefore deliberate**, and section 1 should be read with that in
mind. The ones still worth questioning are the small macros: Background Color, Foreign Lang in
Color, the DAISY tagger, and Remove Para Formatting from Text Files — each does between one and six
replaces and then empties the list.

**One case Jerry's rule does not cover, still open.** `Dx_Is_BANA_Template_Attached` (`:4734`) runs
the whole attach by itself when no BANA template is on the document. So pressing a *small* braille
button — Dashes/Primes/Fractions, Spelling List, Exercise Levels — can empty the undo list without
the transcriber ever asking for an attach. That is not a whole-book job they chose.

**The line to work to:** a macro that works on the **whole book** to put it into shape needs no
undo. A macro the transcriber points at a **selection** to make one specific change does.

**Two of it are now done.** On 9/16/2026 Jerry asked for the clears in the two shared helpers to
come out, and they did — `Sh_Color_Dollar_PG_Red` and `Sh_Para_Before_Dollar`. Both rows below are
struck through and marked DONE. **Everything else in this file is still a survey and untouched.**

What this covers: every button on the VistaType LP and Braille Macros tabs, everything the menu
dialogs reach behind them, and the eight keyboard shortcuts in `src/keymap`. Line numbers are
`src/vba/LPandBrlMacros.bas` unless another file is named.

**How the numbers were reached, and what they are worth.** They are counted off the code, not
measured in Word. Each `Find ... Replace:=wdReplaceAll` pass is counted as one undo press, each
edit inside a loop as one per item, each `FormattedText` assignment as one. Word sometimes rolls
neighboring operations into a single undo entry, so a real press count can come out **lower**.
Where a count says *estimated*, the number depends on the book — how many pictures, blank
paragraphs or hyperlinks it holds. **Nothing here was measured on a real file.** Anything acted on
should be checked against the actual press count in Word first, the way the 39 presses for Format
TOC were measured on Jerry's own file on 9/7/2026.

Jerry's limit, 9/7/2026: *"multiple (i mean more then 3 or 4) is not an acceptable undo
requirment."*

> **Every line number in this file is wrong.** It was read at build 3.0.432; `LPandBrlMacros.bas`
> has grown since, and the drift runs from roughly +280 lines near the top to +580 near the
> bottom. Spot-checked 9/20/2026: `Lp_Normalize_Styles` cited at `:22285` is now at `:22693`,
> `Lp_TOC_CleanAndFormat_TOC` cited at `:22787` is now at `:23366`. **Find macros by name, not by
> the numbers here.** They are left in place because in several rows the number is the only thing
> distinguishing one caller from another; stripping them mangled the prose when it was tried.


## The doctrine this audit is measured against

- A custom undo record is **safe only where the macro runs no `Find`**. A `Find` with
  `Replace:=wdReplaceAll` inside an open `StartCustomRecord` **crashes Word** — access violation
  in `wwlib.dll`, nothing raised, nothing logged. Measured twice (3.0.345, 3.0.347).
- Where there are replace-all passes, the shape that works is the **hidden scratch document**:
  `Documents.Add(Visible:=False)`, passes done on **ranges**, result home in one `FormattedText`
  assignment. One press, or two when the landing is a table.
- `ActiveDocument.UndoClear` is not a cost in presses — it destroys the history outright,
  including work done before the button was pressed.
- Every `Find` in a macro converted onto a range must be `wdFindStop`. `wdFindContinue` on a range
  means the whole book.

---

## 1. The eighteen places that destroy the undo history

This is the headline. Until these go, nothing else about those macros' undo matters.

| Macro | Line | What the clear costs |
|---|---|---|
| ~~**`Sh_Color_Dollar_PG_Red`**~~ **DONE 9/16/2026** | was `:20948` | **One** replace, then the history goes. Six callers: `:9981` and `:10089` (LP File Cleanup, twice), `:12734` (LP AutoTag), `:14996` (**Validate $pg** — a review button), `:22606` (Normalize Styles), ~~`DN_Tag_Daisy_Nimas_Form.frm:80`~~ (that form was removed 9/17/2026, so five callers now) |
| ~~**`Sh_Para_Before_Dollar`**~~ **DONE 9/16/2026** | was `:26319` | Three replaces, then the history goes. Called from `:5168` (braille File Cleanup) and `:10033` (LP File Cleanup) |
| `Lp_Fix_Common_File_Errors` | `:10092` | plus the two above from inside it |
| `Dx_Fix_Common_File_Errors_Run` | `:5236` | plus `:5761` and `:26319` from inside it |
| `Lp_AutoTag_Page_Numbers` | `:12736` | plus `:20948` |
| `Dx_AutoTag_Page_Numbers` | `:6633` | plus `:3138` if it has to attach the template first |
| `Lp_Format_Page_Numbers` | `:11974` | |
| `Dx_Format_Tagged_Page_Numbers` | `:4193` **and** `:4196` | cleared twice; only `MS_Clear_F_and_R_Params_and_Clipboard` sits between them |
| `Lp_Normalize_Styles` | `:22609` | plus `:20948` |
| `Lp_Set_Page_To_Black` | `:11430` | about 4 presses' worth of work |
| `Lp_Set_Page_To_White` | `:11546` | about 4 presses' worth of work |
| `Dx_Add_Color_To_Foreign_Language_Words` | `:5761` | 6 presses' worth of work |
| `Dx_Attach_BANA_Template_Run` | `:3138` | unconditional, on every route in — see the trap below |
| `Lp_Attach_The_Template` | `:18416`, `:18453` | **arguably fair**: it ends with a Save As and `Sh_Close_And_Reopen` (`:18537`), which ends the history anyway. Lowest priority |
| ~~`DN_Add_PgNo_Tags_To_DAISY_or_NIMAS`~~ **GONE 9/17/2026** | was `:20829` | The whole manual DAISY/NIMAS route was removed — the macro, the "Choose XML Conversion Type" menu and `DN_Tag_Daisy_Nimas_Form`. Nothing left to fix |
| `DN_Remove_Para_Formatting_From_Text_Files` | `DN_Text_File_Para_Fix_Warning.frm:109` | 5 presses' worth of work. The comment at `:112–116` of that same file records an identical clear being taken off `Lp_Replace_Section_Break_With_Page_Break` — this one was left behind |

**One clear is harmless and should be left alone:** `Sh_Convert_XML_File_To_Word_Document:28074`
lands on a document VistaType LP has just created, which the transcriber has never typed in.

### The trap that widens the braille half

`Dx_Is_BANA_Template_Attached` (`:4707`) is the guard at the top of nearly every braille button,
and at **`:4734`** it does not merely report — it calls `Dx_Attach_BANA_Template_Run False`
outright when no BANA template is attached. So on a document without the template, pressing
Dashes/Primes/Fractions, Format Spelling List, Foreign Lang in Color, Exercise Levels 1 & 2,
Compress Linear Math, Full File Cleanup or Selection Clean Up runs the whole attach — the clear
at `:3138` and roughly 75 edit passes — before the button's own work starts. **Any fix that leaves
the guard alone leaves that hole open.**

### A second thing the clears hide

Both AutoTags type a paragraph mark at the top of the book and delete it at the end (LP at
`:12315` / `:12739`). The delete is placed **after** the `UndoClear`, so one press of Ctrl+Z today
re-inserts a stray paragraph at the top of the document and does nothing else. Worth fixing in
the same change; it is the most confusing thing a transcriber can hit here.

---

## 2. Macros over Jerry's limit, with no clear masking it

### Large print

| Macro | Line | Presses | Counted or estimated |
|---|---|---|---|
| `Lp_Fix_Common_File_Errors` | `:9895` | 104 fixed passes + one per blank paragraph, image, hyperlink, text box | counted / estimated |
| `Lp_Normalize_Styles` | `:22285` | 100+ | estimated |
| `Sh_Replace_Multiple_Para_Marks_No_Warning` | loop `:14170`, delete `:14177` | **one per blank paragraph removed** — hundreds on a raw book | counted shape |
| `Sh_Remove_Spaces_Before_Punctuation` | `:26749` | **15** (fifteen `Sh_Punct_Repl` calls) | counted |
| `Lp_Italics_To_Dashed_Underline` | `:10346` | **13** | counted |
| `Lp_Set_Table_Border_Weights` | `:18753` | ~21 property writes per table; Word may coalesce borders, so 3–21 per table | estimated |
| `Lp_SetSelectedTableBorderWeight` | `:19661` | 3–21 | estimated |
| `Sh_Apply_Title_Case_Capitalization` | `:20952` | one per word — a ten-word chapter title is ~10 | estimated |
| `Sh_Delete_Prodnote_Paragraphs` | `:19416` | one per prodnote + one per residual paragraph restyled | estimated |
| `Sh_Remove_Txt_Bxs_And_Frames` | `:13443` | one per shape and frame, plus a `FormattedText` each | estimated |
| `Sh_Kill_The_Hyperlinks` → `Sh_Remove_Hyperlinks` | `:12802`, `:20848` | one replace + one per hyperlink | estimated |
| `Lp_Convert_Hyper_To_Addresses` | `:13347` | one per hyperlink + 1 | estimated |
| `Lp_Remove_Box_Bullets_Bullets_and_Numbers` | `:9789`, loop `:9833` | one per list paragraph + ~10 | estimated |
| TOC **Remove Color Bars** | `Lp_TOC_Format_And_Color_Form.frm:49–117` | 3–5 per TOC entry — a 160-entry contents page runs into the hundreds | estimated |
| TOC **color bars Okay** | `Lp_TOC_Color_Bars_Form.frm:180–224` | one per TOC paragraph recolored + one per tab stop changed | estimated |
| Bkgrnd & Picture Tools → **Pictures to Inline** | `Lp_Bakgrnd_Picture_Menu_Form.frm:44` | 2 per floating picture + 1 per inline picture | estimated |
| **Picture Alignment** dialog → Okay | `LP_Picture_Alignment_Form.frm` `CmdOkay_Click` | one per picture | estimated |

Carried along by their callers rather than reachable alone: `Lp_Add_Para_After_Image` (`:17428`),
`Lp_SetPicturesToInlineAndLockAspectRatio` (`:20671`),
`Lp_ResizePicturesAndShapesToFitPageWidthAndPageHeight` (`:19807`),
`Lp_Set_TOC_and_Print_Page_Num_Tab_Stops` (`:18666`), `Lp_Merge_Adjacent_Pg_Tags` (`:11988`).

### Braille

| Macro | Line | Presses | Counted or estimated |
|---|---|---|---|
| `Dx_Fix_Common_File_Errors` | `:5008` / `:5028` | **~75 fixed passes** + unbounded per-item loops | counted / estimated |
| `Dx_Selected_File_CleanUp` | `:5376` | ~40+, depending on what is ticked (13 macros on the list) | estimated |
| `Dx_Remove_Bullets` | `:7750` | ~19 (6 minimum, +5 on a selection, +8 via `Dx_Convert_Auto_List_To_Text`, +1/link) | counted |
| `Dx_Compress_Linear_Math` | `:8688` | **11** | counted |
| `Dx_Spelling_List` | `:5852` | **10** (9 replaces + 1 style set, in `Dx_Spelling_List_Options_Form.frm`) | counted |
| `Dx_Convert_Auto_List_To_Text` | `:4785` | **8** | counted |
| `Dx_UnEmbed_Ref_Pg_No` | `:4283` | ~8 — **and it uses Cut/Paste, so it empties the transcriber's own clipboard** | counted |
| `Dx_Embed_Ref_Pg_No` | `:4220` | ~6 | counted |
| `Dx_Encode_Fractions_With_DBT_Codes` | `:9120` | one per fraction + 4 | counted shape |
| `Dx_Convert_Hyper_To_Addresses` | `:8098` | one per hyperlink + 1 | counted shape |
| `Dx_Convert_Hyperliks_To_Text` | `:8842` | one per hyperlink × 2 loops + 1 | counted shape |
| `Dx_Delete_Images` | `:8163` | one per floating shape + 2 | counted shape |
| `Dx_Replace_Fraction_Text_With_Compact_Fractions` | `:8897` | one per fraction, 18 patterns scanned | counted shape |

### Both sides

`Sh_Fix_Ref_Pages_Before_and_After_Tables` (`:27191`) costs **2–3 per table** — `SplitTable`
`:27217`, `InsertAfter vbCr` `:27223`, `WrapAroundText` `:27227`. A 20-table book is 40 to 60
presses on its own. It is reached only from the two AutoTags, so it rides along with whatever
they get — **but it drives the table through `Selection`**, so it must be moved onto a range
before either AutoTag can run in a hidden document.

### Where the big two get their totals

`Sh_Replace_All_Until_Done` (`:26106`) repeats one replace until it stops changing things —
two or three rounds typically. It is the fix for alternate-number tagging, not tidiness, so it is
not removable. But it means each of the 22 patterns in `Dx_AutoTag_Page_Numbers` and the 12 in the
LP twin is 2 or 3 undo entries rather than 1. That is most of why the AutoTags are so expensive,
and the strongest argument for the hidden round trip there — it makes the number irrelevant.

Braille Full File Cleanup's ~75 comes from the 37 steps at `:5028–5251`. The four largest are
`Sh_Remove_Spaces_Before_Punctuation` (15), `Dx_Convert_Auto_List_To_Text` (8),
`Dx_Add_Color_To_Foreign_Language_Words` (6), `Dx_Fix_Para_Space_Errors` (5). On top sit loops
with no ceiling: per Abbyy-restyled paragraph and deleted style (`:8505–8557`), per shape and
frame (`:13526–13597`), per header and footer (`:26237–26245`), per hyperlink, per fraction, per
blank paragraph mark (`:14170–14203`).

---

## 3. Already done — do not revisit

**Hidden scratch document, at 1 or 2 presses:** `Lp_Table_Convert_Hidden` (`:21674`, 2 —
`srcTbl.Delete` plus one `FormattedText`), `Lp_TOC_CleanAndFormat_TOC` (`:22787`, 1 normally, 4
when the TOC ends the book), `Lp_Horz_To_Vert_Hidden` (`:15292`, 1),
`Lp_Exercise_Levels_Hidden` (`:14612`, 1), `Dx_Exercise_Levels_Hidden` (`:7092`, 1 — **the
worked example on the braille side**), both Import Exported Selection macros (1 each).

All ten table helpers are reached **only** from inside `Lp_Table_Convert_Hidden`, working on the
scratch document's table, so none of them touches the book. Dialogs 178 and 200 truthfully say
two presses.

**Custom undo record, correctly used, at 1 press:** `Lp_Resize_Images` (`:19854`),
`Lp_Change_Image_Color` (`:20522`),
`Lp_Type_Fill_In_Line_To_Margin` (`:15975`), `Lp_Type_Counted_Fill_In_Lines` (`:16443`),
`Lp_Compress_Linear_Math` and `_In_Selection` via `Lp_Clm_Run_As_One_Undo` (`:16919`),
`Lp_Replace_Section_Break_With_Page_Break` (`Lp_Section_Brk_Caution.frm`),
`Lp_Convert_Table_To_Pseudo_Columns` (`:17632`), `Lp_Convert_Table_To_Real_Columns` (`:17787`),
both Export Selection macros.

**Resize Pictures Used Throughout a Selected Range, 9/22/2026: one press of Ctrl+Z per press of
Apply, and a session cannot be fewer.** The box stays open while the transcriber works, so an undo
record held open across it would swallow her own typing and her own hand resize of the next
picture. Each Apply opens and closes its own record; dialog 381 says how many presses it will
take to put a whole session back, rather than promising one.

**Tried and rejected, 9/21/2026: an `UndoClear` at the start of `Lp_Resize_Same_Picture_Throughout`
(that macro was retired on 9/22/2026; the rule holds for the range loop that replaced it).**
It stopped a third Ctrl+Z from going on into earlier edits, but it also wiped the transcriber's
hand resize of the model picture, which comes just before the macro and must stay undoable.
Word cannot empty only part of the undo list, so Jerry had it taken out (3.0.473 → reverted).
Do not add it again.

**One real edit, nothing needed:** `Lp_Keep_With_Next_Para`, `Lp_Toggle_Space_After_Current_Para`,
`Sh_Keep_Lines_Of_Para_Together`, `Sh_Move_Paragraph_To_Next_Page`, both Manual Tag macros (3–4),
`Dx_Change_Prodnotes_To_Transcriber_Notes`, `Dx_Set_Whole_Document_To_Times_New_Roman_14`,
`Dx_Remove_Breaks`, `Dx_Remove_Column_Breaks`, `Dx_Remove_Keep_With_Next`,
`Dx_Remove_Optional_Hyphens`, `Sh_Remove_Multi_Spaces`, `Sh_Replace_Manual_Line_Break`,
`Sh_Replace_Small_Caps_With_All_Caps`, `Lp_Replace_Tabs_With_Single_Space`,
`Lp_Fix_Para_Space_Errors` (4), `Lp_Convert_Auto_List_To_Text` (3).

**Changes nothing, so needs no undo:** `Sh_Doc_Info`, both About boxes, both Video Links (they
only open a web page), `Dx_Save_Braille_File`, `MS_Reset_Word_Configuration`,
`Sh_Show_Recommended_Styles_Pane`, the whole `Sh_PgVal_*` validation walk-through (checked every
write in `ShNonModalMessage.bas` — there are none; the list document is closed
`SaveChanges:=wdDoNotSaveChanges` at `:375`), and every menu starter.

**`Lp_Validate_Dollar_PG` is the exception in that last group** — it is presented as a review
button but it calls `Sh_Color_Dollar_PG_Red`, so it edits once and then destroys the history.

### The crash shape is not present anywhere

All four readers checked specifically for a `wdReplaceAll` inside an open custom undo record,
statement by statement through every record-opener in the project. **There is none.** Verified in
particular:

- `Lp_Type_Fill_In_Line_To_Margin`: between `StartCustomRecord` `:16095` and `EndCustomRecord`
  `:16411` there is no `Find` and no `.Execute` — only `TypeText`, `TypeBackspace`, `InsertAfter`.
  It calls `Lp_Remove_Fill_In_Line_At_Cursor` inside the record (`:16112`); that is
  `doc.Range(a,b).Delete` only.
- `Lp_Clm_Run_As_One_Undo` `:16919` → `Lp_Clm_Compress_Span` `:17190`: edits by
  `doc.Range(...).Text = ""` and `.InsertAfter`, no `Find` at all.
- `DxExportImportSelectedText.bas:243` wraps a bare `rngSel.Delete`.
- `LpExportImportSelectedText.bas`: the `wdReplaceAll` at `:156` is on the **new export
  document** and closes well before the record opens at `:209`.

**`Lp_Tahoma_The_Fill_Ins` (`:15949`) does run a `wdReplaceAll`**, but never inside an open record
— its callers are `Lp_Ex_Passes:14771` (on the hidden scratch document) and
`Lp_Attach_The_Template:18162` (no record open). No risk as things stand. Worth a comment there so
nobody later wraps the attach in a record.

---

## 4. Found while reading — not about undo

These were defects noticed while reading every macro, none of them about undo. **They are now
GitHub issues**, which is where this project records defects from 9/20/2026. They were lifted
out of here because a closed survey is the wrong place to keep live work.

| What | Issue |
|---|---|
| `Lp_Set_Table_Border_Weights` writes the Word-wide `Options.DefaultBorderLineWidth` and never puts it back | #2 |
| Braille Remove Section Breaks still carries the book-corrupting `^b` delete (no callers) | #3 |
| `Lp_Italics_To_Dashed_Underline` uses `wdFindContinue` on all 13 passes, though offered as a selection repair | #4 |
| `Lp_TOC_Color_Bars_Form` selects the whole document before its remove-leaders loop | #5 |
| Compress Linear Math retired on the large-print tab, still live on the braille tab | #7 |
| `Dx_Is_BANA_Template_Attached` runs the whole attach, so a small button empties the undo list | #8 |
| Five macros with no callers | #9 |
| Three more that need a read to confirm | #10 |

The one item that closed itself: the DAISY `Replace:=wdReplaceAl` typo went with the manual
DAISY route on 9/17/2026.

## Before any of this ships

Nothing in the build compiles VBA. Any change here needs **Debug → Compile by hand on the build
box** (Alt+F11 → Debug → Compile LPandBRL, against the STARTUP `.dotm` opened directly) and a real
`make try` in Word. Changes to File Cleanup and the macros are covered by `make try`; anything
touching `LargePrintTemplate.dotx`, the ribbon or the keymap needs a full `Setup.exe`.

And add a row to `docs/Reported-Errors.md` in the **same** change as any fix, not afterwards.

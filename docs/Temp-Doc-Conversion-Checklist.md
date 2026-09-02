# Hiding the temporary document — what to test

Horiz List to Vertical was converted first (3.0.147) and is the worked example. This is the
list of everything else that goes through the same round trip, so each one can be ticked off as
it is converted and tried.

## What the change is

`Lp_Copy_To_Temp_Doc` and `Dx_Copy_To_Temp_Doc` create the scratch document **hidden**, and then
deliberately give it a window, show it, maximize it and activate it:

```vba
If tempDoc.Windows.count = 0 Then tempDoc.Windows.Add
With tempDoc.Windows(1)
    .Visible = True
    .WindowState = wdWindowStateMaximize
End With
tempDoc.Activate
```

That is the flashing. It is there because those callers work through `Selection`, and `Selection`
only reaches the document that is active. Converting a macro means moving its passes onto a
**range** — the scratch document's own content — so the window is never needed.

The round trip itself stays. It is what gives the single undo and the private workspace, and it
has never lost anyone's work.

## What to look for on each one

- **No flash.** The screen should not blink to a blank document and back.
- **One Ctrl+Z** puts the document back as it was.
- **Formatting survives**, indentation especially — see the traps below.
- **The paragraph before and after the selection is untouched**, both its text and its formatting.
- **A selection that stops mid-paragraph** does not gain a paragraph mark.
- The result is otherwise **identical to the old version**. Run the same case on the previous
  build if there is any doubt.

## Two traps, both about paragraph marks

Both of these were found the hard way on 8/12/2026 and pull in opposite directions:

1. **The marks must travel with the text.** A paragraph's formatting — its indentation among it —
   lives *in* its paragraph mark. Leave the last mark behind and the last item comes back with
   the scratch document's defaults.
2. **They must come back in the same number.** The cleanup passes collapse runs of paragraph
   marks, so the text can return ending with fewer than it left with. That deletes the mark after
   the last item and welds it onto the paragraph below.

`Lp_Horz_To_Vert_Hidden` handles both: it sends everything, then restores the count before
handing the text home. Copy that approach.

Also: build the scratch document from the **source document's own attached template**, so styles
resolve the same way. That is what lets one routine serve both large print and braille — a BANA
document gets a BANA scratch document.

## Merge a pair when the conversion makes them identical

**Jerry's rule, 8/12/2026.** Everything that differs between an `Lp_` and a `Dx_` version of the
same helper usually exists only to serve the round trip, so taking the round trip out tends to
leave the two the same job written twice. When that happens, merge them into one `Sh_` macro and
repoint the callers.

The test is whether they differ in WORK or only in plumbing. `Replace_Tabs_With_Single_Space`
converted to ranges and still stayed two macros, because the braille side has underlined-tab
passes the large-print side has no use for.

`Sh_Remove_Multi_Spaces` was the first merge. Before the conversion the two genuinely differed — the
braille copy reached up a paragraph before copying out, the large-print copy wiped the whole undo
history afterwards. After it, the only difference left was the order of five tidy-up lines.

Helpers are usually not on the ribbon, so merging them costs nothing in button ids — unlike
`Lp_Horz_List_To_Vertical`, where two ids had to be kept alive.

## Two more traps, from Fix_Para_Space_Errors

Both were things the temporary document was quietly handling, and both are easy to lose.

**Never include the document's own final paragraph mark.** Any pass that rewrites `^013` will
rewrite that one too and ADD a paragraph. The old code cleaned up afterwards with "go to the end,
delete one character" — which looked like pure temp-document tidying and was removed as such. It
was not: on the no-selection route it was also removing the paragraph this pass had just created,
and it deleted the last character of the transcriber's text into the bargain. Not touching the
mark beats tidying up after it.

**When there is a selection, reach back one character to take in the paragraph mark in front of
it.** "Spaces after a paragraph mark" cannot match the first selected paragraph otherwise,
because its opening mark sits outside the selection — so leading spaces on the first line
survive. The temporary document got this right by accident: it began by typing a paragraph mark
at the top so the pattern had something to match. Including the real one is the same trick
without the scratch document, and nothing before the selection is harmed, because every one of
these patterns puts the mark back.

`Sh_Para_Fix_Range` does both. Use it for any macro whose passes touch `^013`.

## Two traps that are not about paragraph marks

**A quirk may be load-bearing.** `Dx_Replace_Manual_Line_Break` decided whether to ask by whether
text was selected. That looked like an inconsistency next to the large-print copy, which always
asked — so it was "fixed" to always ask, and that broke Full File Cleanup on the braille side,
which runs it with nothing selected and must not stop for a dialog. The old test was doing the
right thing for the wrong reason. What actually separates the cases is **who is calling**: a
sequence knows the answer, a person is asked. So the answer became a parameter.

Before removing an oddity, find out which sequences reach the macro and what they rely on.

**Nothing in the build compiles VBA.** A compile error ships and reaches the transcriber as
"Compile error in hidden module". Running any macro against the built `.dotm` forces Word to
compile the whole project, and takes seconds:

```powershell
$doc = $word.Documents.Open("...\dist\LPandBRL.dotm", $false, $false)
$word.Run("Sh_Doc_Config_Type")     # any macro will do
```

Do it before handing over a build. `Application.Run MacroName:="X", "arg"` is the mistake that
made the point: once `MacroName:=` is named, every later argument must be named too. A public sub
in the same module can simply be called — `Sh_Replace_Manual_Line_Break "Para"`.

## Buttons that reach the round trip

Derived by following the call graph from each ribbon button, so it may over-report: some of
these reach it only through a shared cleanup helper. Confirm per button as you go.

**It DID over-report, and by a lot — traced 8/31/2026.** Following every macro each button can
reach, dialogs included, only **seven places** in the whole project still called a
scratch-document routine that morning: `Dx_Format_Exercise_Lv_1_and_Lv_2`,
`Lp_Format_Exercise_Lv_1_and_Lv_2`, `Lp_Table_Convert_Options_Form` (five calls),
`Lp_TOC_CleanAndFormat_TOC`, `Lp_Resize_Images` (two calls), `Lp_Change_Image_Color_Form` and
`Lp_Section_Brk_Caution`.

**Six of those seven are done, and ONE is left.** Settled at 3.0.309
(`Dx_Format_Exercise_Lv_1_and_Lv_2`, converted), 3.0.319
(`Lp_Format_Exercise_Lv_1_and_Lv_2`, converted), 3.0.321 (`Lp_Resize_Images`, round trip
**deleted** — it never needed one), 3.0.326 (`Lp_Section_Brk_Caution`, round trip **deleted**,
and a book-corrupting defect found and cured with it) and 3.0.338
(`Lp_TOC_CleanAndFormat_TOC`, round trip **deleted** — the third of these that never needed one).
Still calling `Lp_Copy_To_Temp_Doc`: `Lp_Table_Convert_Options_Form` (five calls).

**Three of the six settled were deletions, not conversions.** That is now the expected answer, not
the surprise one: read what the macro does to the scratch document before planning how to move it
onto a range.

**None of the six braille $pg buttons is among them** — AutoTag Ref Pages, Validate $pg Tags,
Manual Tag Ref Page, Format $pg Tags, Embed and UnEmbed reach no scratch document at all. The
three things in their reach that DO show or activate a document are all deliberate: the
`Documents.Add` in `Dx_Attach_BANA_Template_Run` fires only when nothing is open at all; the
second document in `Sh_Copy_Ref_Pg_Tags_To_Temp_File` is the validation list, which is meant to
be read; and the `.Activate` calls in `ShNonModalMessage` put the transcriber back in her own
document. Leave those alone.

Two more findings from the same trace, both worth acting on separately. **Both were acted on at
3.0.318, 8/31/2026 — the five dead procedures are gone:**

- **`Sh_Copy_To_Temp_Doc` and `Sh_Copy_From_Temp_Doc` were called from nowhere** - not by a macro,
  a form, a ribbon button or a keyboard shortcut. They were written on 8/11/2026 to pick the
  right route for the horizontal-list merge, and whatever used them had since moved to the
  hidden route. Removed, with the `Sh_TempDocRoute` variable that only they wrote.
- **`Dx_Copy_To_Temp_Doc` had no live caller either from 3.0.309**, the braille exercise macro
  having been its last one. The braille side of this list is finished. Removed, along with
  `Dx_Copy_From_Temp_Doc` and `Dx_Attach_Same_BANA_Template` — which nothing but
  `Dx_Copy_To_Temp_Doc` ever called — and the `Dx_Attached_BANA_Template` variable, whose only
  reader was `Dx_Attach_Same_BANA_Template`. **Dialog 270 went with them and its number is
  retired, not reused.**

  Why remove rather than leave: dead code that still reads `Documents.Add`, `Selection.Copy` and
  `Selection.Paste` is exactly the thing that gets called again by someone tidying up later, and
  it would put the flashing scratch document straight back. A tombstone comment stands where each
  one was.

**`Lp_Copy_To_Temp_Doc` and `Lp_Copy_From_Temp_Doc` stay, and still show their document.**
**One** place calls them from 9/2/2026: the `Lp_Table_Convert_Options_Form` — what is left of
the large-print half of this list. `Lp_TOC_CleanAndFormat_TOC` came off at 3.0.338.
`Lp_Change_Image_Color_Form` is off it too: it reaches `Lp_Copy_To_Temp_Doc` only through the
Picture Tools menu, and was counted separately above.

**Ask first whether the round trip is needed at all.** `Lp_Resize_Images` was on this list as a
conversion and turned out to be a deletion: nothing it does needs a document of its own. Two
tickets off this list so far have been of that kind. Read what the macro actually does to the
scratch document before planning how to move it onto a range — if every pass is something a
`Range` can be asked for directly, the answer is to delete the round trip, not to rebuild it.

### Braille Macros tab

- [x] Attach BANA Template — `Dx_Attach_BANA_Template` (Jerry, 8/13/2026, build 3.0.174)
- [x] Full File Cleanup — `Dx_File_Fix_Sequence` (Jerry, 8/13/2026, build 3.0.174)
- [x] Selection Clean Up — `Dx_Selected_File_CleanUp` (Jerry, 8/13/2026, build 3.0.174)
- [ ] AutoTag Ref Pages — `Dx_AutoTag_Page_Numbers`
- [ ] Validate $pg Tags — `Dx_Ref_Pg_Number_Sequence_Menu`
- [ ] Manual Tag Ref Page — `Dx_Manual_Tag_with_Dollar_pg`
- [ ] Format $pg Tags — `Dx_Format_Tagged_Page_Numbers`
- [ ] Embed Ref Pg Numb — `Dx_Embed_Ref_Pg_No`
- [ ] UnEmbed Ref Pg Numb — `Dx_UnEmbed_Ref_Pg_No`
- [ ] Foreign Lang in Color — `Dx_Add_Color_To_Foreign_Language_Words`
- [ ] Format Spelling List — `Dx_Spelling_List`
- [x] Exercise Levels 1 & 2 — `Dx_Format_Exercise_Lv_1_and_Lv_2` (3.0.309, 8/31/2026).
      **The 8/13/2026 tick on this line was wrong, and it stood for eighteen days.** What was
      tested that day was Convert Auto List To Text, which this button reaches and which had
      just been converted; the button itself still made its own scratch document with
      `Dx_Copy_To_Temp_Doc` and still put it on screen. Jerry reported it again on 8/31/2026 —
      "the screen goes blank while the macro is working on the temp doc". A tick here means THIS
      BUTTON no longer round-trips through a visible document, not that a helper it calls was
      converted. Converted at 3.0.309: `Dx_Exercise_Levels_Hidden` does the round trip on a
      document created `Visible:=False`, and `Dx_Tabs_To_Fill_Ins` and `Dx_Fix_Para_Space_Errors`
      took an optional range so they could be run against it.
- [ ] Dashes/Primes/Fractions — `Dx_Type_Dashes`
- [ ] Compress Linear Math — `Dx_Compress_Linear_Math`
- [ ] DAISY or NIMAS to Word — `DN_Menu_Starter`

### VistaType LP tab

- [x] Attach LP Template — `Lp_Attach_Lp_Template` (Jerry, 8/13/2026, build 3.0.174). Not on
      this list originally, and it should have been: the attach runs `Lp_Fix_Common_File_Errors`
      end to end, so every helper converted for the cleanup reaches a transcriber through the
      attach as well — on a whole document, with nothing selected.
- [x] Full File Cleanup — `Lp_File_Fix_Sequence` (Jerry, 8/13/2026, build 3.0.174)
- [x] Selection Cleanup — `Lp_Selected_File_CleanUp` (Jerry, 8/13/2026, build 3.0.174). One
      option on that menu, **replace section breaks**, still round-tripped until 3.0.326
      (9/1/2026) — `Lp_Replace_Section_Break_With_Page_Break` shows `Lp_Section_Brk_Caution`, and
      the Okay button on that form was what called `Lp_Copy_To_Temp_Doc`. A tick against a menu
      is not a tick against everything under it.

      **And the round trip was the least of it.** Replacing every section break with a manual
      page break collapses the document into ONE section, so every section's page setup is
      discarded. Jerry's own test book — eight sections, seven starting on an odd page, each one
      immediately before a `Print Pg Num` paragraph — was "fine before the macro, corrupted
      after", and repaginating what came out killed Word outright. That had been true for as long
      as the macro existed.

      Cured by not deleting the breaks at all: whether a section starts on the next page or the
      next *odd* page is a **property**, so `SectionStart` moves from `wdSectionOddPage` to
      `wdSectionNewPage`. Every section and its page setup survive, no paragraph mark is touched,
      the blank filler pages are gone and each chapter still opens on a page of its own — which
      is the purpose of the button. It works on the **whole book** now, never a selection, and no
      longer demands one; both dialogs were reworded and dialog 129 is retired. Full account,
      including two wrong turns, in `docs/Reported-Errors.md`.
- [ ] AutoTag Ref Pages — `Lp_AutoTag_Page_Numbers`
- [x] Format Exercise — `Lp_Format_Exercise_Lv_1_and_Lv_2` (3.0.319, 9/1/2026). It had been
      ticked on 8/13/2026 for the same wrong reason as its braille twin above: what was tested
      was Convert Auto List To Text, which this button reaches and which had just been
      converted. The button itself still called `Lp_Copy_To_Temp_Doc` and the screen still went
      blank. Converted onto `Lp_Exercise_Levels_Hidden`, `Lp_Ex_Passes` and `Lp_Ex_Repl`, with
      `Lp_Fix_Para_Space_Errors` and `Sh_Remove_Spaces_Before_Punctuation` gaining an optional
      range so they could be run against the hidden document.

      **NOT merged with the braille twin, and the merge rule expected it would be.** Once both
      were on ranges they turned out not to be the same job written twice: this one applies
      `List` and `List 2`, doubles the paragraph marks so each item gets a blank line, and
      rebuilds fill-in lines as underlined Tahoma underscores; the braille one applies
      `Exercise1`, `Exercise2` and `Ex2Nemeth2`, places `[[*kps*]]` and `[[*kpe*]]` markers
      hidden and plum, and writes UEB or EBAE fill-in indicators. The shape of the round trip is
      shared — and it IS shared, through `Sh_Trailing_Para_Marks` and
      `Sh_Set_Trailing_Para_Marks` — but nothing between the two ends is.

      Three faults went with the old route, all of them invisible until they bit:
      `Lp_Copy_To_Temp_Doc` looks `LargePrintTemplate.dotx` up **by path**, and on a miss it
      shows a message and does a plain `Exit Sub` the caller never sees — so all twenty finds
      ran on the transcriber's own book, each with `wdFindContinue`, meaning the whole book and
      not the selection. A selection made inside a **table** was widened to the whole table on
      the way out but pasted back over the selection on the way home. And the way home was a
      guess: `Documents(1)`, or `Documents(2)` when that was the scratch one, which is wrong the
      moment a third document is open.
      (Also recorded here because it was worth finding out: this is the LP button that reaches
      Convert Auto List To Text — NOT Fill-In Line, which was reported to Jerry as its caller by
      mistake. `Lp_Format_Exercise_Lv_1_and_Lv_2` begins two lines after `Lp_Type_Fill_In_Line`
      ends, and a script that walked back to the nearest preceding Sub landed on the wrong one.
      Fill-In Line reaches nothing on this list.)
- [ ] Table and TOC Tools — `Lp_Table_Tools`
      - The **TOC** half is done (3.0.338). `Lp_TOC_CleanAndFormat_TOC` works on the selected
        range in the book; the round trip was **deleted**, not converted. What went with it:
        the whole book on a machine that cannot find `LargePrintTemplate.dotx` (the macro ran
        all nineteen replaces across it and then closed it with `SaveChanges:=wdDoNotSaveChanges`,
        because `Lp_Copy_To_Temp_Doc` reports a missing template and does a plain `Exit Sub` that
        `Application.Run` never hands back); a TOC inside a table, widened on the way out and
        pasted back narrow; the transcriber's view, flipped to Print Layout every run; her
        clipboard, emptied although nothing was ever put on it; and two dead regular expressions.
        **Every `Find` is `wdFindStop` now** — the old code used `wdFindContinue`, which is safe
        only inside a scratch document and means "the whole book" on a range.
        `Sh_Para_Fix_Range` answers both paragraph-mark traps; it is used ONLY for the two passes
        that need the mark in front of the selection, because every other pass would otherwise
        reach into the paragraph above the TOC.
        **Left alone deliberately:** the two `Print Pg Num` passes at the end are not repeatable —
        formatting the same TOC twice gives every reference page number two leading non-breaking
        spaces and two tabs.
      - The **table** half is what remains — `Lp_Table_Convert_Options_Form`, five calls.
- [ ] Bkgrnd & Picture Tools — `Lp_Picture_Tools_Menu_Starter`. **Resize Images, one of the
      things under it, is done** (3.0.321, 9/1/2026) — and it was a DELETION, not a conversion.
      All the macro does is walk images and set their scale, and an `InlineShapes` collection
      comes off a `Range` as readily as off a document; two of its four branches asked
      `ActiveDocument` for its images instead of asking the selection for its own, which is the
      only reason the selection had to *be* a document.

      Five faults went with the round trip, none of them ever reported:

      - **The clipboard, twice.** The table branch carried the table home through
        `Selection.Copy` / `Selection.Paste`, and `MS_Clear_F_and_R_Params_and_Clipboard` then
        emptied the clipboard on *every* route out, including the three that never used it.
      - **The table, deleted and pasted back.** `Selection.rows.Delete`, a `TypeBackspace` and a
        delete-one-character removed the transcriber's own table so the scratch document's copy
        could be pasted in its place.
      - **Cells she had not selected.** `Lp_Copy_To_Temp_Doc` widens a selection made inside a
        table to the entire table, so images in other cells were resized too.
      - **The whole undo history**, thrown away by `ActiveDocument.UndoClear` — so a resize could
        not be taken back, and neither could anything done before it. It is one custom undo
        record now, `"Resize Images"`, the same idiom as Reflow Table and the two exports.
      - **The book**, on a machine without `LargePrintTemplate.dotx`. `Lp_Copy_To_Temp_Doc` looks
        it up by path and quietly gives up on a miss; on the table branch that meant deleting her
        table and pasting whatever was on the clipboard.

      The other buttons under this menu are still to do — `Lp_Change_Image_Color_Form` still
      round-trips.
- [ ] DAISY or NIMAS to Word — `DN_Menu_Starter`

### Already done

- [x] Horiz List to Vertical — `Lp_Horz_List_To_Vertical` (both tabs, 3.0.147)
- [x] Remove Multi Spaces — now `Sh_Remove_Multi_Spaces` (3.0.149, first merge under the rule).
      Reached by Full File Cleanup and Selection Cleanup on both tabs, so test it through those.
- [x] Replace Tabs With Single Space — `Lp_` and `Dx_` both converted (3.0.150), and deliberately
      **not** merged: the braille one turns UNDERLINED tabs into underscores first, because a
      scanner leaves a ruled fill-in line as an underlined tab and a braille transcriber needs
      the underscores. That is different work, not different plumbing — the exception that shows
      what the merge rule actually means. Also reached by all four Cleanup buttons.
- [x] Replace Manual Line Break — now `Sh_Replace_Manual_Line_Break` (3.0.153). Converted at
      3.0.151, then merged once Jerry settled the four ways the two sides disagreed: it always
      asks; cancelling stops the run and the cleanup sequence with it; multiple spaces are
      removed only for "Space"; and the `Dx_Fix_Para_Space_Errors` call is gone from the
      large-print side. The answer is a **parameter**: a caller that knows it passes it and
      nothing is asked, otherwise the transcriber is asked. That distinction matters — see below.
      **Cannot be tested from a background session at all**, since asking needs a person. Test
      Full File Cleanup on the braille tab first (it must run straight through, no dialog), then
      Selection Cleanup on both tabs, including Cancel.
- [x] Fix Para Space Errors — `Lp_` and `Dx_` both converted (3.0.157), **not** merged: the
      large-print one protects the "1 point" style, the braille one handles middle dots and tabs
      after paragraph marks. Merging would give each side passes nobody asked for. New shared
      `Sh_Para_Fix_Range` — see the traps above.
- [x] Replace NonBreaking Spaces — three macros became one (3.0.159). `Dx_Replace_NonBreaking_Spaces`
      (temp document) and `Dx_Replace_NonBreaking_Space_With_Space` (a bare whole-document replace
      with no scoping at all) are gone; everything runs through
      `Sh_ReplaceNonBreakingSpacesWithNormalSpace`, the only one that protects the non-breaking
      spaces holding a **Print Pg Num** bar together. Runs are collapsed afterwards on the same
      range, because the slow path must replace one character with one character.
- [x] Kill The Hyperlinks — now `Sh_Kill_The_Hyperlinks` (3.0.160). The scratch document was
      doing real work: `Sh_Remove_Hyperlinks` deletes every hyperlink in the ACTIVE DOCUMENT, so
      copying the selection out was the only thing confining it. It takes a range now; the
      default is still the whole document, so its three unconverted callers are unaffected.
- [x] Replace Multiple Para Marks — now `Sh_Replace_Multiple_Para_Marks_No_Warning` (3.0.161).
      Two bugs, both found by Jerry running it: the braille side was still doing
      `^013{2,}` -> `^p` through the temp document, which needs the marks to be TOUCHING and so
      never sees a line holding a single space; and the large-print paragraph walk added in July
      ignored the selection entirely, while the caution box was telling the transcriber that
      selecting text is how you limit it.
      At 3.0.163 the blank test became Jerry's own `Lp_IsEmptyPara` — tabs and non-breaking
      spaces count as empty too — and `Lp_IsEmptyPara` itself is gone. `Sh_IsBlankParaMark` is
      the one blank test in the project now. It leaves `Chr(7)` alone on purpose: that is what
      stops an empty table cell reading as an empty paragraph.

      **The rule, settled by Jerry 8/13/2026: EVERY empty paragraph goes — a run of them, and a
      lone blank line too — and every macro that goes looking for empty paragraph marks works
      that way.** "Replace multiple paragraph marks with a single paragraph mark" means one mark
      between the two paragraphs, which is no blank line at all; the caution box has said since
      2018 that deliberately added blank lines will be removed. Read as "collapse a run down to
      one blank line" for two builds, which is what Jerry was reporting each time. The only
      paragraph mark left standing is the document's own final one, which Word will not give up. Macros that touch runs
      of `^p` as part of a different job — Format Exercise, `Sh_Para_Before_Dollar`,
      `Sh_Remove_Empty_Para_Before_Tables`, `Dx_Remove_Page_Breaks`, the DAISY/NIMAS text
      re-paragraphing, the two table tools — are not covered by it and were audited and left
      alone; the count is the point in each of them.

      And one that was not on any list. `Lp_Convert_Hyper_To_Addresses` finished by running
      `Selection.Range.AutoFormat` over the whole document to turn plain URLs blue — and
      **AutoFormat removes empty paragraphs.** It is reached by Fix Common File Errors, so it
      ran on every cleanup and every template attach, and it took blank lines out of documents
      with no hyperlinks in them at all. Jerry found it by noticing that Fix Common File Errors
      disagreed with itself. Turning off every AutoFormat option except `ReplaceHyperlinks` does
      **not** stop it — that was tried and measured; the paragraph analysis is not one of the
      options. AutoFormat is gone (3.0.166) and `Sh_Linkify_Range` makes the links directly.
      Two things it has to get right: skip text that is already inside a hyperlink, and walk
      each paragraph's tokens **back to front**, because `Hyperlinks.Add` inserts a field and
      every offset after it moves. Front to back linked the first address on a line and silently
      missed the rest.

      The braille twin never had the bug — `Dx_Convert_Hyper_To_Addresses` has no AutoFormat
      call. It also makes no links; it only strips them to text.

### Group 2 — shapes, frames and numbering (3.0.169)

Jerry's guess was right: **the round trip was doing real work in every one of these**, and more
of it than anywhere else so far. `ActiveDocument.Shapes`, `.Frames`, `.ListParagraphs` and
`ConvertNumbersToText` are whole-document by nature — there is no "the shapes in this selection"
to ask for — so copying the selection into a scratch document was the only thing keeping any of
them off the rest of the book.

- [x] Remove Txt Bxs And Frames — now `Sh_Remove_Txt_Bxs_And_Frames` (3.0.170). Converted as two
      macros and reported as "different work, not merged"; Jerry's answer was to make the
      large-print one behave like the braille one — ungroup first, take text from ANY shape that
      has some, keep the formatting — and with that they were the same job written twice, so they
      merged. **The rule cuts both ways.** A real difference in behavior is a question for Jerry,
      not a reason to stop: he may want one of the two behaviors everywhere. Ask before assuming
      the difference is wanted.
      Two large-print faults went with it: the text came out as a plain STRING, so a bold word in
      a text box arrived flat, and the closing marker had no paragraph mark after it, so it welded
      onto the front of the anchor paragraph.
- [x] Remove Bullets — `Dx_Remove_Bullets`. Three separate things reached past the selection
      here: `Sh_Remove_Hyperlinks` (whose default is still the whole document),
      `Dx_Convert_Auto_List_To_Text`, and every `Selection.Find` with `wdFindContinue`.
- [x] Remove Box Bullets, Bullets and Numbers — `Lp_Remove_Box_Bullets_Bullets_and_Numbers`.
- [x] Convert Auto List To Text — `Lp_` and `Dx_`, converted here rather than in group 3 because
      `Dx_Remove_Bullets` calls it. **A macro scoped to a range cannot call one that goes through
      a temporary document**: the paste back invalidates the range it was holding. Convert the
      callee first, always.

New traps from this group:

**A shape belongs to a range when its ANCHOR does.** A floating shape has no place in the text of
its own; the anchor is what travels with copied text, so the anchor is what decides.
`Sh_Shape_In_Range` and `Sh_Frame_In_Range` do it, and both check `StoryType` first — a shape
anchored in a header has an anchor whose `.Start` counts from the beginning of THAT story, and it
would otherwise fall inside the main story's numbers by coincidence.

**Delete shapes backwards by index.** `For Each` over `Shapes` while deleting from it skips the
next shape. `Lp_Remove_Txt_Bxs_And_Frames` was working around that by running the whole pass four
times; backwards by index, one pass is enough and nothing is missed.

**`Range.FormattedText` replaces the clipboard.** The braille copy kept the text box's formatting
by copying and pasting. Assigning `FormattedText` to a COLLAPSED range inserts at that point and
keeps the formatting, without taking the transcriber's clipboard away.

**Selecting a range does not flash.** Where a helper still reads `Selection` — the large-print
Convert Auto List To Text needs `Lp_Toggle_Space_After_Current_Para` — select the range and call
it. The flashing came from ACTIVATING another document, not from moving the selection within one.

### Group 3 — small caps and hyperlink addresses (3.0.173)

- [x] Replace Small Caps With All Caps — now `Sh_Replace_Small_Caps_With_All_Caps`. The whole
      body of both copies was ONE formatted find-and-replace; everything else was round-trip
      plumbing, so they were identical the moment it came out. Small caps are a look, not
      letters — the text underneath stays lower case, and a screen reader or DBT reads it as
      lower case — which is why this exists at all.
      The closing "turn small caps and all caps off" now runs on a COLLAPSED selection. That is
      what it always amounted to on the large-print side: it is there so the transcriber's next
      keystrokes are not in small caps, NOT to strip the capitals just applied. On the braille
      side it ran right after the paste, where the selection was the pasted text — a line that
      could have undone the macro's own work.
- [x] Convert Hyper To Addresses — `Lp_` and `Dx_` both converted, **not** merged. The
      large-print one turns the plain addresses back into live links (`Sh_Linkify_Range`); the
      braille one stops at plain text, because a DBT source file has no use for a clickable
      link. A real difference in what the transcriber gets, so it stays two — but it is Jerry's
      call, not ours: he settled the text-box pair the other way. Ask.
      `ActiveDocument.Hyperlinks` is a whole-document collection, so the scratch document was
      the scope here too.

Trap from this group: **rewrite hyperlinks backwards by index.** Writing `r.Text` takes the
link out of the collection AND changes the length of the text, so anything counted from the
front has moved by the next turn of the loop.

## Three faults Jerry's text-box test found (3.0.175)

All three were in `Sh_Remove_Txt_Bxs_And_Frames`, and none of them showed up in the first round
of headless tests because those tests all had either the whole document or a tidy text selection.

**A swallowed error deleted the box anyway.** The whole rescue block sat under one
`On Error Resume Next`, so a failed `FormattedText` assignment was passed over and the macro
carried straight on to `shp.Delete` — leaving a pair of markers with nothing between them and
the text gone for good. The shape is now deleted ONLY once the text has actually arrived; if it
has not, the opening marker is taken back out and the box is left where it is. Losing a text
box's contents is far worse than leaving the box for the transcriber to deal with.

**A selection inside a text box scoped the macro to nothing.** Clicking into a box makes
`Selection.Type = wdSelectionNormal` — same as selecting document text — but the range's story
is the box's own, so no shape's anchor could ever fall inside it and the macro silently did
nothing at all. A selection whose `StoryType` is not `wdMainTextStory` is now treated as no
selection.

**Off by one at the end of the range.** `Sh_Shape_In_Range` used `anchor.Start <= rng.End`. A
selection of one paragraph ENDS at the start of the next one, so a text box anchored to the
paragraph just below the selection was pulled in and converted. It is `< rng.End` now. Same fix
in `Sh_Frame_In_Range`.

Worth knowing about the route in: the Selection Cleanup forms run `Lp_Is_Text_Selected` before
every option, which stops with "Text must be selected first!" and `End` when the selection is
not document text — so with only shapes ctrl-clicked, that button cannot reach this macro at all.

## Still never tested: grouped shapes

Word refuses to group shapes programmatically in a headless instance ("Grouping is disabled for
the selected shapes"), so the `msoGroup` / `Ungroup` branch of `Sh_Remove_Txt_Bxs_And_Frames` has
never run under test. It is the braille code unchanged, but from 3.0.170 it is reached from the
LARGE-PRINT side too, where it had never run before. The case: two text boxes, grouped, cleaned
from the VistaType LP tab.

## The helpers underneath

These are what the buttons above reach. Converting a helper converts every button that uses it,
which is why the two Full File Cleanup buttons and the two Selection Cleanup buttons cover a
large part of the list between them.

`Lp_`/`Dx_Convert_Auto_List_To_Text`, `..._Convert_Hyper_To_Addresses`,
`..._Fix_Para_Space_Errors`, `..._Kill_The_Hyperlinks`, `..._Remove_Bullets`,
`..._Remove_Multi_Spaces`, `..._Remove_Txt_Bxs_And_Frames`, `..._Replace_Manual_Line_Break`,
`..._Replace_Multiple_Para_Marks_No_Warning`, `..._Replace_NonBreaking_Spaces`,
`..._Replace_Small_Caps_With_All_Caps`, `..._Replace_Tabs_With_Single_Space`,
`Lp_Remove_Box_Bullets_Bullets_and_Numbers`, `Lp_Resize_Images`, `Lp_TOC_CleanAndFormat_TOC`,
and the forms `Lp_Change_Image_Color_Form` and `Lp_Section_Brk_Caution`.

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

### Braille Macros tab

- [ ] Attach BANA Template — `Dx_Attach_BANA_Template`
- [ ] Full File Cleanup — `Dx_File_Fix_Sequence`
- [ ] Selection Clean Up — `Dx_Selected_File_CleanUp`
- [ ] AutoTag Ref Pages — `Dx_AutoTag_Page_Numbers`
- [ ] Validate $pg Tags — `Dx_Ref_Pg_Number_Sequence_Menu`
- [ ] Manual Tag Ref Page — `Dx_Manual_Tag_with_Dollar_pg`
- [ ] Format $pg Tags — `Dx_Format_Tagged_Page_Numbers`
- [ ] Embed Ref Pg Numb — `Dx_Embed_Ref_Pg_No`
- [ ] UnEmbed Ref Pg Numb — `Dx_UnEmbed_Ref_Pg_No`
- [ ] Foreign Lang in Color — `Dx_Add_Color_To_Foreign_Language_Words`
- [ ] Format Spelling List — `Dx_Spelling_List`
- [ ] Exercise Levels 1 & 2 — `Dx_Format_Exercise_Lv_1_and_Lv_2`
- [ ] Dashes/Primes/Fractions — `Dx_Type_Dashes`
- [ ] Compress Linear Math — `Dx_Compress_Linear_Math`
- [ ] DAISY or NIMAS to Word — `DN_Menu_Starter`

### VistaType LP tab

- [ ] Full File Cleanup — `Lp_File_Fix_Sequence`
- [ ] Selection Cleanup — `Lp_Selected_File_CleanUp`
- [ ] AutoTag Ref Pages — `Lp_AutoTag_Page_Numbers`
- [ ] Format Exercise — `Lp_Format_Exercise_Lv_1_and_Lv_2`
- [ ] Table and TOC Tools — `Lp_Table_Tools`
- [ ] Bkgrnd & Picture Tools — `Lp_Picture_Tools_Menu_Starter`
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

# Domain concepts

What the words mean: large print, braille and DBT, the typeface, reference page numbers,
file cleanup. Moved out of `CLAUDE.md` on 20 September 2026. Same text, no changes.

---

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

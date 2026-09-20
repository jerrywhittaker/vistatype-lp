# VistaType LP - Ribbon Test Checklist

**The single ribbon test document.** Every button on both tabs, the Quick Access Toolbar
icons and the retired-buttons tab, with the dialogs each one opens.

Built 7/26/2026 from `src/ribbon/customUI14.xml` and the UserForms, originally for the
"return the cursor where I left it" change - which is why the **RETURNS YOU** marking runs
through it. `LP-Ribbon-Test-List.md` was a second rendering of the same data covering only
the large-print half; it was merged in here and deleted on 9/20/2026, since two copies of
one inventory had already drifted apart.

> **This ought to be generated, not maintained by hand.** Every stale entry found on
> 9/20/2026 - a dialog that no longer exists, two QAT buttons that were never listed, a form
> named for the wrong tab - is the kind of drift `build_ribbon_dispatch.py` and
> `build_ribbon_tabs.py` already prevent for their own outputs. A
> `tools/lib/build_ribbon_test_list.py` beside them would end it. The dialog button lists
> below are also still truncated ("+22 more"), which a generator would fix for free.

## How to test the cursor fix

For every item marked **RETURNS YOU**, do this:

1. Open a long document (20+ pages) and click somewhere well down it - say page 12.
2. Note the page number and the paragraph you are sitting in.
3. Run the macro.
4. **You should be looking at that same paragraph again**, or within a line or two of it.
   The document may have got shorter, so a small drift is expected and fine.

If you land at the top of the document instead, press an arrow key. If the view jumps
back to page 12, the cursor was right and only the screen failed to follow - tell me,
that is a different fault from the cursor genuinely being lost.

Items marked *should not land at the top* are a weaker promise: the operation itself does
not save your place, but the cleanup steps inside it now do, so you should end up near
where you were rather than on page 1. Worth a look, but don't treat drift as a failure.

Items with no marking are either untouched by this work, or are meant to leave you
somewhere else on purpose:

| Left alone deliberately | Why |
|---|---|
| **Validate $pg Tags** (LP) | Stops at the tag it is complaining about - that is the point of it |
| **Export Sel to New Doc** / **Import to Current Doc** (both tabs) | Work on a selection; the import leaves you at what was inserted |
| **Move Paragraph to Next Page** (QAT) | You want to see where the paragraph landed |
| **Attach LP Template** / **Attach BANA Template** | These do a Save As, so the document is no longer the one your position was measured in |
| **Convert Paragraph to Title Case**, **Keep With Next**, **No Space After**, **Do Not Split** (QAT) | Point operations on the paragraph you are already sitting in - they never move you |

---


## Tab: Braille Macros


### Attach

- [ ] **Attach BANA Template**  `Dx_Attach_BANA_Template`  -- **RETURNS YOU**
  - opens dialog **Dx Choose BANA Template Form**
  - opens dialog **Dx Choose Translation Form**: Cmd BANA EBAE Button, Cmd BANA EBAE Nemeth Button, Cmd BANA UEB Button, Cmd BANA UEB Nemeth Button


### File Cleanup

- [ ] **Full File Cleanup**  `Dx_File_Fix_Sequence`  -- **RETURNS YOU**
  - opens dialog **Dx File Cleanup Sub Menu Form**: Fix Errors Button, Del Para Marks Button
- [ ] **Selection Clean Up**  `Dx_Selected_File_CleanUp`  -- **RETURNS YOU**
  - opens dialog **Dx Selected Cleanup Form**: Okay
- [ ] **Delete Prodnotes**  `Sh_Delete_Prodnote_Paragraphs`
- [ ] **Prodnote to TN**  `Dx_Change_Prodnotes_To_Transcriber_Notes`  -- **RETURNS YOU**


### Reference Page Number Formatting

- [ ] **AutoTag Ref Pages**  `Dx_AutoTag_Page_Numbers`  -- **RETURNS YOU**
  - opens dialog **Sh Validation Choices Form**: Navigation Button, Temp File Button, Show More Button
    - opens dialog **Sh Pg Validation Overview Form**
- [ ] **Validate $pg Tags**  `Dx_Ref_Pg_Number_Sequence_Menu`  -- *should not land at the top*
  - opens dialog **Sh Validation Choices Form**: Navigation Button, Temp File Button, Show More Button
    - opens dialog **Sh Pg Validation Overview Form**
- [ ] **Manual Tag Ref Page**  `Dx_Manual_Tag_with_Dollar_pg`  -- *should not land at the top*
- [ ] **Format $pg Tags**  `Dx_Format_Tagged_Page_Numbers`  -- *should not land at the top*
- [ ] **Embed Ref Pg Numb**  `Dx_Embed_Ref_Pg_No`  -- *should not land at the top*
- [ ] **UnEmbed Ref Pg Numb**  `Dx_UnEmbed_Ref_Pg_No`  -- *should not land at the top*


### Other Formatting Tools

- [ ] **Foreign Lang in Color**  `Dx_Add_Color_To_Foreign_Language_Words`  -- *should not land at the top*
- [ ] **Format Spelling List**  `Dx_Spelling_List`  -- *should not land at the top*
  - opens dialog **Dx Spelling List Options Form**: Cmd Ok
- [ ] **Horiz List to Vertical**  `Lp_Horz_List_To_Vertical`  -- **RETURNS YOU** (shared with the LP tab from 3.0.140)
  - opens dialog **Lp Horz To Vert List Form**: Cmd Ok  (the braille button uses the large-print form)
- [ ] **Exercise Levels 1 && 2**  `Dx_Format_Exercise_Lv_1_and_Lv_2`  -- **RETURNS YOU**
  - opens dialog **Dx UEB EBAE Fill In YN Form**: Cmd NO, Cmd YES
- [ ] **Dashes/Primes/Fractions**  `Dx_Type_Dashes`  -- *should not land at the top*
  - opens dialog **Dx Type Dashes Form**: Cmd Em, Cmd En, Cmd Encode, Cmd Long, Cmd Minus, Cmd Single, Cmd Double, Button Five Eighths ... (+20 more)
- [ ] **Compress Linear Math**  `Dx_Compress_Linear_Math`  -- *should not land at the top*


### Export/Import

- [ ] **Export Sel to New Doc**  `Dx_ExportSelectionToNewFile`
- [ ] **Import to Current Doc**  `Dx_Import_Exported_Selection_File`


### DAISY, NIMAS, and Text Tools

- [ ] **DAISY or NIMAS to Word**  `DN_Menu_Starter`
  - opens dialog **DN XML Type Form**, then **DN Keep Or Omit Images Form**
    (the manual route and `DN Auto Or Manual Form` were removed 9/17/2026)
- [ ] **Text File Para Fix**  `DN_Remove_Para_Formatting_From_Text_Files`
  - opens dialog **DN Text File Para Fix Warning**: Continue Button


### Help

- [ ] **Video Links && Practice Files**  `Dx_Video_Links`
  - opens dialog **Dx Video Download Link Page**: View the Page
- [ ] **Version && Updates**  `Dx_About`
  - opens dialog **Dx About Title And Agreement**: View Version


## Tab: VistaType LP


### Attach

- [ ] **Attach LP Template** &nbsp;`Lp_Attach_Lp_Template`
    - dialog **Lp Re Attach Warning Form**
        - [ ] Cmd No Button
        - [ ] Cmd Yes Button - "Yes Buttonon"
      - dialog **LP Attach An Lp Template Form**
          - [ ] Attach Okay
          - [ ] Mirrored Check Box
          - [ ] Attach Okay
          - [ ] Custom Orient Landscape
          - [ ] Custom Orient Portrait
          - [ ] Customize Check Box - "Customize paper or screen settings check box"
          - [ ] Margin Half Inch - "8.5 x 11 Paper - portrait orientation with 1/2 inch margin"
          - [ ] Margin Thee Fourths Inch - "8.5 x 11 Paper - portrait orientation with 3/4 inch margin"
          - [ ] Margin One Inch - "8.5 x 11 Paper - portrait orientation with 1 inch margin"
          - [ ] Nine Point Seven Tablet
          - [ ] ...and 20 more buttons on this dialog


### File Cleanup

- [ ] **Full File Cleanup** &nbsp;`Lp_File_Fix_Sequence`  **RETURNS YOU**
    - dialog **Lp File Cleanup Sub Menu Form**
        - [ ] Okay Button - "Okay Button"
- [ ] **Selection Cleanup** &nbsp;`Lp_Selected_File_CleanUp`  **RETURNS YOU**
    - dialog **Lp Selected Cleanup Form**
        - [ ] Okay Button
        - [ ] Compress Linear Math - "Compress the spacing in text math equations*" (9/15/2026: moved here off the ribbon; also runs in Fix Common File Errors)
- [ ] **Delete Prodnotes** &nbsp;`Sh_Delete_Prodnote_Paragraphs`


### Reference Page Numbers

- [ ] **AutoTag Ref Pages** &nbsp;`Lp_AutoTag_Page_Numbers`  **RETURNS YOU**
- [ ] **Validate $pg Tags** &nbsp;`Lp_Validate_Dollar_PG`
    - dialog **Sh Validation Choices Form**
        - [ ] Navigation Button
        - [ ] Temp File Button - "Validate by copying the tags to a temporary document"
        - [ ] Show More Button - "Show me more information"
      - dialog **Sh Pg Validation Overview Form**
          - [ ] Close Button
- [ ] **Manual Tag Ref Page** &nbsp;`Lp_Manual_Tag_with_Dollar_pg`
- [ ] **Format $pg Tags** &nbsp;`Lp_Format_Page_Numbers`  **RETURNS YOU**


### Other Formatting Tools

- [ ] **Format Exercise** &nbsp;`Lp_Format_Exercise_Lv_1_and_Lv_2`  **RETURNS YOU**
- [ ] **Fill-In Line** &nbsp;`Lp_Type_Fill_In_Line`
    - dialog **Lp Type Fill In Line Form**
        - [ ] Command Button1
        - [ ] Command Button2 - "Type 2 fill-in characters"
        - [ ] Command Button3 - "Type 3 fill-in characters"
        - [ ] Command Button4 - "Type 4 fill-in characters"
        - [ ] Command Button5 - "Type 5 fill-in characters"
        - [ ] Command Button6 - "Type 6 fill-in characters"
        - [ ] Command Button7 - "Type 7 fill-in characters"
        - [ ] Command Button8 - "Type 8 fill-in characters"
        - [ ] Command Button9 - "Type 9 fill-in charactersM"
        - [ ] Command Button10 - "Type 10 fill-in characters"
        - [ ] ...and 12 more buttons on this dialog
- [ ] **Horiz List to Vertical** &nbsp;`Lp_Horz_List_To_Vertical`  **RETURNS YOU**
    - dialog **Lp Horz To Vert List Form**
        - [ ] Cmd Ok
- [ ] **Table and TOC Tools** &nbsp;`Lp_Table_Tools`  **RETURNS YOU**
    - dialog **Lp Table Tools Menu Form**
        - [ ] Default Color Selected
        - [ ] Yellow Color Selected
        - [ ] Gray Color Selected
        - [ ] Black And White Selected - "Change selected table to plain black and white."
        - [ ] List Button
        - [ ] Pseudo Button - "Pseudo Columns: A Table With no Borders"
        - [ ] Pseudo Question - "Information about pseudo columns button"
        - [ ] Real Button - "Convert Table to Real Columns Button"
        - [ ] Real Question
        - [ ] Rotate Button - "Rotate Table 90 Degrees  button"
        - [ ] ...and 4 more buttons on this dialog
      - dialog **Lp Table Convert Options Form**
          - [ ] Cmd Okay - "Okay ButtonB"
          - [ ] RCTable Image Button
          - [ ] Column Only Table Image Button - "Table has column headers only"
          - [ ] Row Only Image Button
          - [ ] No Row Column Image Button - "Table has row headers only"
          - [ ] Row And Column Table - "Table has row and columns headings radio button"
          - [ ] Column Only Table - "Tabel has column heading only radio button3"
          - [ ] Row Only Table - "Table has row headings only radio button"
          - [ ] No Row Or Column
          - [ ] Use Colors
          - [ ] ...and 6 more buttons on this dialog
      - dialog **LP Pseudo Column Info Form**
          - [ ] Command Button1
      - dialog **LP Real Column Info Form**
          - [ ] Command Button1
    - dialog **Lp TOC Format And Color Form**
        - [ ] Okay Button
      - dialog **Lp TOC Color Bars Form**
          - [ ] Yellow Image Button
          - [ ] Pink Image Button
          - [ ] Tan Image Button
          - [ ] Blue Image Button
          - [ ] Aqua Image Button
          - [ ] Green Image Button
          - [ ] Okay Button
- [ ] **Bkgrnd && Picture Tools** &nbsp;`Lp_Picture_Tools_Menu_Starter`  **RETURNS YOU**
    - dialog **Lp Bakgrnd Picture Menu Form**
        - [ ] Background Color Button - "Document Backgroud Color Button"
        - [ ] Image Size Button - "Resize Pictures Buttonor"
        - [ ] Pictures To Inline Button - "     All Pictures to Inline with fixed aspect ratio"
        - [ ] Pictures Color Grayscale Button - "Pictures Color/Grayscale Button"
        - [ ] Center Or Left Align Pictures Button - "Center or Left Align Pictures Button"
        - [ ] Same Picture Size Button - "Resize This Picture Throughout the Book" (9/15/2026: select one picture, set its size, press it; every other copy of that picture takes the same size)
      - dialog **Lp Resize Images Form**
          - [ ] Cmd Okay
      - dialog **Lp Change Image Color Form**
          - [ ] Okay Button
      - dialog **LP Picture Alignment Form**
          - [ ] Cmd Okay


### Export/Import

- [ ] **Export Sel to New Doc** &nbsp;`Lp_Export_Selection_To_NewFile`
- [ ] **Import to Current Doc** &nbsp;`Lp_Import_Exported_Selection_File`


### DAISY, NIMAS and Text Tools

- [ ] **DAISY or NIMAS to Word** &nbsp;`DN_Menu_Starter`
    - dialog **DN XML Type Form**
        - [ ] (the manual route and `DN Auto Or Manual Form` were removed 9/17/2026;
              `DN_Menu_Starter` now calls `Sh_Convert_XML_File_To_Word_Document` directly)
    - dialog **DN Keep Or Omit Images Form**
- [ ] **Text File Para Fix** &nbsp;`DN_Remove_Para_Formatting_From_Text_Files`
    - dialog **DN Text File Para Fix Warning**
        - [ ] Continue Button - "Continue Button"


### Help

- [ ] **Videos and Practice Files** &nbsp;`Lp_Video_Links`
    - dialog **Lp Video Download Link Page**
        - [ ] View the Page - "View/Download Link Page"
- [ ] **Version && Updates** &nbsp;`Lp_About`
    - dialog **Lp About Title And Agreement**
        - [ ] Label6
        - [ ] View Version


## Tab: LP and BRL QAT Icons

> Hidden by design (`visible="false"`) — it will not appear on the ribbon. It exists only
> so the Quick Access Toolbar's entries have ribbon controls to resolve against. Test these
> **eight** from the Quick Access Toolbar (or by keyboard shortcut), not from a tab.
> Document Settings is Ctrl+Alt+Shift+I.

### QAT icons

- [ ] **Document Settings**  `Sh_Doc_Info`
- [ ] **No Space After this Paragraph**  `Lp_Toggle_Space_After_Current_Para`
- [ ] **Keep Paragraph With Next Paragraph**  `Lp_Keep_With_Next_Para`
- [ ] **Move Paragraph to Next Page**  `Sh_Move_Paragraph_To_Next_Page`
- [ ] **Do Not Split Paragraph Across Pages**  `Sh_Keep_Lines_Of_Para_Together`
- [ ] **Convert Paragraph to Title Case**  `Sh_Apply_Title_Case_Capitalization`
- [ ] **Styles Pane: Recommended**  `Sh_Show_Recommended_Styles_Pane`
- [ ] **Reset Word Configuration**  `MS_Reset_Word_Configuration`

## Tab: VistaType LP Retired Buttons

> Also hidden. A button is moved here rather than deleted, because every toolbar already
> installed in the field references it by id and a deleted id draws as a blank button.
> **The test is that a retired button still works when pressed from an old toolbar.**

- [ ] **Compress Linear Math** (large print)  `Lp_Compress_Linear_Math` — retired 9/15/2026

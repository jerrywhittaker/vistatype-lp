# VistaType LP - Ribbon Test Checklist

Generated 7/26/2026, for the "return the cursor where I left it" change.

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
  - opens dialog **Dx Horz To Vert List Form**: Cmd Ok
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
  - opens dialog **DN Auto Or Manual Form**: Automatic Button, Manual Button
- [ ] **Text File Para Fix**  `DN_Remove_Para_Formatting_From_Text_Files`
  - opens dialog **DN Text File Para Fix Warning**: Continue Button


### Help

- [ ] **Video Links && Practice Files**  `Dx_Video_Links`
  - opens dialog **Dx Video Download Link Page**: View the Page
- [ ] **Version && Updates**  `Dx_About`
  - opens dialog **Dx About Title And Agreement**: View Version


## Tab: VistaType LP


### Attach

- [ ] **Attach LP Template**  `Lp_Attach_Lp_Template`
  - opens dialog **Lp Re Attach Warning Form**: Cmd No Button, Cmd Yes Button
    - opens dialog **LP Attach An Lp Template Form**: Attach Okay, Mirrored Check Box, Attach Okay, Custom Orient Landscape, Custom Orient Portrait, Customize Check Box, Margin Half Inch, Margin Thee Fourths Inch ... (+22 more)


### File Cleanup

- [ ] **Full File Cleanup**  `Lp_File_Fix_Sequence`  -- **RETURNS YOU**
  - opens dialog **Lp File Cleanup Sub Menu Form**: Okay Button
- [ ] **Selection Cleanup**  `Lp_Selected_File_CleanUp`  -- **RETURNS YOU**
  - opens dialog **Lp Selected Cleanup Form**: Okay Button, Compress Linear Math (moved here off the ribbon 9/15/2026)
- [ ] **Delete Prodnotes**  `Sh_Delete_Prodnote_Paragraphs`


### Reference Page Numbers

- [ ] **AutoTag Ref Pages**  `Lp_AutoTag_Page_Numbers`  -- **RETURNS YOU**
- [ ] **Validate $pg Tags**  `Lp_Validate_Dollar_PG`
  - opens dialog **Sh Validation Choices Form**: Navigation Button, Temp File Button, Show More Button
    - opens dialog **Sh Pg Validation Overview Form**
- [ ] **Manual Tag Ref Page**  `Lp_Manual_Tag_with_Dollar_pg`
- [ ] **Format $pg Tags**  `Lp_Format_Page_Numbers`  -- **RETURNS YOU**


### Other Formatting Tools

- [ ] **Format Exercise**  `Lp_Format_Exercise_Lv_1_and_Lv_2`  -- **RETURNS YOU**
- [ ] **Fill-In Line**  `Lp_Type_Fill_In_Line`
  - opens dialog **Lp Type Fill In Line Form**: Command Button1, Command Button2, Command Button3, Command Button4, Command Button5, Command Button6, Command Button7, Command Button8 ... (+14 more)
- [ ] **Horiz List to Vertical**  `Lp_Horz_List_To_Vertical`  -- **RETURNS YOU**
  - opens dialog **Lp Horz To Vert List Form**: Cmd Ok
- [ ] **Table and TOC Tools**  `Lp_Table_Tools`  -- **RETURNS YOU**
  - opens dialog **Lp TOC Format And Color Form**: Okay Button
    - opens dialog **Lp TOC Color Bars Form**: Yellow Image Button, Pink Image Button, Tan Image Button, Blue Image Button, Aqua Image Button, Green Image Button, Okay Button
  - opens dialog **Lp Table Tools Menu Form**: Default Color Selected, Yellow Color Selected, Gray Color Selected, Black And White Selected, List Button, Pseudo Button, Pseudo Question, Real Button ... (+6 more)
    - opens dialog **LP Pseudo Column Info Form**: Command Button1
    - opens dialog **LP Real Column Info Form**: Command Button1
    - opens dialog **Lp Table Convert Options Form**: Cmd Okay, RCTable Image Button, Column Only Table Image Button, Row Only Image Button, No Row Column Image Button, Row And Column Table, Column Only Table, Row Only Table ... (+8 more)
- [ ] **Bkgrnd && Picture Tools**  `Lp_Picture_Tools_Menu_Starter`  -- **RETURNS YOU**
  - opens dialog **Lp Bakgrnd Picture Menu Form**: Background Color Button, Image Size Button, Pictures To Inline Button, Pictures Color Grayscale Button, Center Or Left Align Pictures Button, Same Picture Size Button (9/15/2026)
    - opens dialog **LP Picture Alignment Form**: Cmd Okay
    - opens dialog **Lp Change Image Color Form**: Okay Button
    - opens dialog **Lp Resize Images Form**: Cmd Okay


### Export/Import

- [ ] **Export Sel to New Doc**  `Lp_Export_Selection_To_NewFile`
- [ ] **Import to Current Doc**  `Lp_Import_Exported_Selection_File`


### DAISY, NIMAS and Text Tools

- [ ] **DAISY or NIMAS to Word**  `DN_Menu_Starter`
  - opens dialog **DN Auto Or Manual Form**: Automatic Button, Manual Button
- [ ] **Text File Para Fix**  `DN_Remove_Para_Formatting_From_Text_Files`
  - opens dialog **DN Text File Para Fix Warning**: Continue Button


### Help

- [ ] **Videos and Practice Files**  `Lp_Video_Links`
  - opens dialog **Lp Video Download Link Page**: View the Page
- [ ] **Version && Updates**  `Lp_About`
  - opens dialog **Lp About Title And Agreement**: View Version


## Tab: LP and BRL QAT Icons

> Hidden by design (`visible="false"`) — it will not appear on the ribbon. It exists only
> so the Quick Access Toolbar's entries have ribbon controls to resolve against. Test these
> six from the **Quick Access Toolbar** (or by keyboard shortcut), not from a tab.

### (no label - QAT icons)

- [ ] **Document Settings**  `Sh_Doc_Info`
- [ ] **No Space After this Paragraph**  `Lp_Toggle_Space_After_Current_Para`
- [ ] **Keep Paragraph With Next Paragraph**  `Lp_Keep_With_Next_Para`
- [ ] **Move Paragraph to Next Page**  `Sh_Move_Paragraph_To_Next_Page`
- [ ] **Do Not Split Paragraph Across Pages**  `Sh_Keep_Lines_Of_Para_Together`
- [ ] **Convert Paragraph to Title Case**  `Sh_Apply_Title_Case_Capitalization`

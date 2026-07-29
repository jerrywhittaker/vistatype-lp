# LP Ribbon - Test List

Built 7/26/2026 from `src/ribbon/customUI14.xml` + the UserForms. The **VistaType LP**
tab and the QAT icons. Braille tab is in `Ribbon-Test-Checklist.md`.

**RETURNS YOU** = this operation should put the cursor back where it was. Sit on page 12
of a long document, note the paragraph, run it, check you come back.

Dialog buttons are listed by the name in the code, with the on-screen wording in
quotes where I could read it out of the form's `.frx`.
---


## VistaType LP


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
- [ ] **Compress Linear Math** &nbsp;`Lp_Compress_Linear_Math`  **RETURNS YOU**
- [ ] **Bkgrnd && Picture Tools** &nbsp;`Lp_Picture_Tools_Menu_Starter`  **RETURNS YOU**
    - dialog **Lp Bakgrnd Picture Menu Form**
        - [ ] Background Color Button - "Document Backgroud Color Button"
        - [ ] Image Size Button - "Resize Pictures Buttonor"
        - [ ] Pictures To Inline Button - "     All Pictures to Inline with fixed aspect ratio"
        - [ ] Pictures Color Grayscale Button - "Pictures Color/Grayscale Button"
        - [ ] Center Or Left Align Pictures Button - "Center or Left Align Pictures Button"
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
    - dialog **DN Auto Or Manual Form**
        - [ ] Automatic Button - "Convert the XML File to Word"
        - [ ] Manual Button
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


## LP and BRL QAT Icons

> This tab is **hidden** (`visible="false"`) and will not appear on the ribbon — that is
> correct. It exists only so the Quick Access Toolbar entries have ribbon controls to point
> at. Test these six **from the Quick Access Toolbar**, not from a tab. They also have
> keyboard shortcuts (Document Settings is Ctrl+Alt+Shift+I).

### QAT icons

- [ ] **Document Settings** &nbsp;`Sh_Doc_Info`
- [ ] **No Space After this Paragraph** &nbsp;`Lp_Toggle_Space_After_Current_Para`
- [ ] **Keep Paragraph With Next Paragraph** &nbsp;`Lp_Keep_With_Next_Para`
- [ ] **Move Paragraph to Next Page** &nbsp;`Sh_Move_Paragraph_To_Next_Page`
- [ ] **Do Not Split Paragraph Across Pages** &nbsp;`Sh_Keep_Lines_Of_Para_Together`
- [ ] **Convert Paragraph to Title Case** &nbsp;`Sh_Apply_Title_Case_Capitalization`

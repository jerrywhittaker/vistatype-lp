Attribute VB_Name = "RibbonDispatch"
' VistaType LP - generated ribbon dispatch table.
'
' ############################################################################
' ##  GENERATED FILE - DO NOT EDIT.                                         ##
' ##  Written by tools/lib/build_ribbon_dispatch.py from                    ##
' ##  src/ribbon/customUI14.xml on every `make build`. Change the ribbon,   ##
' ##  not this file; anything typed here is lost on the next build.         ##
' ############################################################################
'
' WHY A SELECT CASE AND NOT Application.Run
'
' RibbonAction used to reach every macro with `Application.Run control.Tag`, and wrapped it in
' On Error GoTo so one handler could report a failure in any of the 47 buttons.
'
' That does not work. WORD DOES NOT PASS AN ERROR BACK OUT OF Application.Run. It takes the
' error itself and shows its own "Run-time error" dialog - the one with End and Debug, which
' also offers to open this source on the user's machine - and the caller's handler is never
' entered. Measured on the build box 8/26/2026: four real button presses wrote the marker set
' immediately BEFORE the call and never the one immediately after it, nor the one first thing
' in the handler.
'
' A DIRECT call propagates normally, which is what this table provides. Every macro named by a
' button's tag is a plain no-argument Sub, which is what makes the table possible at all; the
' generator checks that and refuses to write a table if it ever stops being true.
'
Option Explicit

' True when macroName was found and run. False means it is not on the ribbon, and the caller
' decides what to do about that - RibbonAction falls back to Application.Run, so a button can
' never do nothing at all just because this file is out of date.
'
Public Function Sh_Dispatch(ByVal macroName As String) As Boolean
    Sh_Dispatch = True

    Select Case macroName
        Case "DN_Menu_Starter": DN_Menu_Starter
        Case "DN_Remove_Para_Formatting_From_Text_Files": DN_Remove_Para_Formatting_From_Text_Files
        Case "Dx_About": Dx_About
        Case "Dx_Add_Color_To_Foreign_Language_Words": Dx_Add_Color_To_Foreign_Language_Words
        Case "Dx_Attach_BANA_Template": Dx_Attach_BANA_Template
        Case "Dx_AutoTag_Page_Numbers": Dx_AutoTag_Page_Numbers
        Case "Dx_Change_Prodnotes_To_Transcriber_Notes": Dx_Change_Prodnotes_To_Transcriber_Notes
        Case "Dx_Compress_Linear_Math": Dx_Compress_Linear_Math
        Case "Dx_Embed_Ref_Pg_No": Dx_Embed_Ref_Pg_No
        Case "Dx_ExportSelectionToNewFile": Dx_ExportSelectionToNewFile
        Case "Dx_File_Fix_Sequence": Dx_File_Fix_Sequence
        Case "Dx_Format_Exercise_Lv_1_and_Lv_2": Dx_Format_Exercise_Lv_1_and_Lv_2
        Case "Dx_Format_Tagged_Page_Numbers": Dx_Format_Tagged_Page_Numbers
        Case "Dx_Import_Exported_Selection_File": Dx_Import_Exported_Selection_File
        Case "Dx_Manual_Tag_with_Dollar_pg": Dx_Manual_Tag_with_Dollar_pg
        Case "Dx_Ref_Pg_Number_Sequence_Menu": Dx_Ref_Pg_Number_Sequence_Menu
        Case "Dx_Selected_File_CleanUp": Dx_Selected_File_CleanUp
        Case "Dx_Spelling_List": Dx_Spelling_List
        Case "Dx_Type_Dashes": Dx_Type_Dashes
        Case "Dx_UnEmbed_Ref_Pg_No": Dx_UnEmbed_Ref_Pg_No
        Case "Dx_Video_Links": Dx_Video_Links
        Case "Lp_About": Lp_About
        Case "Lp_Attach_Lp_Template": Lp_Attach_Lp_Template
        Case "Lp_AutoTag_Page_Numbers": Lp_AutoTag_Page_Numbers
        Case "Lp_Compress_Linear_Math": Lp_Compress_Linear_Math
        Case "Lp_Export_Selection_To_NewFile": Lp_Export_Selection_To_NewFile
        Case "Lp_File_Fix_Sequence": Lp_File_Fix_Sequence
        Case "Lp_Format_Exercise_Lv_1_and_Lv_2": Lp_Format_Exercise_Lv_1_and_Lv_2
        Case "Lp_Format_Page_Numbers": Lp_Format_Page_Numbers
        Case "Lp_Horz_List_To_Vertical": Lp_Horz_List_To_Vertical
        Case "Lp_Import_Exported_Selection_File": Lp_Import_Exported_Selection_File
        Case "Lp_Keep_With_Next_Para": Lp_Keep_With_Next_Para
        Case "Lp_Manual_Tag_with_Dollar_pg": Lp_Manual_Tag_with_Dollar_pg
        Case "Lp_Picture_Tools_Menu_Starter": Lp_Picture_Tools_Menu_Starter
        Case "Lp_Selected_File_CleanUp": Lp_Selected_File_CleanUp
        Case "Lp_Table_Tools": Lp_Table_Tools
        Case "Lp_Toggle_Space_After_Current_Para": Lp_Toggle_Space_After_Current_Para
        Case "Lp_Type_Fill_In_Line": Lp_Type_Fill_In_Line
        Case "Lp_Validate_Dollar_PG": Lp_Validate_Dollar_PG
        Case "Lp_Video_Links": Lp_Video_Links
        Case "MS_Reset_Word_Configuration": MS_Reset_Word_Configuration
        Case "Sh_Apply_Title_Case_Capitalization": Sh_Apply_Title_Case_Capitalization
        Case "Sh_Delete_Prodnote_Paragraphs": Sh_Delete_Prodnote_Paragraphs
        Case "Sh_Doc_Info": Sh_Doc_Info
        Case "Sh_Keep_Lines_Of_Para_Together": Sh_Keep_Lines_Of_Para_Together
        Case "Sh_Move_Paragraph_To_Next_Page": Sh_Move_Paragraph_To_Next_Page
        Case "Sh_Show_Recommended_Styles_Pane": Sh_Show_Recommended_Styles_Pane
        Case Else
            Sh_Dispatch = False
    End Select
End Function   '*** end of Sh_Dispatch ***

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Table_Convert_Options_Form 
   Caption         =   "Large Print Table Conversion Options"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9585.001
   OleObjectBlob   =   "Lp_Table_Convert_Options_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Table_Convert_Options_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Lp_Table_Convert_Options_Form
'
' Author: Jerry Whittaker - jerry@vistatypelp.org
'
' Version: 2.0  Date: 9/3/2026  - NO SCRATCH DOCUMENT. Both conversions run on the table where it
'                                 stands in the transcriber's own book, so nothing is shown on the
'                                 screen, nothing goes through the clipboard, no document is
'                                 activated, and each conversion is one Ctrl+Z instead of two or
'                                 four. The deliberate minimize/maximize of the Word window at the
'                                 end of a rotation - "the Jolt" - is gone with it
' Version: 1.9  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.8  Date: 6/22/2026 - ajusted timing with "do events" to make sure that the original document is shown at the end of the rotation
' Version: 1.7  Date: 5/5/2026  - fixed bug in table rotation that prevented temp files from being closed
' Version: 1.6  Date: 1/28/2026 - revised import code into original doc from temp doc
' Version: 1.5  Date: 9/15/2025 - bypassed type Z tables by making them Y type tables
' Version: 1.4  Date: 8/13/2025 - default header color ajustments - added Application.Run MacroName:="Lp_Table_Insert_Transcriber_Note"
'                                 replaced sleep with DoEvents
' Version: 1.3  Date: 7/25/2025 - limited color choices based on value UseColors
' Version: 1.2  Date: 6/27/2025 - added table no row or column headers table type - fixed minor button bugs for table to list
' Version: 1.1  Date: 11/7/2023 - fixed form label for "Table Has Row & Column Headers"
' Version: 1.0  Date: 12/5/2019
'
' Controls the options used by:
'          Lp_Table_Convert_RC_Table_To_List and Lp_Table_Convert_R_Only_Table_To_List
'

Private Sub CmdCancel_Click()
    Unload Lp_Table_Tools_Menu_Form
    Unload Me
    End
End Sub

Private Sub CmdOkay_Click()

    Dim su_Prev As Boolean
    Dim workTbl As Table

    ' NO CUSTOM UNDO RECORD, AND THESE TWO CONVERSIONS MUST NEVER HAVE ONE. 3.0.345 and 3.0.347
    ' each opened one so that Ctrl+Z would take the whole job back in a single press, and each
    ' CRASHED WORD OUTRIGHT - Jerry, testing 3.0.347: "converting table to a list and transposing
    ' (rotating) a table, both crashed word". It is the Word defect measured on the build box on
    ' 9/2/2026 and written up at the top of Lp_TOC_CleanAndFormat_TOC: a Find with
    ' Replace:=wdReplaceAll inside an open StartCustomRecord takes Word down with an access
    ' violation in wwlib.dll, nothing raised and nothing logged.
    '
    ' None is needed. The work is done in a scratch document that is never shown, and comes home
    ' in two steps - delete the table, put the result where it stood. Ctrl+Z is two presses, the
    ' same as the old list conversion and better than the old rotation's four. See
    ' Lp_Table_Convert_Hidden for why it cannot be one.
    '
    ' The handler below is not optional. The form calls its macros directly rather than through
    ' RibbonAction, so without it a failure reaches the transcriber as Word's own Run-time error
    ' dialog, Debug button and all - and Debug opens this source on their machine.
    On Error GoTo ConversionFailed

    su_Prev = Application.ScreenUpdating

    Lp_Table_Convert_Options_Form.Hide
    
    ' Lp_GP_String_3 is concatinated with the following codes

        ' 1=Black (wdColorAutomatic)
        ' 2=Red
        ' 3=Orange
        ' 4=Blue
        ' 5=Violet
        ' 6=Green
        ' B=LeaveBlankCellsBlank
        ' C=UseColors for headers
        ' E=Create an Empty Para After List Groups
        ' H=Display Headers In Bold
        ' L=Make table into a list a list - set in Lp_Table_Tools_Menu_Form
        ' M=Fill Blank Cells With EMDash
        ' N=Fill Blank Cells With "NA"
        ' O=Headers in Lower Case
        ' P=Headers in Title Case
        ' R=rotate this table - set in Lp_Table_Tools_Menu_Form
        ' S=Keep List Groups On Same Page
        ' U=Headers in Upper Case
        ' W=a table with no row or column headers
        ' X=a row and column table
        ' Y=a column header only table
        ' Z=a row header only table

    '****************** control string settings ****************
    
    If KeepListGroupsOnSamePage Then
        Lp_GP_String_3 = Lp_GP_String_3 + "S "
    End If
    
    If EmptyParaAfterLists Then
        Lp_GP_String_3 = Lp_GP_String_3 + "E "
    End If
    
   If TitleCaseRadioButton Then
        Lp_GP_String_3 = Lp_GP_String_3 + "P "
    End If
    
    If DisplayHeadersInBold Then
        Lp_GP_String_3 = Lp_GP_String_3 + "H "
    End If
    
    If UseColors Then
        Lp_GP_String_3 = Lp_GP_String_3 + "C "
    End If
    
    If Black Then
        Lp_GP_String_3 = Lp_GP_String_3 + "1 "
    End If
    
    If Red Then
        Lp_GP_String_3 = Lp_GP_String_3 + "2 "
    End If
    
    If Orange Then
        Lp_GP_String_3 = Lp_GP_String_3 + "3 "
    End If
    
    If Blue Then
        Lp_GP_String_3 = Lp_GP_String_3 + "4 "
    End If
    
    If Violet Then
        Lp_GP_String_3 = Lp_GP_String_3 + "5 "
    End If
    
    If Green Then
        Lp_GP_String_3 = Lp_GP_String_3 + "6 "
    End If
    
    If ProperCaseRadioButton Then
        Lp_GP_String_3 = Lp_GP_String_3 + "P "
    End If
    
    If FillBlankCellsWithEMDash Then
        Lp_GP_String_3 = Lp_GP_String_3 + "M "
    End If
    
    If FillBlankCellsWithNA Then
        Lp_GP_String_3 = Lp_GP_String_3 + "N "
    End If

    If LeaveBlankCellsBlank Then
        Lp_GP_String_3 = Lp_GP_String_3 + "B "
    End If
    
    If HeadersInLowerCaseRadioButton Then
        Lp_GP_String_3 = Lp_GP_String_3 + "O "
    End If
    
    If HeadersInAllCapsRadioButton Then
        Lp_GP_String_3 = Lp_GP_String_3 + "U "
    End If
    
    If NoRowOrColumn Then
        Lp_GP_String_3 = Lp_GP_String_3 + "W "  ' This is table without Row and column headers
    End If
    
    If RowAndColumnTable Then
        Lp_GP_String_3 = Lp_GP_String_3 + "X "   ' This is a table with row and column headers
    End If

    If ColumnOnlyTable Then
        Lp_GP_String_3 = Lp_GP_String_3 + "Y "  ' This is a table with column headers only
    End If
    
    If RowOnlyTable Then
        Lp_GP_String_3 = Lp_GP_String_3 + "Z "  ' This is a table with Row headers only
    End If

'****************** run the selected conversion **************************************
'
' 9/3/2026 - Both conversions now do their work in a scratch document that is NEVER SHOWN, and
' bring the result home in a single assignment. See Lp_Table_Convert_Hidden for the whole story;
' in short, the scratch document does two jobs and only one of them was ever the fault. It had to
' be SHOWN because the passes reached their table through Selection - that is the flash of yellow
' - and the passes now take the table instead. But it is also what keeps the undo short, because
' edits made in another document are not in this book's undo stack at all. Taking the round trip
' out altogether made Ctrl+Z 30 to 50 presses, and a custom undo record to cure that crashed Word.

'***************** List Table ******************************

    If InStr(Lp_GP_String_3, "L") > 0 Then

        If Not Selection.Information(wdWithInTable) Then
            Sh_Say "Place the cursor in the table first.", "VistaType LP (288)"
            Unload Lp_Table_Tools_Menu_Form
            Unload Me
            End
        End If
        Set workTbl = Selection.Tables(1)

        ' A row-and-column list needs a header in row 1, column 1, and this stops with a message
        ' if there is not one. Asked before anything is changed.
        If InStr(Lp_GP_String_3, "X") > 0 Then
            Lp_Table_Is_R1C1_Empty workTbl
        End If

        ' The row-only and column-only lists are made by rotating the table, so a merged one
        ' cannot be done - and the transcriber must hear that BEFORE anything is changed.
        If InStr(Lp_GP_String_3, "Y") > 0 Or InStr(Lp_GP_String_3, "Z") > 0 Then
            If Not Lp_Table_Is_Uniform_Grid(workTbl) Then
                Sh_Say "This table has merged cells. Unmerge them before converting it to a list.", "VistaType LP (287)"
                Unload Lp_Table_Tools_Menu_Form
                Unload Me
                End
            End If
        End If

        Application.ScreenUpdating = False

        Lp_Table_Convert_Hidden workTbl, False

        ' The find parameters are put back, but NOT through
        ' MS_Clear_F_and_R_Params_and_Clipboard: nothing here uses the clipboard any more, so
        ' emptying it would throw away whatever the transcriber had copied for no reason.
        Sh_Reset_Find_Parameters

        Application.ScreenUpdating = su_Prev
        ActiveWindow.ScrollIntoView Selection.Range
        Application.ScreenRefresh

        Sh_Say "Press Ctrl+Z two times to restore the original table.", "VistaType (178)"
        Unload Lp_Table_Tools_Menu_Form
        Unload Me
        End
    End If

'************ Rotate (transpose) Table ****************

    If InStr(Lp_GP_String_3, "R") > 0 Then

        If Not Selection.Information(wdWithInTable) Then
            Sh_Say "Place the cursor in the table first.", "VistaType LP (288)"
            Unload Lp_Table_Tools_Menu_Form
            Unload Me
            End
        End If
        Set workTbl = Selection.Tables(1)

        ' Asked FIRST, so that a table that cannot be rotated is never changed at all.
        If Not Lp_Table_Is_Uniform_Grid(workTbl) Then
            Sh_Say "This table has merged cells. Unmerge them before rotating the table.", "VistaType LP (287)"
            Unload Lp_Table_Tools_Menu_Form
            Unload Me
            End
        End If

        Application.ScreenUpdating = False

        Lp_Table_Convert_Hidden workTbl, True

        ' See the note in the list branch: find parameters back, clipboard left alone.
        Sh_Reset_Find_Parameters

        Application.ScreenUpdating = su_Prev
        ActiveWindow.ScrollIntoView Selection.Range
        Application.ScreenRefresh

        Sh_Say "Press Ctrl+Z two times to restore the original table.", "VistaType LP (200)"

        If InStr(Lp_GP_String_3, "X") > 0 Then
            Sh_Say "Check the header in column 1 row 1. The title should describe the data in the column.", "VistaType LP (202)"
        End If

        Unload Lp_Table_Tools_Menu_Form
        Unload Me
    End If

    Exit Sub

ConversionFailed:
    ' Put the screen back BEFORE saying anything, then report through the add-in's own dialog
    ' rather than Word's. Nothing in here may raise.
    Dim failNumber As Long
    Dim failText As String

    failNumber = Err.Number
    failText = Err.Description

    On Error Resume Next
    Sh_Reset_Find_Parameters
    Application.ScreenUpdating = True
    Application.ScreenRefresh

    ' Nothing here opens a custom undo record and nothing saves a cursor position any more -
    ' Lp_Table_Convert_Hidden closes its own scratch document on the way out - so the screen is
    ' the only thing to put back.
    Sh_Report_Error "Lp_Table_Convert_Options_Form", failNumber, failText

    Unload Lp_Table_Tools_Menu_Form
    Unload Me

End Sub

Private Sub RCTableImageButton_Click()
    CmdOkay.Enabled = True
    OptFrame.Visible = True
    HowToDisplayEmptyCellsFrame.Visible = True
    RowAndColumnTable.Value = True
End Sub

Private Sub ColumnOnlyTableImageButton_Click()
    CmdOkay.Enabled = True
    OptFrame.Visible = True
    HowToDisplayEmptyCellsFrame.Visible = True
    ColumnOnlyTable.Value = True
End Sub

Private Sub RowOnlyImageButton_Click()
    CmdOkay.Enabled = True
    OptFrame.Visible = True
    HowToDisplayEmptyCellsFrame.Visible = True
    RowOnlyTable.Value = True
End Sub

Private Sub NoRowColumnImageButton_Click()
    CmdOkay.Enabled = True
    OptFrame.Visible = True
    HowToDisplayEmptyCellsFrame.Visible = True
    LeaveBlankCellsBlank.Enabled = False
    LeaveBlankCellsBlank.Visible = True
    NoRowOrColumn.Value = True
End Sub
'********* End Table Type Image Buttons ********

'*************** Begin Table Type Radio Buttons *****************
Private Sub RowAndColumnTable_Click() 'radio button
    ListGroupSettingsFrame.Visible = True
    OutputListOptionsFrame.Visible = True
    CmdOkay.Enabled = True
    OptFrame.Visible = True
    HowToDisplayEmptyCellsFrame.Visible = True
    
    If InStr(Lp_GP_String_3, "L") > 0 Then
        UseColors.Value = True
        UseColors.Enabled = True
        DisplayHeadersInBold.Enabled = True
        LeaveBlankCellsBlank.Visible = False
        If FillBlankCellsWithEMDash = True Then
        FillBlankCellsWithNA.Value = False
        End If
    End If
    
    If InStr(Lp_GP_String_3, "R") > 0 Then
        FillBlankCellsWithNA.Value = True
        LeaveBlankCellsBlank.Enabled = True
        LeaveBlankCellsBlank.Visible = True
        LeaveBlankCellsBlank.Value = True
    End If
End Sub

Private Sub ColumnOnlyTable_Click() 'radio button
    ListGroupSettingsFrame.Visible = True
    OutputListOptionsFrame.Visible = True
    LeaveBlankCellsBlank.Enabled = False
    CmdOkay.Enabled = True
    OptFrame.Visible = True
    HowToDisplayEmptyCellsFrame.Visible = True
    
    If InStr(Lp_GP_String_3, "L") > 0 Then
        UseColors.Enabled = True
        UseColors.Value = True
        Black.Enabled = True
        DisplayHeadersInBold.Enabled = True
        LeaveBlankCellsBlank.Visible = False
        If FillBlankCellsWithEMDash = True Then
        FillBlankCellsWithNA.Value = False
        End If
    End If
    
    If InStr(Lp_GP_String_3, "R") > 0 Then
        FillBlankCellsWithNA.Value = True
        LeaveBlankCellsBlank.Enabled = True
        LeaveBlankCellsBlank.Visible = True
        LeaveBlankCellsBlank.Value = True
    End If
End Sub

Private Sub RowOnlyTable_Click() 'radio button
    ListGroupSettingsFrame.Visible = True
    OutputListOptionsFrame.Visible = True
    LeaveBlankCellsBlank.Enabled = False
    LeaveBlankCellsBlank.Visible = True
    OptFrame.Visible = True
    HowToDisplayEmptyCellsFrame.Visible = True
    
     If InStr(Lp_GP_String_3, "L") > 0 Then
        UseColors.Enabled = True
        UseColors.Value = True
        Black.Enabled = True
        DisplayHeadersInBold.Enabled = True
        LeaveBlankCellsBlank.Visible = False
        If FillBlankCellsWithEMDash = True Then
        FillBlankCellsWithNA.Value = False
        End If
    End If
        
    If InStr(Lp_GP_String_3, "R") > 0 Then
        FillBlankCellsWithNA.Value = True
        LeaveBlankCellsBlank.Enabled = True
        LeaveBlankCellsBlank.Value = True
    End If
End Sub

Private Sub NoRowOrColumn_Click() 'Radio Button
     If InStr(Lp_GP_String_3, "L") > 0 Then
        LeaveBlankCellsBlank.Visible = False
        If FillBlankCellsWithEMDash = True Then
            FillBlankCellsWithNA.Value = False
        End If
    End If
    
    If InStr(Lp_GP_String_3, "R") > 0 Then
ListGroupSettingsFrame.Visible = False
OutputListOptionsFrame.Visible = False
        UseColors.Value = False
        LeaveBlankCellsBlank.Enabled = True
        LeaveBlankCellsBlank.Visible = True
        LeaveBlankCellsBlank.Value = True
    End If
End Sub
'*************** End Table Type Radio Buttons *****************

Private Sub UseColors_Click() 'radio button
    If InStr(Lp_GP_String_3, "R") > 0 Then 'rotate the table
        If UseColors = True Then
            DisplayHeadersInBold.Value = False
            DisplayHeadersInBold.Enabled = True
            Red.Enabled = True
            Orange.Enabled = True
            Blue.Enabled = True
            Violet.Enabled = True
            Green.Enabled = True
            Black.Enabled = True
        Else  'UseColors = False
            DisplayHeadersInBold.Enabled = False
            DisplayHeadersInBold.Value = False
            Black.Enabled = True
            Red.Enabled = True
            Orange.Enabled = True
            Blue.Enabled = True
            Violet.Enabled = True
            Green.Enabled = True
        End If
    Else  ' request is for list
        If UseColors = True Then
            DisplayHeadersInBold.Value = True
            Black.Value = False
            Red.Value = True
            Orange.Value = False
            Blue.Value = False
            Violet.Value = False
            Green.Value = False
            DisplayHeadersInBold.Enabled = True
            Black.Enabled = True
            Red.Enabled = True
            Orange.Enabled = True
            Blue.Enabled = True
            Violet.Enabled = True
            Green.Enabled = True
        Else  'UseColors = False
            DisplayHeadersInBold.Enabled = False
            DisplayHeadersInBold.Value = False
            Black.Value = False
            Red.Value = True
            Orange.Value = False
            Blue.Value = False
            Violet.Value = False
            Green.Value = False
            Black.Enabled = False
            Red.Enabled = False
            Orange.Enabled = False
            Blue.Enabled = False
            Violet.Enabled = False
            Green.Enabled = False
        End If
    End If
End Sub
Private Sub Black_Click()
    UseColors.Value = True
End Sub

Private Sub Red_Click()
    UseColors.Value = True
End Sub

Private Sub Green_Click()
    UseColors.Value = True
End Sub

Private Sub Violet_Click()
    UseColors.Value = True
End Sub

Private Sub Orange_Click()
    UseColors.Value = True
End Sub

Private Sub Blue_Click()
    UseColors.Value = True
End Sub

Sub UserForm_Initialize()
    Lp_Table_Tools_Menu_Form.Hide
    CmdOkay.Enabled = False
    
    If InStr(Lp_GP_String_3, "R") > 0 Then 'rotation wanted
        FormTitle.Caption = "Table Rotation"
        FormTitle.ControlTipText = "Table Rotation"
        ListGroupSettingsFrame.Caption = "Row/Column Header Colors"
        ListGroupSettingsFrame.ControlTipText = "Row/Column Header Colors Group Frame"
        DisplayHeadersInBold.ControlTipText = "Display colors in Bold Radio Button"
        DisplayHeadersInBold = False
        KeepListGroupsOnSamePage.Value = False
        KeepListGroupsOnSamePage.Enabled = False
        EmptyParaAfterLists.Value = False
        EmptyParaAfterLists.Enabled = False

        FillBlankCellsWithNA.Value = False
        UseColors = False
        Red.Value = True
        Black.Enabled = True
        Red.Enabled = True
        Orange.Enabled = True
        Blue.Enabled = True
        Violet.Enabled = True
        Green.Enabled = True

    Else  'list wanted

        Red.Value = True
        EmptyParaAfterLists.Value = True
        FillBlankCellsWithNA.Value = True
        LeaveBlankCellsBlank.Enabled = False
        LeaveBlankCellsBlank.Visible = False
        LeaveBlankCellsBlank.Value = False
        FormTitle.Caption = "Convert Table to List"
        FormTitle.ControlTipText = "Convert Table to List"
        DisplayHeadersInBold = False
    End If
    
    ' 7/24/2026 - removed "MS_Set_Word_Config_For_Large_Print": opening the document already
    '             configures Word for large print, and re-running it here reset the user's
    '             Styles-pane options (show filter / sort order) every time this form opened.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    ' 9/3/2026 - removed a block that walked the tables to store the selected table's index
    ' in Lp_GP_Counter_1. Its own "Dim Lp_GP_Counter_1 As Long" made a LOCAL of that name,
    ' so the Public it meant to set was never touched and the whole loop was thrown away.
    ' Nothing needs the index now in any case: both conversions work on Selection.Tables(1)
    ' in the transcriber's own document, so there is no second document to find it again in.

End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Lp_Table_Tools_Menu_Form
    Unload Me
    End
End Sub

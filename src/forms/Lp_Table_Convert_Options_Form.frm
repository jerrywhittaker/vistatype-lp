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

'***************** List Table ******************************

    If InStr(Lp_GP_String_3, "L") > 0 Then
        Application.ScreenUpdating = False
        If InStr(Lp_GP_String_3, "X") > 0 Then  'table with row and column headers
            Application.Run MacroName:="Lp_Table_Is_R1C1_Empty" ' also stores documen name into Lp_GP_String_2
            Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
            Application.Run MacroName:="Lp_Table_Cleanup_For_Roation_And_List"
            Application.Run MacroName:="Lp_Table_Fill_Empty_Cells"
            Application.Run MacroName:="Lp_Table_Row_Column_Header_Setup"
            Application.Run MacroName:="Lp_Table_Convert_RC_Table_To_List"

        ElseIf InStr(Lp_GP_String_3, "W") > 0 Then  'table with no row or column headers
            Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
            Application.Run MacroName:="Lp_Table_Fill_Empty_Cells"
            Application.Run MacroName:="Lp_Table_Convert_NoRC_Table_To_List"
            
        ElseIf InStr(Lp_GP_String_3, "Z") > 0 Then  'table with row only headers
            Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
            Application.Run MacroName:="Lp_Table_Transpose_Table" 'table is now a type "Y"
            Lp_GP_String_3 = Replace(Lp_GP_String_3, "Z ", "Y ") 'change the control string
            GoTo TypeYTable
            
        ElseIf InStr(Lp_GP_String_3, "Y") > 0 Then  'table with column header only
            Application.Run MacroName:="Lp_Copy_To_Temp_Doc" ' type Z table has already been transposed
TypeYTable:
            Application.Run MacroName:="Lp_Table_Fill_Empty_Cells"
            Application.Run MacroName:="Lp_Table_Cleanup_For_Roation_And_List"
            Application.Run MacroName:="Lp_Table_Transpose_Table" 'makes the column only table a row-only table
            Application.Run MacroName:="Lp_Table_Row_Column_Header_Setup"
            Application.Run MacroName:="Lp_Table_Convert_R_Only_Table_To_List"
        End If
        
        Application.ScreenUpdating = True
        Application.ScreenRefresh
        DoEvents    ' Let the UI catch up
        MsgBox "Press Ctrl+Z Two (2) times to restore original table.", , "VistaType (178)"
        Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
        Unload Lp_Table_Tools_Menu_Form
        Unload Me
        End
    End If
    
'************ Rotate (transpose) Table ****************
    If InStr(Lp_GP_String_3, "R") > 0 Then
        Application.ScreenUpdating = False
        Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
        TempFileName = ActiveDocument.fullName ' the name of the temporary File
        Application.Run MacroName:="Lp_Table_Cleanup_For_Roation_And_List"
        Application.Run MacroName:="Lp_Table_Transpose_Table"
        Application.Run MacroName:="Lp_Table_Fill_Empty_Cells"
        If Not InStr(Lp_GP_String_3, "W") = 0 Then
            Application.Run MacroName:="Lp_Table_Row_Column_Header_Setup"
        End If

       '**********  Start put a par above table '*********
        Dim tbl As Table
        Dim tblRange As Range
    
        'place transcriber note at top in temp file
        Selection.HomeKey Unit:=wdStory 'top of temp doc - move to the single para mark at top
        Application.Run MacroName:="Lp_Table_Insert_Transcriber_Note"
        Selection.WholeStory
        Selection.Copy
        DoEvents

        'open original doc and delete the table
        Documents(Lp_GP_String_2).Activate
        DoEvents '6/22/2026 need some time here
        DoEvents
        DoEvents
        DoEvents
        ActiveDocument.Tables(Lp_GP_Counter_1).Select
        ActiveDocument.Tables(Lp_GP_Counter_1).Delete

        ' --- PART 1: The Import & UI Kick ---
        Dim srcDoc As Document
        Dim tgtDoc As Document
        Dim targetRange As Range
        Dim originalTemp As Document
        
        ' Capture the ORIGINAL temp doc before anything overwrites srcDoc
        Set originalTemp = Documents(TempFileName)
        
        ' srcDoc is the one you're currently importing from
        Set srcDoc = Documents(TempFileName)
        Set tgtDoc = ActiveDocument
        
        ' Make sure the target doc owns the selection BEFORE using Selection
        tgtDoc.Activate
        
        ' Define where the content should land
        Set targetRange = Selection.Range
        targetRange.Collapse Direction:=wdCollapseStart
        
        ' Move EVERYTHING (Text, Tables, Styles) from temp doc
        targetRange.FormattedText = srcDoc.Content.FormattedText
        
        ' --- CLOSE BOTH TEMP DOCS ---
        Application.DisplayAlerts = wdAlertsNone
        
        ' Close the one you just imported from
        If Not srcDoc Is Nothing Then
            srcDoc.Close SaveChanges:=wdDoNotSaveChanges
        End If
        
        ' Close the ORIGINAL temp doc (the one that was being left open)
        If Not originalTemp Is Nothing Then
            originalTemp.Close SaveChanges:=wdDoNotSaveChanges
        End If
        
        Application.DisplayAlerts = wdAlertsAll

        ' --- THE VISUAL SHOCK: Applying your Style ---
        ' Applying a style forces Word's layout engine to "Draw" the table NOW
        On Error Resume Next
        With targetRange.Tables(1)
            '.Style = "yellow on white paper table"
            .AllowAutoFit = True
            .AutoFitBehavior (wdAutoFitWindow)
        End With
        On Error GoTo 0

        ' --- THE REFRESH SEQUENCE ---
        ' Force ScreenUpdating back to True
        Do While Application.ScreenUpdating = False
            Application.ScreenUpdating = True
        Loop

        ' Snap the "Camera" to the new table so it's visible behind the MsgBox
        targetRange.Tables(1).Select
        ActiveWindow.ScrollIntoView Selection.Range
        ActiveDocument.Repaginate
        
        ' The "Jolt" - Minimize/Maximize forces Windows OS to repaint the pixels
        ActiveWindow.WindowState = wdWindowStateMinimize
        DoEvents
        ActiveWindow.WindowState = wdWindowStateMaximize

        ' Final hard refresh
        Application.ScreenRefresh
        DoEvents

        ' --- NOTIFICATIONS ---
        ' vbSystemModal forces the box to the front and gives Word a moment to breathe
        MsgBox "Press Ctrl+Z (undo) four (4) times to restore original table.", vbSystemModal, "VistaType LP (200)"
        
        If InStr(Lp_GP_String_3, "X") > 0 Then
            MsgBox "Check the header in column 1 row 1. The title should describe the data in the column.", vbSystemModal, "VistaType LP (202)"
        End If

        ' --- CLEANUP ---
        Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
        
        ' Unload forms cleanly
        On Error Resume Next
        Unload Lp_Table_Tools_Menu_Form
        Unload Me
    End If
End Sub



'********* Begin Table Type Image Buttons - these select the table type radio buttons ********

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
    
    '*** store the index of the selected table ***
    Dim tbl As Table
    Dim Lp_GP_Counter_1 As Long
    Dim i As Long

    If Selection.Information(wdWithInTable) Then
        For i = 1 To ActiveDocument.Tables.count
            Set tbl = ActiveDocument.Tables(i)
            If Selection.Range.start >= tbl.Range.start And Selection.Range.End <= tbl.Range.End Then
                Lp_GP_Counter_1 = i
                Exit For
            End If
        Next i
    End If

End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Lp_Table_Tools_Menu_Form
    Unload Me
    End
End Sub

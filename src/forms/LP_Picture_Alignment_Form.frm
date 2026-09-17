VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Picture_Alignment_Form 
   Caption         =   "Align Pictures (339)"
   ClientHeight    =   5172
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7170
   OleObjectBlob   =   "LP_Picture_Alignment_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LP_Picture_Alignment_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Lp_Picture_Alignment_Form
'
' Version: 1.3  Date: 9/17/2026 - with anything selected, 'All pictures in the document' is now
'                                 grayed out as well as unselected; and a WHOLE TABLE selected on
'                                 its own also selects 'Include pictures in tables' and grays out
'                                 'Exclude pictures in tables', which would otherwise leave the
'                                 macro nothing to do (Jerry)
' Version: 1.2  Date: 9/17/2026 - the scope option now follows what is selected: with a range,
'                                 a picture or a table selected the dialog opens on 'Pictures in
'                                 a selected text range, selected picture, or selected table'
'                                 rather than on the whole document (Jerry)
' Version: 1.1  Date: 7/26/2026 - fixed a slip that called Sh_Create_Temp_Bookmark where the move-and-delete was meant, so the user was never put back and a stale bookmark was left behind; now returns the user to where the cursor was when Okay was clicked
' Version 1.0  Date: 4/11/2025
'

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOkay_Click()
    Sh_Save_User_Position
    Application.ScreenUpdating = False
    Unload Me
      
    If AllPicturesInDocButton Then
        If CenterButton Then
            For Each oILShp In ActiveDocument.InlineShapes
                oILShp.Select
                If ExcludeTablesButton And Selection.Information(wdWithInTable) Then
                    GoTo skip1
                End If
                Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
skip1:
            Next
        Else 'LeftAllineButton
            For Each oILShp In ActiveDocument.InlineShapes
            oILShp.Select
            If ExcludeTablesButton And Selection.Information(wdWithInTable) Then
                GoTo skip2
            End If
            Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
skip2:
            Next
        End If
    End If

    If SelectedItemOrRangeButton Then

        If Selection.Type = wdSelectionNormal Or Selection.Information(wdWithInTable) Or Selection.InlineShapes.count = 1 Then
        Else
            MsgBox "Nothing is selected or the selected picture is NOT inline", , "VistaType LP (148)"
            ' PUT THE SCREEN BACK before stopping. Screen updating went off at the top of this
            ' handler and End does not restore it, so the user pressed OK on the message and
            ' Word sat there frozen. Jerry, 8/26/2026.
            '
            ' End is KEPT here, unlike the two in Dx_UnEmbed_Ref_Pg_No. This is a form's button
            ' handler: Exit Sub would return to the form, which would still be on screen and
            ' waiting - a different behavior, not just a tidier one. Worth doing if that is what
            ' is wanted, but it is a decision about the dialog rather than a repair.
            Application.ScreenUpdating = True
            End
        End If
        
        '****************************
        ' selected Picture in a table
        '****************************
        ' Check if the selection is within a table
        If Selection.Information(wdWithInTable) Then
            Dim tblf As Table
            Set tblf = Selection.Tables(1) ' Get the active table
            ' check if  full table selection by comparing ranges bypass if true - only single pic in table selected
            If Selection.Range.start = tblf.Range.start And Selection.Range.End = tblf.Range.End Then
                GoTo Skip3
            End If
            ' Check if the selection is an InlineShape (e.g., picture)
            If Selection.InlineShapes.count > 0 Then
                'On Error Resume Next
                Dim pic1 As inlineShape
                Set pic1 = Selection.InlineShapes(1)
                ' Align the picture within the table cell
                If LeftAlignButton Then
                    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                Else 'Center Button
                    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End If
            Else
                MsgBox "The selection in the table is not a picture.", , "VistaType LP (149)"
            End If
            GoTo eom
        End If 'SelectedItemOrRangeButton
        
Skip3:
        '****************************
        ' selected Picture
        '****************************
        ' Check if the selection contains exactly one inline shape (e.g., a picture)
        If Selection.InlineShapes.count = 1 Then
            Dim shp1 As inlineShape
            Set shp1 = Selection.InlineShapes(1)
    
            ' Center the picture by aligning its paragraph
            If LeftAlignButton Then
                shp1.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            Else 'Center Button
                shp1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            GoTo eom
            End If 'Selection.InlineShapes.Count = 1
        End If
       
        '****************************
        ' selected table
        '****************************
        Dim tbl As Table
        Dim cell As cell
        Dim shp As inlineShape
        
        ' Check if the entire table is selected
        If Selection.Information(wdWithInTable) Then
            Set tbl = Selection.Tables(1) ' Get the selected table
            
            ' Verify if the entire table is selected
            If Selection.Range.start = tbl.Range.start And Selection.Range.End = tbl.Range.End Then
                ' Loop through each cell in the table
                For Each cell In tbl.Range.Cells
                    ' Loop through inline shapes (pictures) in the cell
                    For Each shp In cell.Range.InlineShapes
                        shp.Select
                        'Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        If LeftAlignButton Then
                            Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                        Else 'Center Button
                            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        End If
                    Next shp
                Next cell
            End If
            GoTo eom:
        End If
        
        '*******************************************
        ' Selected Text range with or without tables
        '*******************************************
        Dim selRange As Range
        Dim inlineShape As inlineShape
        
        ' Validate the selection
        If Selection.Type <> wdSelectionNormal Then
            MsgBox "Please select text containing pictures to align.", vbExclamation, "VistaType LP (297)"
            Exit Sub
        End If
        
        ' Set the range to the selection
        Set selRange = Selection.Range
        
        ' Loop through all inline shapes in the document
        For Each inlineShape In ActiveDocument.InlineShapes
            With inlineShape
                ' Check if the inline shape is within the selected range
                If .Range.InRange(selRange) Then
                    ' Determine inclusion/exclusion of pictures in tables
                    If IncludeTablesButton Or (ExcludeTablesButton And .Range.Tables.count = 0) Then
                        ' Apply alignment based on button settings
                        If LeftAlignButton Then
                            .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                        ElseIf CenterButton Then
                            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        End If
                    End If
                End If
            End With
        Next inlineShape
        
    End If 'SelectedItemOrRangeButton
    
eom: 'End of Macro

    ' This used to call Sh_Create_Temp_Bookmark, which was a slip - it dropped a FRESH mark
    ' here instead of returning to the one made on entry, so the user was never put back and
    ' a stale TempPlaceholder was left in the document for the next macro to trip over.
    Application.ScreenUpdating = True
    Sh_Return_User_To_Start_Position
    On Error GoTo 0

End Sub

Private Sub UserForm_Initialize()

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

    ' No position is saved here on purpose. CmdOkay_Click saves it when work actually
    ' starts, so cancelling the dialog leaves nothing behind.

    ' THE DIALOG FOLLOWS WHAT IS SELECTED (Jerry, 9/17/2026). Somebody who has selected a
    ' picture, a table or a run of text has already said what they mean, and the dialog
    ' opening on 'All pictures in the document' invited one click that realigned the whole
    ' book. With nothing selected the selected-range option cannot work anyway - CmdOkay_Click
    ' stops on dialog 148 - so the whole-document option is the only sensible default there.
    '
    ' Selection.Start <> Selection.End covers selected text, a selected table and a selected
    ' picture, since a picture selected by clicking it is a one-character range. The type test
    ' beside it is belt and braces for a shape selected some other way. It is NOT the
    ' wdWithInTable test CmdOkay_Click uses: a cursor merely sitting in a table selects
    ' nothing, and would have flipped the default with nothing to work on.
    '
    ' SET THE VALUE BEFORE DISABLING THE OTHER BUTTON, both times below. Disabling an option
    ' button that is still the selected one leaves it grayed AND chosen, which reads as the
    ' dialog having made an impossible choice.
    Dim selTbl As Table
    Dim wholeTableOnly As Boolean

    If Selection.Start <> Selection.End Or Selection.Type = wdSelectionInlineShape Then
        SelectedItemOrRangeButton.Value = True

        ' Nothing is selected by accident, so offering to realign the whole book from here is
        ' offering the one thing that would undo the transcriber's own choice of scope.
        AllPicturesInDocButton.Enabled = False

        ' A WHOLE TABLE AND NOTHING ELSE. The same test CmdOkay_Click uses to recognize one:
        ' the selection's range is exactly the table's range, so a table that is merely part
        ' of a longer selection does not match. Guarded, because Tables(1) raises when the
        ' selection is not in a table at all.
        On Error Resume Next
        If Selection.Information(wdWithInTable) Then
            Set selTbl = Selection.Tables(1)
            If Not selTbl Is Nothing Then
                wholeTableOnly = (Selection.Range.Start = selTbl.Range.Start _
                              And Selection.Range.End = selTbl.Range.End)
            End If
        End If
        Err.Clear
        On Error GoTo 0

        ' Every picture in reach is inside that table, so excluding pictures in tables would
        ' leave the macro nothing at all to do (Jerry, 9/17/2026).
        If wholeTableOnly Then
            IncludeTablesButton.Value = True
            ExcludeTablesButton.Enabled = False
        End If
    Else
        AllPicturesInDocButton.Value = True
    End If

End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub



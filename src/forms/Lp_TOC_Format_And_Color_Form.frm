VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_TOC_Format_And_Color_Form 
   Caption         =   "Format TOC - Add/Remove Color Bars"
   ClientHeight    =   2628
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4095
   OleObjectBlob   =   "Lp_TOC_Format_And_Color_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_TOC_Format_And_Color_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lp_TOC_Format_And_Color_Form
'
' Version: 1.1  Date: 9/2/2026 - dialog 207 no longer promises a number of Ctrl+Z presses.
'                               Lp_TOC_CleanAndFormat_TOC stopped round-tripping through a
'                               scratch document, so the two presses that bought are gone; and
'                               it CANNOT have a custom undo record instead - one crashed Word
'                               outright. See the note in that macro.
' Version: 1.0  Date: 8/25/2025
'
Private Sub CancelButton_Click()
    Unload Lp_TOC_Format_And_Color_Form
    End
End Sub

Private Sub OkayButton_Click()
    Lp_TOC_Format_And_Color_Form.Hide
    
    If FormatTheTOCButton Then
        Application.Run MacroName:="Lp_TOC_CleanAndFormat_TOC"
        MsgBox "Press Ctrl+Z, more than once, to return to the original TOC.", , "VistaType LP (207)"
        Exit Sub
        
    ElseIf AddColorBarsButton Then
        Lp_TOC_Color_Bars_Form.Show
        Selection.Collapse Direction:=wdCollapseStart
        DoEvents
        MsgBox "Color bars added.", , "VistaType LP (208)"
        Exit Sub

    ElseIf RemoveColorBarsButton Then
        Dim para As Paragraph
        Dim st As Style
        Dim stylesToMatch As New Collection
        Dim i As Long
        Dim rngPara As Range, rr As Range
        Dim changed As Boolean
        Dim pf As ParagraphFormat
        Dim ts As tabStop
        Dim pos As Single, align As WdTabAlignment
    
        ' Collect only character styles beginning with "Words"
        For Each st In ActiveDocument.Styles
            If st.Type = wdStyleTypeCharacter Then
                If UCase$(st.NameLocal) Like "WORDS*" Then stylesToMatch.Add st
            End If
        Next st
        If stylesToMatch.count = 0 Then Exit Sub
    
        For Each para In Selection.Paragraphs
            changed = False
    
            ' Paragraph-scoped search range (extend by 1 to catch end-run)
            Set rngPara = para.Range.Duplicate
            If rngPara.End < ActiveDocument.Content.End Then rngPara.End = rngPara.End + 1
    
            ' Strip only the Words* character style runs
            For i = 1 To stylesToMatch.count
                Set rr = rngPara.Duplicate
                With rr.Find
                    .ClearFormatting
                    .Text = ""
                    .Format = True
                    .Style = stylesToMatch(i)
                    .Forward = True
                    .Wrap = wdFindStop
                    Do While .Execute
                        ' Replace char style with Default Paragraph Font (keeps direct formatting)
                        rr.Style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
                        changed = True
                        rr.Collapse wdCollapseEnd
                    Loop
                End With
            Next i
    
            ' Only paragraphs that actually contained Words* runs get re-tagged to TOC 1
            If changed Then
                Set pf = para.Range.ParagraphFormat.Duplicate  ' preserve spacing/indents etc.
                para.Style = ActiveDocument.Styles("TOC 1")
                para.Range.ParagraphFormat = pf                ' restore paragraph-level formatting
            End If
        Next para
        
        For Each para In Selection.Paragraphs
            If UCase$(para.Style.NameLocal) Like "TOC*" Then
                If para.TabStops.count > 0 Then
                    ' Store the first tab stop’s position and alignment
                    pos = para.TabStops(1).Position
                    align = para.TabStops(1).Alignment
    
                    ' Clear all direct tab stops
                    para.TabStops.ClearAll
    
                    ' Add a new tab stop at the same position/alignment with a dot leader
                    para.TabStops.Add Position:=pos, Alignment:=align, Leader:=wdTabLeaderDots
                End If
            End If
        Next para
    End If
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    ActiveWindow.View.Type = wdNormalView
    ActiveWindow.View.Type = wdPrintView
    Selection.Collapse Direction:=wdCollapseStart
    DoEvents
    MsgBox "Color bars removed.", , "VistaType LP (209)"
    
    Exit Sub
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
    End
End Sub

Private Sub UserForm_Initialize()

    Lp_GP_String_1 = ActiveDocument.Name
    
    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DN_Text_File_Para_Fix_Warning 
   Caption         =   "Caution"
   ClientHeight    =   6288
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   OleObjectBlob   =   "DN_Text_File_Para_Fix_Warning.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DN_Text_File_Para_Fix_Warning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
    Unload Me
    End
End Sub

Private Sub ContinueButton_Click()
'
' Author: Jerry Whittaker -  jerry@vistatypelp.org
'
' Date: 2/3/2018
' Version: 1.3
'
' Change Log:
'        V1.2 to 1.3
'           -fixed end of macro message
'
    DN_Text_File_Para_Fix_Warning.Hide
    
    ' convert the whole document into normal style
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("Normal")
    Selection.HomeKey Unit:=wdStory
    
    ' replace two ajcent para marks with a marker code
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "$#*@"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' replace all para marks with a space
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' replace each marker code with a para mark
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$#*@"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
    ' replace multiple consecutive para marks with just one para mark
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
      
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
    ' 9/1/2026 - removed: MsgBox "End of Macro" (Jerry). It told the transcriber nothing, and it
    ' carried no title, so Word captioned it "Microsoft Word" - it did not even look like part of
    ' VistaType LP. The twin of the one taken off Lp_Replace_Section_Break_With_Page_Break the
    ' same day; both have the shape of a line left in from debugging. No dialog number to retire.
    '
    ' The two lines above it were NOT touched, and both are worth a look one day: the clipboard
    ' is emptied on purpose, and ActiveDocument.UndoClear throws away the WHOLE undo history
    ' rather than this macro's part of it.
    Unload Me
    
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub


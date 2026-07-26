VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Section_Brk_Caution 
   Caption         =   "Caution"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   OleObjectBlob   =   "Lp_Section_Brk_Caution.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Section_Brk_Caution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Version: 1.5  Date: 7/26/2026 - returns the user to where the cursor was when Okay was clicked
' Version: 1.4  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version 1.3  Date: 1/6/2019
' Version 1.2  Date: 11/15/2018
'
Private Sub CmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub CmdOkay_Click()

    Dim Limited_Selection As Boolean
    
    Application.ScreenUpdating = False ' Turn screen updating off
    
    Sh_Save_User_Position

    If Selection.Type <> wdSelectionNormal Then   'text is NOT selected"
        Limited_Selection = False
    Else
        Limited_Selection = True
        Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
    End If
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = "^m"
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
    
    If Limited_Selection = True Then
        Selection.Delete Unit:=wdCharacter, count:=1
        Application.Run MacroName:="Lp_Copy_From_Temp_Doc"
        Selection.TypeBackspace
    End If
    
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    Application.ScreenUpdating = True ' Turn screen updating on
    Sh_Return_User_To_Start_Position
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
    
    Unload Me
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
    Application.ScreenUpdating = True ' Turn screen updating on
    End
End Sub

Private Sub UserForm_Initialize()
    ' 7/24/2026 - removed "MS_Set_Word_Config_For_Large_Print": opening the document already
    '             configures Word for large print, and re-running it here reset the user's
    '             Styles-pane options (show filter / sort order) every time this form opened.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub


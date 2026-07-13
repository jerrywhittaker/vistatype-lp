VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Validation_Choices_Form 
   Caption         =   "Validate Reference Page Tags"
   ClientHeight    =   4968
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   OleObjectBlob   =   "Sh_Validation_Choices_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Validation_Choices_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version: 1.2  Date: 2/12/2024 - added code to place reference page tags in 3 column temp doc
' Version: 1.1  Date: 2/14/2021 - fixed problem with info display page
' Version: 1.0  Date: 9/26/2018
' Author: Jerry Whittaker  jerry@thewhittakers.org

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub NavigationButton_Click()
    ActiveWindow.DocumentMap = True
    Selection.Find.ClearFormatting
    Selection.HomeKey Unit:=wdStory
    
    With Selection.Find
            .Text = "$pg"
    End With

    Sh_SetBarVisible "Navigation", True
    SendKeys "^" + "f" + "{right}{Enter}{right}{right}{right}"  'Turn on navigation pane and set up search
    Unload Me
End Sub

Private Sub TempFileButton_Click()

    Sh_Validation_Choices_Form.Hide
    Sh_SetBarVisible "Navigation", False
    Application.ScreenUpdating = False    ' Turn screen updating off
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

    With Selection.Find
    .Text = "$pg*^013"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .MatchWildcards = True
    End With

    SendKeys "^h"      '  Ctrl+h - open find and replace dialog
    SendKeys "%d"      ' Alt+d - move to find tab
    SendKeys "%i"       ' Alt+i - Find In
    SendKeys "m"         ' Main Document
    SendKeys "{Esc}"    ' close find and replace dialog box

    Application.ScreenUpdating = True    ' Turn screen updating on
    
    ' pause for 1 second then run the "Sh_Copy_Ref_Pg_Tags_To_Temp_File" macro
    Application.OnTime When:=Now + TimeValue("00:00:01"), Name:="Sh_Copy_Ref_Pg_Tags_To_Temp_File"
    
    Unload Sh_Validation_Choices_Form

End Sub
Private Sub ShowMoreButton_Click()
    Sh_Pg_Validation_Overview_Form.Show
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



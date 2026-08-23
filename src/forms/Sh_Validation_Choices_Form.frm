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
' Version: 1.3  Date: 8/23/2026 - the temp-file route no longer drives Find and Replace with
'                                 SendKeys, so the dialog cannot flash on screen (Jerry)
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

' NO SendKeys AND NO PAUSE ANY MORE. Jerry, 8/23/2026: the Find and Replace dialog flashed on
' screen. It flashed because it was really being opened - this used to type Ctrl+H, Alt+D, Alt+I,
' M, Esc into it to reach "Find In > Main Document", which selects every $pg paragraph at once so
' that one Copy takes the lot. ScreenUpdating cannot hide a dialog; it stops the DOCUMENT
' repainting and nothing else. So the dialog had to go, not be hidden.
'
' The one-second Application.OnTime went with it. That was not a courtesy pause - it was the only
' way to wait for keystrokes SendKeys hands to Windows and cannot follow. The list is now built on
' a range inside Sh_Copy_Ref_Pg_Tags_To_Temp_File, so it can simply be called.
'
' Version: 1.0  Date: 8/23/2026
Private Sub TempFileButton_Click()

    Sh_Validation_Choices_Form.Hide
    Sh_SetBarVisible "Navigation", False

    Application.Run MacroName:="Sh_Copy_Ref_Pg_Tags_To_Temp_File"

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



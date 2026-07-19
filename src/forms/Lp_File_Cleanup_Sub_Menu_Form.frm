VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_File_Cleanup_Sub_Menu_Form 
   Caption         =   "Large Print File Cleanup"
   ClientHeight    =   2388
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "Lp_File_Cleanup_Sub_Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_File_Cleanup_Sub_Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lp_File_Cleanup_Sub_Menu_Form
'
' Version: 1.1  Date: 1/22/2026 - refresh screen before messages
' Version: 1.0  Date: 5/7/2025 - full rewright
'

Private Sub CancelButton_Click()
    Selection.Collapse 'clear selection
    Unload Me
End Sub

Private Sub OkayButton_Click()

    Lp_File_Cleanup_Sub_Menu_Form.Hide
    
    ' Hold screen updating off across the selected cleanup(s); restored to True below.
    Application.ScreenUpdating = False
    
    If FixCommonErrors Then
        Application.Run MacroName:="Lp_Fix_Common_File_Errors"
    End If
    
    If RemoveParaMarks Then
        Application.Run MacroName:="Lp_Replace_Multiple_Para_Marks_With_Warning"
    End If
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True

    ' Force Word to repaint the screen NOW
    DoEvents

    If FixCommonErrors Or RemoveParaMarks Then
        MsgBox "Selected cleanup(s) complete", , "VistaType LP (187)"
    Else
        If Not (FixCommonErrors Or RemoveParaMarks) Then
            MsgBox "Canceled... No cleanup selected", , "VistaType LP (188)"
        End If
    End If
    
    Selection.Collapse 'clear selection
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub



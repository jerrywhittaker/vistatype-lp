VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Space_Or_Para_Form 
   Caption         =   "Choose Replacement Character (362)"
   ClientHeight    =   2064
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   OleObjectBlob   =   "Sh_Space_Or_Para_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Space_Or_Para_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sh_Space_Or_Para_Form
'
' Version: 1.1  Date: 10/11/2026 - the red X acts like Cancel (Jerry, 10/11/2026). It only unloaded
'                                the box and left Sh_GP_String_1 as it was, so a "Para" or "Space"
'                                from an earlier answer could carry the macro on.
'
Private Sub Cmd_Cancel_Click()
    Sh_GP_String_1 = ""
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Handles the red X - the same as Cancel
    If CloseMode = vbFormControlMenu Then
        Cancel = True                 'prevents default close behavior
        Cmd_Cancel_Click
    End If
End Sub

Private Sub Cmd_Okay_Click()
    If Para_Mark Then
        Sh_GP_String_1 = "Para"
    End If
    If Space Then
        Sh_GP_String_1 = "Space"
    End If
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub



VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_UEB_EBAE_Fill_In_YN_Form 
   Caption         =   "Two Level Exercise Fill-In Indicators (336)"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7785
   OleObjectBlob   =   "Dx_UEB_EBAE_Fill_In_YN_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_UEB_EBAE_Fill_In_YN_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Version: 1.6  Date: 10/11/2026 - the red X acts like Cancel (Jerry, 10/11/2026). It only unloaded
'                                the box, so Dx_Format_Exercise_Lv_1_and_Lv_2 carried on and
'                                formatted the exercise with the last Yes or No given.
' Date: 6/23/2018
' Version: 1.5

Private Sub CmdCancel_Click()
    Unload Me 'close this form
    End
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Handles the red X - the same as Cancel
    If CloseMode = vbFormControlMenu Then
        Cancel = True                 'prevents default close behavior
        CmdCancel_Click
    End If
End Sub
Private Sub CmdNO_Click()
    Dx_UEB_EBAE_Boolean = False
    Unload Me 'close this form
End Sub
Private Sub CmdYES_Click()
    Dx_UEB_EBAE_Boolean = True
    Unload Me 'close this form
End Sub
Private Sub UserForm_Initialize()
    If Dx_UEB_EBAE_Boolean = True Then
         Dx_UEB_EBAE_Fill_In_YN_Form.CmdYES.SetFocus
    Else
        Dx_UEB_EBAE_Fill_In_YN_Form.CmdNO.SetFocus
    End If
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub



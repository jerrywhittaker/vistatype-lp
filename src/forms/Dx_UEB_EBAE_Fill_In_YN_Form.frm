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

' Version: 1.8  Date: 10/12/2026 - the red X did nothing: QueryClose set Cancel = True,
'                                then unloaded, and Word ignores an Unload there (Jerry, 10/12/2026).
'                                QueryClose is gone: a plain close leaves Dx_Fill_In_Answered
'                                False, which is already Cancel.
' Version: 1.7  Date: 10/11/2026 - Cancel (and so the red X) no longer ends with End (issue #32).
'                                End reset the whole VBA project and closed every box that stays
'                                open (375, the TOC box, the $pg menus). Cancel now only unloads
'                                the box; Yes and No raise Dx_Fill_In_Answered, so
'                                Dx_Format_Exercise_Lv_1_and_Lv_2 can tell an answer from Cancel.
' Version: 1.6  Date: 10/11/2026 - the red X acts like Cancel (Jerry, 10/11/2026). It only unloaded
'                                the box, so Dx_Format_Exercise_Lv_1_and_Lv_2 carried on and
'                                formatted the exercise with the last Yes or No given.
' Date: 6/23/2018
' Version: 1.5

Private Sub CmdCancel_Click()
    Unload Me 'close this form - Dx_Fill_In_Answered stays False, so the exercise is not formatted
End Sub

Private Sub CmdNO_Click()
    Dx_UEB_EBAE_Boolean = False
    Dx_Fill_In_Answered = True
    Unload Me 'close this form
End Sub
Private Sub CmdYES_Click()
    Dx_UEB_EBAE_Boolean = True
    Dx_Fill_In_Answered = True
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



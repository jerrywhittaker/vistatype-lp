VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Resize_Images_Form 
   Caption         =   "Resize Selected Pictures (350)"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5835
   OleObjectBlob   =   "Lp_Resize_Images_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Resize_Images_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOkay_Click()
    ' Version: 2.0  Date: 10/14/2026 - Jerry: a selection is required, so the "All Pictures"
    '                               choice and its two option buttons are gone. A percent out of
    '                               range now stops here: before, the message showed and the
    '                               resize ran anyway with the last percent used.

    ' check the percent value range
    If Val(Percent_Requested) > 575 Or Val(Percent_Requested) < 1 Then
        MsgBox "Percent is out of range", , "VistaType LP (300)"
        Exit Sub
    End If

    ' Val first: Int("50%") stops with Type mismatch, and Val("50%") is 50
    Lp_Pic_Percent = Str(Int(Val(Percent_Requested)))  'save the choice
    Lp_Resize_Images_Form.Hide

    Application.Run MacroName:="Lp_Resize_Images"
    Unload Me 'close this form

End Sub

Private Sub UserForm_Initialize()

    If Val(Lp_Pic_Percent) < 1 Then
        Lp_Pic_Percent = "50"
    End If

    Percent_Requested = Val(Lp_Pic_Percent)

    With Me.Percent_Requested
        .selStart = 0
        .SelLength = 3
    End With

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)


End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

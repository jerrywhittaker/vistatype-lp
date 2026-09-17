VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Bullet_Removal_Form 
   Caption         =   "Choose paragraph style (329)"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4155
   OleObjectBlob   =   "Dx_Bullet_Removal_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Bullet_Removal_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub CmdOkay_Click()
    If TOC1 Then
        Dx_GP_String_2 = "T1"
        Unload Me
    End If
    
    If TOC2 Then
        Dx_GP_String_2 = "T2"
        Unload Me
    End If

    If List1 Then
        Dx_GP_String_2 = "L1"
        Unload Me
    End If
    
    If List2 Then
        Dx_GP_String_2 = "L2"
        Unload Me
    End If
    
    If BodyText Then
        Dx_GP_String_2 = "B"
        Unload Me
    End If
    
    If Unchanged Then
        Dx_GP_String_2 = "U"
        Unload Me
    End If

End Sub

Private Sub More_Info_Click()
    Sh_Hyperlink_Info_Form.Show
End Sub

Private Sub Remove_Hyperlinks_Click()
        Dx_GP_String_1 = "Remove_Hyper"
End Sub

Private Sub UserForm_Initialize()
    Remove_Hyperlinks = False
    Dx_GP_String_1 = ""
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub



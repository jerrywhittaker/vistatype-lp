VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Create_Brl_Bullets_Form 
   Caption         =   "Format /Remove Bullets"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5385
   OleObjectBlob   =   "Dx_Create_Brl_Bullets_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Create_Brl_Bullets_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdCancel_Click()
    Unload Me 'close this form
    End
End Sub

Private Sub CmdOkay_Click()
    Unload Me 'close this form
    
    If Want_Primary Then
        Dx_GP_Counter_1 = 1
    Else ' Want_Secondary
        Dx_GP_Counter_1 = 2
    End If
    
    If Has_Bullets Then
        Dx_GP_Boolean_1 = True
    Else ' has no bullets
        Dx_GP_Boolean_1 = False
    End If
    
    If Remove_Bullets Then
        Dx_GP_String_2 = "Remove_Bullets"
    End If
        
End Sub

Private Sub Has_Bullets_Click()
    Remove_Bullets.Enabled = True
End Sub

Private Sub Has_No_Bullets_Click()
    Remove_Bullets.Enabled = False
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
    Dx_GP_String_2 = "Keep_Bullets"
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub


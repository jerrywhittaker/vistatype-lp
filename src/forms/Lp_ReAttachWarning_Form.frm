VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_ReAttachWarning_Form 
   Caption         =   "Template Re-Attachment Cautions"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "Lp_ReAttachWarning_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_ReAttachWarning_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNoButton_Click()
    Unload Me
End Sub

Private Sub CmdYesButton_Click()
    Lp_GP_String_1 = "Doc_Is_Already_LP"
    Lp_ReAttachWarning_Form.Hide
    LP_Attach_An_Lp_Template_Form.Show
    Unload Me
End Sub

Private Sub userform_terminate() 'red X was clicked
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


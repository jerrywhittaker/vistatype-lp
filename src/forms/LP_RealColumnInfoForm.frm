VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LP_RealColumnInfoForm 
   Caption         =   "What is a real column?"
   ClientHeight    =   6948
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "LP_RealColumnInfoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LP_RealColumnInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
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

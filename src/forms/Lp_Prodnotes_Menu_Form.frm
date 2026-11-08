VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Prodnotes_Menu_Form 
   Caption         =   "Prodnote Review (425)"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   OleObjectBlob   =   "Lp_Prodnotes_Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Prodnotes_Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Lp_Prodnotes__Menu_Form - Prodnotes Review (425), VistaType LP File Cleanup group

Private Sub Run_After_Menu(ByVal macroName As String)
    Dx_Prodnotes_Alt_Text_Menu_Form.Hide
    Sh_BridgeTargetMacro = macroName
    Application.OnTime Now, "Sh_StartSpinnerBridge"
    Unload Me
End Sub

Private Sub ReviewProdnotes_Click()
    Run_After_Menu "Dx_Review_Prodnotes"
End Sub

Private Sub DeleteProdnotesButton_Click()
    Run_After_Menu "Sh_Delete_Prodnote_Paragraphs"
End Sub

Private Sub ProdnoteToTNButton_Click()
    Run_After_Menu "Dx_Change_Prodnotes_To_Transcriber_Notes"
End Sub

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub userform_terminate() 'red X was clicked - nothing runs, the same as Cancel
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Start the form centered inside the Word window (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub


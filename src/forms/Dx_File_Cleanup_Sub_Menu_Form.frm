VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_File_Cleanup_Sub_Menu_Form 
   Caption         =   "Braille File Cleanup (332)"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4905
   OleObjectBlob   =   "Dx_File_Cleanup_Sub_Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_File_Cleanup_Sub_Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Dx_File_Cleanup_Sub_Menu_Form
' Version 1.0
' Date: 5/22/2018


Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Fix_Errors_Button_Click()
    Dx_File_Cleanup_Sub_Menu_Form.Hide
    Application.Run MacroName:="Dx_Fix_Common_File_Errors"
End Sub

Private Sub Del_Para_Marks_Button_Click()
    Dx_File_Cleanup_Sub_Menu_Form.Hide
    Application.Run MacroName:="Dx_Replace_Multiple_Para_Marks_With_Warning"
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub


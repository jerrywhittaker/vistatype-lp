VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Choose_Translation_Form 
   Caption         =   "What braille translation will be used for this document?"
   ClientHeight    =   3144
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5355
   OleObjectBlob   =   "Dx_Choose_Translation_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Choose_Translation_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dx_Choose_Translation_Form
' Version: 1.1 Date: 3/9/2024 - removed "End" from Sub Cmd_Cancel_Click
'
Private Sub Cmd_Cancel_Click()
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_EBAE_Button_Click()
    Dx_UEB_EBAE_String = "EBAE"
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ActiveDocument.Variables("BrailleType").Delete
    ActiveDocument.Variables.Add Name:="BrailleType", Value:="EBAT" 'place EBAE Textbook type into document info
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_EBAE_Nemeth_Button_Click()
    Dx_UEB_EBAE_String = "EBAE Nemeth"
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ActiveDocument.Variables("BrailleType").Delete
    ActiveDocument.Variables.Add Name:="BrailleType", Value:="EBAN" 'place EBAE Nemeth type into document info
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_UEB_Button_Click()
    Dx_UEB_EBAE_String = "UEB"
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ActiveDocument.Variables("BrailleType").Delete
    ActiveDocument.Variables.Add Name:="BrailleType", Value:="UEBT" 'place UEB Textbook type into document info
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_UEB_Nemeth_Button_Click()
    Dx_UEB_EBAE_String = "UEB Nemeth"
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ActiveDocument.Variables("BrailleType").Delete
    ActiveDocument.Variables.Add Name:="BrailleType", Value:="UEBN" 'place UEB Nemeth type into document info
    Unload Me 'close this form
End Sub



Private Sub UserForm_Initialize()

    Application.Run MacroName:="MS_Set_Word_Config_For_Braille"
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DN_Auto_Or_Manual_Form 
   Caption         =   "Choose XML Conversion Type"
   ClientHeight    =   4608
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5790
   OleObjectBlob   =   "DN_Auto_Or_Manual_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DN_Auto_Or_Manual_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'DN_Auto_Or_Manual_Form
'
'Gives user the choice of adding $pg tags a manually converted NIMAS or DAISY .xml file to to converta and tag automatically
'
' Version: 1.1  Date' 7/8/2026


Private Sub Automatic_Button_Click()
    Me.Hide 'hide this user form
    Sh_BridgeTargetMacro = "Sh_Convert_XML_File_To_Word_Document" ' there must be a complete end of this macro in order to display the non-modal message
    Application.OnTime Now, "Sh_StartSpinnerBridge" ' after the non-modal message is displayed the "Sh_StartSpinnerBridge" will initiate the
                                                    ' second portion of the LP attachment "Lp_Attach_The_Template" will begin
    Unload Me 'unload this user form
    'Application.Run MacroName:="Sh_Convert_XML_File_To_Word_Document"
End Sub

Private Sub Manual_Button_Click()
    Application.Run MacroName:="DN_Add_PgNo_Tags_To_DAISY_or_NIMAS"
End Sub

Private Sub Cancel_Button_Click()
    Unload Me
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

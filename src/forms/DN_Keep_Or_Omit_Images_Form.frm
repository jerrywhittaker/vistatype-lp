VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DN_Keep_Or_Omit_Images_Form 
   Caption         =   "Large Print or Braille (325)"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "DN_Keep_Or_Omit_Images_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DN_Keep_Or_Omit_Images_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
' DN_Keep_Or_Omit_Images_Form  -  "Large Print or Braille (325)"
'
' Version: 2.0  Date: 10/18/2026 - asks LARGE PRINT OR BRAILLE instead of keep or omit images
'                                (Jerry, 10/18/2026, issue #37). Both answers keep the pictures;
'                                braille drops them later, elsewhere, so the conversion no longer
'                                offers to. The answer decides what happens to a book's MathML
'                                equations. The form keeps its name so nothing that refers to it
'                                breaks; the option buttons are LargePrint and Braille now.
' Version: 1.1  Date: 10/12/2026 - the red X did nothing: QueryClose set Cancel = True,
'                                then unloaded, and Word ignores an Unload there (Jerry, 10/12/2026).
'                                The X now only clears the answer and lets the box close.
' Version: 1.0  Date: 8/5/2026
'
' Author: Jerry Whittaker - jerry@vistatypelp.org
'
' Asked straight after the DAISY/NIMAS choice, by Sh_Convert_XML_File_To_Word_Document.
'
' Answers through Sh_GP_String_2, "LP" or "BRL", the way the type form answers through
' Sh_GP_String_1. An empty string means canceled, and the caller stops.

Private Sub CmdOkay_Click()
    If Braille.Value = True Then
        Sh_GP_String_2 = "BRL"
    Else
        Sh_GP_String_2 = "LP"
    End If

    Unload Me
End Sub

Private Sub CmdCancel_Click()
    Sh_GP_String_2 = vbNullString  'signal cancel
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'The red X is Cancel: it leaves no answer, and the box then closes by itself. Do not set
    'Cancel = True and unload from here - Word ignores an Unload inside QueryClose, so the X
    'did nothing at all (Jerry, 10/12/2026, on 329).
    If CloseMode = vbFormControlMenu Then Sh_GP_String_2 = ""
End Sub

Private Sub UserForm_Initialize()
    'Large print is the default, as keeping the images was before it.
    LargePrint.Value = True

    'Center UserForm within the Word application window
    Me.StartUpPosition = 0

    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

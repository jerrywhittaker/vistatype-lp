VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DN_Keep_Or_Omit_Images_Form 
   Caption         =   "Keep or Omit Images (325)"
   ClientHeight    =   1668
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
' DN_Keep_Or_Omit_Images_Form
'
' Version: 1.1  Date: 10/12/2026 - the red X did nothing: QueryClose set Cancel = True,
'                                then unloaded, and Word ignores an Unload there (Jerry, 10/12/2026).
'                                The X now only clears the answer and lets the box close.
' Version: 1.0  Date: 8/5/2026
'
' Author: Jerry Whittaker - jerry@vistatypelp.org
'
' Asked straight after the DAISY/NIMAS choice, by Sh_Convert_XML_File_To_Word_Document.
' Braille almost never wants the pictures; large print always does. Omitting them also skips
' the longest stage of the conversion by a wide margin. Jerry timed the same book on the same
' VM both ways, 8/5/2026: 5 min 30 sec keeping the images, 2 min 30 sec omitting them.
'
' Answers through Sh_GP_String_2, "KEEP" or "OMIT", the way the type form answers through
' Sh_GP_String_1. An empty string means canceled, and the caller stops.

Private Sub CmdOkay_Click()
    If OmitImages = True Then
        Sh_GP_String_2 = "OMIT"
    Else
        Sh_GP_String_2 = "KEEP"
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
    'Keep images is the default - large print is the common case and the safe one, since
    'omitting throws the pictures away and the conversion would have to be run again.
    KeepImages.Value = True

    'Center UserForm within the Word application window
    Me.StartUpPosition = 0

    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

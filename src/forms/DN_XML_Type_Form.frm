VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DN_XML_Type_Form 
   Caption         =   "Convert DAISY or NIMAS .xml to Word (327)"
   ClientHeight    =   1896
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "DN_XML_Type_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DN_XML_Type_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' DN_XML_Type_Form
'
' Version: 1.2  Date: 9/29/2026 - the title was cut off before its number (issue #22).
'                               The form is 228 points wide inside (ClientWidth 4560), not 207.75, so the
'                               title, measured in Segoe UI 9 point, fits with about 20 pixels to spare. Every
'                               control moved 10.2 points right, so the layout stays centered. Widened in Word's
'                               designer on the build box; the size lives in the .frx as well as in ClientWidth.
'                               Set by measurement, not yet seen on screen.
' Version: 1.1  Date: 3/1/2026 - bug fixes
' Version: 1.0  Date: 2/22/2026
'

Private Sub CmdOkay_Click()
    'Set the choice based on selected options
    If IsDAISY = True Then
        Sh_GP_String_1 = "DAISY"
    ElseIf IsNIMAS = True Then
        Sh_GP_String_1 = "NIMAS"
    Else
        Sh_GP_String_1 = vbNullString  'or keep previous value
    End If

    Unload Me
End Sub

Private Sub CmdCancel_Click()
    Sh_GP_String_1 = vbNullString  'optional: signal cancel
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Handles the red X
    If CloseMode = vbFormControlMenu Then
        Sh_GP_String_1 = vbNullString  'optional: signal cancel
        Unload Me
        Cancel = True                 'prevents default close behavior
    End If
End Sub

Private Sub UserForm_Initialize()
    'Center UserForm within the Word application window
    Me.StartUpPosition = 0

    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

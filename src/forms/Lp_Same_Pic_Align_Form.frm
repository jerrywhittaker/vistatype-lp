VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Same_Pic_Align_Form 
   Caption         =   "Align These Pictures (374)"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "Lp_Same_Pic_Align_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Same_Pic_Align_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Same_Pic_Align_Form
'
' Version: 1.0  Date: 9/17/2026
'
' Asked by Lp_Resize_Same_Picture_Throughout once it knows which picture it is working from
' (Jerry, 9/17/2026). The copies it is about to resize should all sit the same way on the page,
' and it cannot be guessed: a lesson icon repeated through a book is usually left aligned, a
' full-width illustration usually centered.
'
' The answer goes back in Sh_GP_String_1 - the same way DN_XML_Type_Form hands back its choice -
' as "LEFT", "CENTER", or EMPTY for cancelled. Empty is what the caller tests, so closing this
' with Cancel or the red X stops the macro rather than quietly picking one.
'
' The choice is read BEFORE Unload Me. Touching any member of a UserForm after it is unloaded
' creates the form again and reads its design-time values, which is a fault this project has
' looked at once already (Change Picture Color, 9/4/2026).

Private Sub UserForm_Initialize()
    ' Left align is offered first: a picture repeated throughout a book is usually an icon
    ' sitting beside text, and that is the commoner answer.
    LeftAlignButton.Value = True

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub CmdOkay_Click()
    If CenterButton.Value Then
        Sh_GP_String_1 = "CENTER"
    Else
        Sh_GP_String_1 = "LEFT"
    End If
    Unload Me
End Sub

Private Sub CmdCancel_Click()
    Sh_GP_String_1 = vbNullString
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' The red X means the same as Cancel. Nothing is done to stop the close - the caller reads
    ' an empty Sh_GP_String_1 and gives up, which is what pressing Cancel does too.
    If CloseMode = vbFormControlMenu Then Sh_GP_String_1 = vbNullString
End Sub

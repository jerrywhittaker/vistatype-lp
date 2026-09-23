VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_TOC_Format_And_Color_Form 
   Caption         =   "Format TOC - Add/Remove Color Bars (354)"
   ClientHeight    =   2628
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4095
   OleObjectBlob   =   "Lp_TOC_Format_And_Color_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_TOC_Format_And_Color_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Lp_TOC_Format_And_Color_Form
'
' Version: 2.0  Date: 9/22/2026 - THE BOX STAYS OPEN, holding the TOC range until Done. Jerry,
'                               9/22/2026: the same kind of box as Resize Pictures in a Selected
'                               Range (375). Shown vbModeless by Lp_Tocb_Start; Cancel is now Done.
'                               Every button calls a macro in LPandBrlMacros and nothing here
'                               touches the book - Remove Color Bars moved out to
'                               Lp_TOC_Remove_Color_Bars. Every End is gone: End wipes the range
'                               the macro holds, and unloads this box under the transcriber.
' Version: 1.2  Date: 9/7/2026 - dialog 207 says ONE press again. The hidden scratch document is
'                               back in Lp_TOC_CleanAndFormat_TOC (3.0.402) because doing the work
'                               in the book cost 39 presses, measured. Jerry: "multiple (i mean
'                               more then 3 or 4) is not an acceptable undo requirment."
' Version: 1.1  Date: 9/2/2026 - dialog 207 no longer promises a number of Ctrl+Z presses.
'                               Lp_TOC_CleanAndFormat_TOC stopped round-tripping through a
'                               scratch document, so the two presses that bought are gone; and
'                               it CANNOT have a custom undo record instead - one crashed Word
'                               outright. See the note in that macro.
' Version: 1.0  Date: 8/25/2025
'
' QueryClose sends the title bar's X through the same Done as the button, or the box would be
' unloaded behind the macro's back with the range still held.

Private Sub CancelButton_Click()
    Lp_Tocb_Done
End Sub

Private Sub OkayButton_Click()
    Dim job As Long

    If FormatTheTOCButton Then
        job = 1
    ElseIf AddColorBarsButton Then
        job = 2
    ElseIf RemoveColorBarsButton Then
        job = 3
    End If
    Lp_Tocb_Okay job
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Lp_Tocb_Done
    End If
End Sub

Private Sub UserForm_Initialize()

    Lp_GP_String_1 = ActiveDocument.Name

    ' CANCEL IS DONE - Jerry, 9/22/2026. The box stays open until this is pressed, so there is
    ' nothing to cancel. Set here so it reads in the source rather than only in the .frx; the
    ' control keeps its name so the .frx does not have to change.
    CancelButton.Caption = "Done"
    CancelButton.Accelerator = "D"

    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

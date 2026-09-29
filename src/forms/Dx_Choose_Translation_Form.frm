VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Choose_Translation_Form 
   Caption         =   "What braille translation will be used for this document? (331)"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5580
   OleObjectBlob   =   "Dx_Choose_Translation_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Choose_Translation_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dx_Choose_Translation_Form
' Version: 1.5 Date: 9/29/2026 - Cancel stops the whole braille attach (Jerry). The attach now shows
'                               this box before it changes anything, so UserForm_Initialize no
'                               longer runs MS_Set_Word_Config_For_Braille - on Cancel it would have
'                               left Word set up for braille on a book that is not braille
' Version: 1.4 Date: 9/29/2026 - two faults Jerry reported on 3.0.519. (1) The title was cut off
'                               before its number: the form is 279 points wide inside (ClientWidth
'                               5580), not 267.75, and the controls sit 4 points right of their
'                               3.0.519 positions. Cancel is now a full-width bar under the four
'                               table buttons (Left 12, Width 250). Layout by Jerry in Word,
'                               9/29/2026, and he saw the whole title, (331) included, on screen.
'                               The size lives in the .frx as well as in ClientWidth above.
'                               (2) Cancel looped for ever:
'                               the braille attach showed this form in a Do While that only a table
'                               button could end. The loop is gone (see Dx_Attach_BANA_Template_Run),
'                               and so is Dx_UEB_EBAE_String, which each button set only to end it
' Version: 1.3 Date: 9/21/2026 - a Cancel button, at last. Cmd_Cancel_Click has been here since
'                               at least 3/9/2024, but no Cmd_Cancel was on the form, so the only
'                               way out was the X. Added in Word's designer on the build box: Cancel,
'                               Alt+C, Tahoma 12 like the four table buttons, Cancel = True so
'                               Esc closes the dialog. Under BANA EBAE Nemeth; form 36 points taller
' Version: 1.2 Date: 8/27/2026 - the four translation buttons call Dx_Set_BrailleType instead of
'                               writing the BrailleType document variable themselves, so the type
'                               and SWIFT's own DBTTemplate property are recorded together and in
'                               one place. Dx_UEB_EBAE_String is still set first, on each button's
'                               own first line - two Do While loops depend on it to terminate
' Version: 1.1 Date: 3/9/2024 - removed "End" from Sub Cmd_Cancel_Click
'
Private Sub Cmd_Cancel_Click()
    ' Records nothing, so Dx_Translation_Answered stays down and whatever showed this box stops:
    ' the braille attach leaves the book as it was, and a braille macro goes no further (Jerry,
    ' 9/29/2026: Cancel "should stop the entire attachment process").
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_EBAE_Button_Click()
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ' Records the type AND, when the document does not already carry one, SWIFT's own
    ' DBTTemplate property - so a book prepared here is legible to SWIFT too. These four
    ' buttons wrote the variable themselves until 8/27/2026, which is why Dx_Set_BrailleType
    ' said it was "one place" while being two. See it for why the property is absent-only.
    Dx_Set_BrailleType "EBAT"   'place EBAE Textbook type into document info
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_EBAE_Nemeth_Button_Click()
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ' Records the type AND, when the document does not already carry one, SWIFT's own
    ' DBTTemplate property - so a book prepared here is legible to SWIFT too. These four
    ' buttons wrote the variable themselves until 8/27/2026, which is why Dx_Set_BrailleType
    ' said it was "one place" while being two. See it for why the property is absent-only.
    Dx_Set_BrailleType "EBAN"   'place EBAE Nemeth type into document info
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_UEB_Button_Click()
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ' Records the type AND, when the document does not already carry one, SWIFT's own
    ' DBTTemplate property - so a book prepared here is legible to SWIFT too. These four
    ' buttons wrote the variable themselves until 8/27/2026, which is why Dx_Set_BrailleType
    ' said it was "one place" while being two. See it for why the property is absent-only.
    Dx_Set_BrailleType "UEBT"   'place UEB Textbook type into document info
    Unload Me 'close this form
End Sub

Private Sub Cmd_BANA_UEB_Nemeth_Button_Click()
    ' if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next
    BrlType = ActiveDocument.Variables("Undefined")
    ' Records the type AND, when the document does not already carry one, SWIFT's own
    ' DBTTemplate property - so a book prepared here is legible to SWIFT too. These four
    ' buttons wrote the variable themselves until 8/27/2026, which is why Dx_Set_BrailleType
    ' said it was "one place" while being two. See it for why the property is absent-only.
    Dx_Set_BrailleType "UEBN"   'place UEB Nemeth type into document info
    Unload Me 'close this form
End Sub



Private Sub UserForm_Initialize()

    ' No MS_Set_Word_Config_For_Braille here any more (9/29/2026). The braille attach now shows this
    ' box BEFORE it attaches anything, and that sub writes to the book and saves the transcriber's
    ' own settings - so Cancel would have left a book that is not braille with Word set up for
    ' braille. The attach configures Word itself once the answer is in; on the other route,
    ' Dx_Is_BANA_Template_Attached now does it after an answer (for a book whose template was put
    ' on while it was on screen, which nothing else notices until the window changes).
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

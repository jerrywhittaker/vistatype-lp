VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Valid_Ref_Pg_No_1_Form 
   Caption         =   "How to validate the $pg tags (363)"
   ClientHeight    =   3264
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10800
   OleObjectBlob   =   "Sh_Valid_Ref_Pg_No_1_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Valid_Ref_Pg_No_1_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Sh_Valid_Ref_Pg_No_1_Form
' Provides directions for operation in the temp doc to validate the $pg tags in the main document
'
' Version 1.2: Date: 8/23/2026 - Alt+C presses Close (Jerry); the directions label grows to fit
'                                its own text, so two more lines cannot clip
' Version 1.1: Date: 8/23/2026 - the F6 / Shift+F6 instruction, first and on its own line (Jerry)
' Version 1.0: Date: 7/27/2026
'

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim chrome As Single
    Me.StartUpPosition = 0

    ' Alt+C presses Close. Jerry, 8/23/2026 - the same accelerator Cancel carries everywhere else
    ' in the product, and this dialog's Close button does the same job.
    '
    ' Set here rather than in the binary layout file, which is where an accelerator normally
    ' lives, because it can be read here: this whole change is for people working by keyboard, and
    ' a keyboard rule that is invisible in the source is a keyboard rule nobody maintains.
    CloseButton.Accelerator = "C"
    
    ' The keyboard instruction comes FIRST, on its own line. Jerry, 8/23/2026: it is here for a
    ' transcriber who cannot use a mouse, and a screen reader reads from the top - buried at the
    ' end it would be the last thing she heard rather than the first thing she needed.
    Directions.Caption = "Use F6 or Shift+F6 to move between the document and the action menu." _
    & vbCrLf & vbCrLf & "This is a temporary document to facilitate validation of the '$pg' reference page tags." _
    & " This document can be closed without saving when the validation is complete by selecting the 'Done - Exit validation' button." _
    & vbCrLf & vbCrLf & "Use Up/Down arrow keys to move the cursor to the $pg item to be located in the document then click the " _
    & "'Locate the selected $pg code in the Document' button." _
    & vbCrLf & vbCrLf & "To add a missing $pg tag, select the $pg tag preceding the missing tag, " _
    & "locate it in the document and scroll the document to the point where the new tag should be located, create an empty paragraph, " _
    & "type the page number and tag it using the 'Manual Tag Ref Page' macro on the Ribbon."

    ' THE DIRECTIONS ARE A LABEL, AND A LABEL CLIPS SILENTLY - text past the bottom edge is
    ' simply not drawn, and there is nothing to scroll. They just gained two lines, so the label
    ' is grown to fit whatever is actually in it and the dialog is grown to match. Done this way
    ' rather than by picking a new height once: the wording here gets edited, and nobody should
    ' have to measure it. With WordWrap already on, AutoSize grows the HEIGHT at the same width.
    Directions.WordWrap = True
    Directions.AutoSize = True

    CloseButton.Top = Directions.Top + Directions.Height + 12
    ' The title bar and border, which are not this dialog's to lay out. RANGE-CHECKED, because
    ' this runs before the form has ever been on screen: InsideHeight answering 0 here would take
    ' the whole of Me.Height off and open a dialog about twice as tall as it should be, with the
    ' Close button below the bottom of the screen. 29 points is what the build box measures.
    chrome = Me.Height - Me.InsideHeight
    If chrome < 6 Or chrome > 120 Then chrome = 30

    Me.Height = CloseButton.Top + CloseButton.Height + 12 + chrome

    ' Centered LAST, so it centers the height the dialog actually ended up with.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

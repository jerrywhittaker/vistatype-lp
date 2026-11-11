VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Prodnote_Info_Form 
   Caption         =   "Prodnotes in This Document (361)"
   ClientHeight    =   6048
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "Sh_Prodnote_Info_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Prodnote_Info_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Close_Me_Click()
    Unload Me 'close this form
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

Private Sub UserForm_Initialize()
'
' Fills in the note, then centers the form on the Word window.
'
' The text lives here rather than in the form's binary layout file so that it can be read and
' edited as text, in the VBE or in src/forms.
'
' Version: 1.5  Date: 11/11/2026 - note replaced with Jerry's new text: one paragraph each for
'                                large print (Prodnote Review) and braille (Prodnotes and Alt-Text)
' Version: 1.4  Date: 11/7/2026 - the braille paragraph sent the reader to Prodnote to TN and
'                                Delete Prodnotes in the File Cleanup group, which moved onto
'                                Prodnotes and Alt-Text (417) on 11/6/2026. It names that button and
'                                its Review/Edit/Delete Prodnotes now, in about the same length
' Version: 1.3  Date: 10/28/2026 - the note's last lines were hidden below the box (Jerry, 3.0.619,
'                                "the text is being cut off at the bottom by the button"): measured
'                                on the build box, the box is 254.4 points high, about 20 lines of
'                                Tahoma 10.2, and the note needs about 21 to 22 (estimated). The box
'                                is made 42 points (about 3 lines) taller, and the button and the
'                                form's bottom move down with it. The Close button gets Alt+C
'                                (Jerry); it had no Alt key
' Version: 1.2  Date: 8/8/2026 - the note now points braille transcribers at Delete Prodnotes on
'                                their OWN tab (it is on both), and names the group the way the
'                                ribbon does - "File Cleanup", not "File Cleanup Group". The
'                                quotation marks say "this is the exact name", so they have to
'                                match what the transcriber is looking at (Jerry)
' Version: 1.1  Date: 8/8/2026 - message rewritten; form made taller to suit. The arrow that
'                                was written as ChrW(8594) is gone - the ribbon button now
'                                reads "Prodnote to TN" - so this file is plain Windows-1252
'                                again and no escape is needed (Jerry)
'
    Dim m As String

    m = "This Word document contains one or more production notes. The prodnotes are formatted " _
      & "with a Prodnote style and are displayed in red text within the document. Prodnotes are " _
      & "not part of the publisher's printed version of the book but are added during the " _
      & "production of the DAISY or NIMAS file. Prodnotes can be used to insert text into the " _
      & "document to improve accessibility and are often descriptions of visual elements such " _
      & "as complex pictures or mathematical charts and graphs."

    m = m & vbCrLf & vbCrLf _
      & "The prodnote content can be viewed for its usefulness and modified or deleted depending " _
      & "on whether the document is for large print or braille. When the desired output is large " _
      & "print, prodnotes may have value only when the goal is to produce an accessible PDF or web " _
      & "page. When the desired output is braille, prodnotes may provide descriptions of visual " _
      & "elements or details of mathematical equations, providing readers with information not " _
      & "found in the text of the book."

    m = m & vbCrLf & vbCrLf _
      & "FOR LARGE PRINT, the prodnote tools can be accessed by selecting the ""Prodnote Review"" " _
      & "icon in the ""File Cleanup"" group on the VistaType LP tab ribbon. The tools provide the " _
      & "ability to review, edit, and delete prodnotes or to convert each or all prodnotes into " _
      & "Box Red or Box Blue styles."

    m = m & vbCrLf & vbCrLf _
      & "FOR BRAILLE, the prodnote tools can be accessed by selecting the ""Prodnotes and Alt-Text"" " _
      & "icon in the ""File Cleanup"" group on the ""Braille Macros"" tab ribbon. The tools provide " _
      & "the ability to review, edit, and delete prodnotes or to convert prodnotes to a transcriber " _
      & "note (TN)."

    Info_Text.Text = m

    ' Put the caret at the start rather than the end, so a note long enough to scroll opens
    ' at its first line.
    Info_Text.selStart = 0

    ' Room for the whole note without scrolling: about 3 lines more than the layout file gives,
    ' with the button and the form's bottom moved down to match.
    Info_Text.Height = Info_Text.Height + 42
    Close_Me.Top = Close_Me.Top + 42
    Me.Height = Me.Height + 42

    ' Close, with Alt+C (Jerry, 10/28/2026).
    Close_Me.Caption = "Close"
    Close_Me.Accelerator = "C"

    ' Start the form centered inside the Word window, which also puts it on the right screen
    ' when there are two. Same approach as Sh_Hyperlink_Info_Form.
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Prodnote_Info_Form 
   Caption         =   "Prodnotes in This Document"
   ClientHeight    =   4056
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
' edited as text, in the VBE or in src/forms. The arrow in "Prodnote -> TN" is written as
' ChrW(8594): it is the same character the ribbon button carries (&#8594;), but it is NOT in
' the Windows-1252 set this file is exported as, so a literal one would not survive a build.
'
' Version: 1.0  Date: 8/1/2026
'
    Dim m As String

    m = "This Word document, created from a DAISY or NIMAS file, contains one or more " _
      & "prodnotes (production notes). These prodnotes are formatted with a Prodnote " _
      & "style and are displayed in red text within the document. Prodnotes are not part " _
      & "of the publisher's printed version of the book but are added during the " _
      & "production of the DAISY or NIMAS file. Prodnotes can be used to insert printable " _
      & "information to improve accessibility and are often descriptions of visual " _
      & "elements such as complex pictures and mathematical charts and graphs."

    m = m & vbCrLf & vbCrLf _
      & "When the desired output is braille, descriptions of visual elements can provide " _
      & "readers with additional information not in print. Edit the prodnotes as needed. " _
      & "The prodnotes in the file may be converted to the ""Transcriber Note"" style " _
      & "using the ""Prodnote " & ChrW(8594) & " TN"" macro after editing is completed."

    m = m & vbCrLf & vbCrLf _
      & "For large print, prodnotes may only have value if the aim is to produce an " _
      & "accessible PDF or web page."

    m = m & vbCrLf _
      & "In either case, the prodnotes should be viewed for their usefulness. Prodnotes " _
      & "that are not useful can be deleted manually, leaving only the useful ones, or all " _
      & "prodnotes can be deleted from the document by running the ""Delete Prodnotes"" " _
      & "macro from the ""File Cleanup"" group of either the ""Braille Macros"" tab or " _
      & """VistaType LP"" tab."

    Info_Text.Text = m

    ' Put the caret at the start rather than the end, so a note long enough to scroll opens
    ' at its first line.
    Info_Text.selStart = 0

    ' Start the form centered inside the Word window, which also puts it on the right screen
    ' when there are two. Same approach as Sh_Hyperlink_Info_Form.
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

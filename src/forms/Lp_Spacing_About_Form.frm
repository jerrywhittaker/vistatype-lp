VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Spacing_About_Form 
   Caption         =   "About Line and Letter Spacing (429)"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "Lp_Spacing_About_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Spacing_About_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Spacing_About_Form - About Line and Letter Spacing (429), opened by the About line/letter
' spacing button (Alt+A) on Spacing and Italics Settings (424). A read-only note: Jerry's
' explanation of the spacing standards, the default and the italics choices.
'
' The note runs to about 3,200 characters, past what a MsgBox shows (about 1,024), so it is a
' form. The text is set in UserForm_Initialize, so it can be read here and not only in the .frx.
' The box scrolls; it opens at the first line. Okay (Alt+O, Enter, Esc) and the red X only close
' it - there is nothing to keep, so a plain close is already Cancel and the form has no
' UserForm_QueryClose.
'
' Version: 1.0  Date: 11/26/2026 - Jerry's text

Private Sub Close_Me_Click()
    Unload Me
End Sub

' The note, paragraph by paragraph. A heading is a paragraph of its own (a text box has no bold).
Private Function AboutText() As String
    Dim m As String

    m = "Using the ""Spacing and Italics Settings"""
    m = m & vbCrLf & vbCrLf _
      & "The two most important readability factors for large print are font size and white " _
      & "space. White space includes margin sizes, line spacing (the size of white space " _
      & "between lines of text), and letter spacing (the size of white space between letters " _
      & "within a line). This form provides a method for increasing or decreasing the size of" _
      & " the white space between lines of text and/or the white space between letters. In " _
      & "addition, it provides the ability to choose color text to indicate the use of " _
      & "italics in place of the dashed underline."
    m = m & vbCrLf & vbCrLf _
      & "The size of the white space is proportionate to the size of the font used in the " _
      & "document; the larger the font, the larger the white space for both line spacing and " _
      & "letter spacing. VistaType LP handles this automatically and it shows the point size " _
      & "of that spacing based on the chosen point size."
    m = m & vbCrLf & vbCrLf _
      & "Choice Buttons Controlling White Space Size"
    m = m & vbCrLf & vbCrLf _
      & """VTLP V2 (APH)"" - the starting settings. The ""VistaType LP Version 2 (APH)"" " _
      & "setting is close to the spacing in earlier versions of VistaType LP, and close to " _
      & "the spacing used by the American Printing House for the Blind."
    m = m & vbCrLf & vbCrLf _
      & """W3C WCAG"" - These white space settings match the World Wide Web Consortium's Web " _
      & "Content Accessibility Guidelines and produce white space at the double-A (AA) level." _
      & " These settings give more white space than the ""VTLP V2 (APH)"" settings."
    m = m & vbCrLf & vbCrLf _
      & "When you attach the large print template to a new document, it uses the settings " _
      & "from this form."
    m = m & vbCrLf & vbCrLf _
      & "A large print book keeps its own line and letter spacing. Editing the book does not " _
      & "change it, and neither does re-attaching the template, unless you pick new settings " _
      & "in this form during the re-attach."
    m = m & vbCrLf & vbCrLf _
      & "An older book that has no settings of its own gets your default when you re-attach " _
      & "the template. If you have never saved a default, it gets the VTLP V2 (APH) settings."
    m = m & vbCrLf & vbCrLf _
      & "Save as My Default"
    m = m & vbCrLf & vbCrLf _
      & "Choose this when you want these settings for most of your books."
    m = m & vbCrLf & vbCrLf _
      & "VistaType LP uses them for this book, and it also remembers them. From then on, they" _
      & " are your starting settings for every new book you attach."
    m = m & vbCrLf & vbCrLf _
      & "Your default does not change a book that already has its own settings. That book " _
      & "keeps what it has. To change it, open this form when you re-attach the template, and" _
      & " pick new settings."
    m = m & vbCrLf & vbCrLf _
      & "Your default is saved on this computer, for your Windows sign-in only. Someone else " _
      & "who signs in on the same computer has their own default."
    m = m & vbCrLf & vbCrLf _
      & "Note: Nothing is saved until you press Okay on Attach Template Form (338). If you " _
      & "press Cancel there, your default stays as it was."
    m = m & vbCrLf & vbCrLf _
      & "This Document Only"
    m = m & vbCrLf & vbCrLf _
      & "Choose this for one special book. The settings go into this book and nowhere else. " _
      & "Your default stays the same."
    m = m & vbCrLf & vbCrLf _
      & "Italics"
    m = m & vbCrLf & vbCrLf _
      & "Large print does not use italic type, because slanted letters are harder to read. " _
      & "VistaType LP takes the slant away and marks those words another way."
    m = m & vbCrLf & vbCrLf _
      & "The starting choice is a dashed underline. You can choose a color instead: blue, " _
      & "green, orange, red or violet. The words then appear in that color, with no " _
      & "underline."
    m = m & vbCrLf & vbCrLf _
      & "Caution: Choose a color the book does not already use for something else. If the " _
      & "book has words in green for another reason, green italics would look the same as " _
      & "those words, and the reader could not tell them apart."

    AboutText = m
End Function

Private Sub UserForm_Initialize()
    Info_Text.Text = AboutText()
    ' Open at the first line, not the last.
    Info_Text.selStart = 0

    ' Centered inside the Word window, which also puts it on the right screen when there are
    ' two. Same approach as Sh_Prodnote_Info_Form.
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

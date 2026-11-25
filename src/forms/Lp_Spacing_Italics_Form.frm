VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Spacing_Italics_Form 
   Caption         =   "Spacing and Italics Settings (424)"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   OleObjectBlob   =   "Lp_Spacing_Italics_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Spacing_Italics_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Spacing_Italics_Form - Spacing and Italics Settings (424), opened by the Spacing and Italics
' Settings button (Alt+T) on Attach LP Template & Select Output Media (338)
'
' Three large print choices for the book being attached (Jerry, 11/24/2026): the line spacing (a
' Word Multiple, 1.00 to 2.00), the letter spacing (a share of the type size, 0.00 to 0.25), and
' how italics are shown (dashed underline, or one of the five Text color styles).
'
' The two spacing rows each have two buttons that put a standard's value in the list (Jerry,
' 11/25/2026): W3C WCAG, 1.25 and 0.12 (Alt+W line, Alt+G letter), and VTLP V2 (APH), 1.00 and
' 0.07 (Alt+V line, Alt+P letter) - the APH values are the built-in ones the box starts with. The
' italics row's Reset puts the built-in marker back (Alt+S). The row labels' Alt keys (Alt+L,
' Alt+T, Alt+I) move to their lists. "This document only" (Alt+D) or "Save as my default" (Alt+M)
' covers all three. "About line/letter spacing" (Alt+A) explains the two standards.
'
' The wording of the labels and buttons is set in UserForm_Initialize, so it can be read here and
' not only in the .frx.
'
' The box opens on the book's own values (Lp_Pref_Line_Multiple and the rest: the book's, else
' the saved default, else built in), or on what was chosen here already when it is reopened
' before 338's Okay. The "= 27 pt at 18 pt" beside the two spacings uses the type size picked on
' 338 (Lp_Base_Font_Size), 18 when none is.
'
' Okay hands the choices to the attach in the Lp_Pref_Chosen publics and writes nothing: the book
' and VistaType.ini are written by Lp_Attach_The_Template, after 338's Okay, so Cancel on 338
' still changes nothing. Cancel and the red X leave the publics alone - a plain close is already
' Cancel, so the form has no UserForm_QueryClose.
'
' Version: 1.1  Date: 11/25/2026 - W3C WCAG and VTLP V2 (APH) buttons on both spacing rows (the
'                               Reset buttons became W3C WCAG), About line/letter spacing, and
'                               the box starts on the APH values (Jerry); wider box
' Version: 1.0  Date: 11/24/2026

Private Function TypeSizeNow() As Double
    TypeSizeNow = Val(Lp_Base_Font_Size)
    If TypeSizeNow <= 0 Then TypeSizeNow = 18
End Function

' The "= nn pt at nn pt" text beside the line and letter spacing lists.
Private Sub ShowPoints()
    Dim vals As Variant
    Dim size As Double

    size = TypeSizeNow()
    vals = Lp_Pref_Line_Values()
    If LineList.ListIndex >= 0 Then
        LinePoints.Caption = Lp_Pref_Line_Label(vals(LineList.ListIndex), size)
    Else
        LinePoints.Caption = ""
    End If
    vals = Lp_Pref_Letter_Values()
    If LetterList.ListIndex >= 0 Then
        LetterPoints.Caption = Lp_Pref_Letter_Label(vals(LetterList.ListIndex), size)
    Else
        LetterPoints.Caption = ""
    End If
End Sub

Private Sub LineList_Change()
    ShowPoints
End Sub

Private Sub LetterList_Change()
    ShowPoints
End Sub

' W3C WCAG (the buttons were Reset until 11/25/2026). Like the other buttons in the rows, they
' only change the list - nothing is kept until Okay.
Private Sub LineReset_Click()
    LineList.ListIndex = Lp_Pref_Line_Index(Lp_Pref_Wcag_Line_Multiple())
End Sub

Private Sub LetterReset_Click()
    LetterList.ListIndex = Lp_Pref_Letter_Index(Lp_Pref_Wcag_Letter_Spacing())
End Sub

' VTLP V2 (APH) - the built-in values (Jerry, 11/25/2026).
Private Sub LineAph_Click()
    LineList.ListIndex = Lp_Pref_Line_Index(Lp_Pref_Builtin_Line_Multiple())
End Sub

Private Sub LetterAph_Click()
    LetterList.ListIndex = Lp_Pref_Letter_Index(Lp_Pref_Builtin_Letter_Spacing())
End Sub

' The explanation of the two standards is still to come from Jerry (11/25/2026).
Private Sub AboutSpacing_Click()
    Sh_Say "The explanation of the line and letter spacing standards is still to come.", _
           "VistaType LP (429)"
End Sub

Private Sub MarkerReset_Click()
    MarkerList.ListIndex = Lp_Pref_Marker_Index(Lp_Pref_Builtin_Italics_Marker())
End Sub

Private Sub OkayButton_Click()
    Dim lineVals As Variant
    Dim letterVals As Variant
    Dim markers As Variant

    lineVals = Lp_Pref_Line_Values()
    letterVals = Lp_Pref_Letter_Values()
    markers = Lp_Pref_Marker_Names()

    If LineList.ListIndex >= 0 Then
        Lp_Pref_Chosen_Line = lineVals(LineList.ListIndex)
    Else
        Lp_Pref_Chosen_Line = Lp_Pref_Builtin_Line_Multiple()
    End If
    If LetterList.ListIndex >= 0 Then
        Lp_Pref_Chosen_Letter = letterVals(LetterList.ListIndex)
    Else
        Lp_Pref_Chosen_Letter = Lp_Pref_Builtin_Letter_Spacing()
    End If
    If MarkerList.ListIndex >= 0 Then
        Lp_Pref_Chosen_Marker = markers(LBound(markers) + MarkerList.ListIndex)
    Else
        Lp_Pref_Chosen_Marker = Lp_Pref_Builtin_Italics_Marker()
    End If
    Lp_Pref_Chosen_AsDefault = (SaveAsDefault.Value = True)
    Lp_Pref_Chosen = True
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me                         'nothing chosen here reaches the attach
End Sub

Private Sub UserForm_Initialize()
    Dim v As Variant
    Dim doc As Document
    Dim iLine As Long
    Dim iLetter As Long
    Dim iMarker As Long

    ' The words on the box, here rather than only in the .frx (Jerry, 11/25/2026).
    LineLabel.Caption = "Line spacing"
    LetterLabel.Caption = "Letter spacing"
    LineReset.Caption = "W3C WCAG"
    LineReset.Accelerator = "W"
    LineReset.ControlTipText = "W3C WCAG Line Spacing Button"
    LetterReset.Caption = "W3C WCAG"
    LetterReset.Accelerator = "G"
    LetterReset.ControlTipText = "W3C WCAG Letter Spacing Button"
    LineAph.Caption = "VTLP V2 (APH)"
    LineAph.Accelerator = "V"
    LineAph.ControlTipText = "VTLP V2 (APH) Line Spacing Button"
    LetterAph.Caption = "VTLP V2 (APH)"
    LetterAph.Accelerator = "P"
    LetterAph.ControlTipText = "VTLP V2 (APH) Letter Spacing Button"
    AboutSpacing.Caption = "About line/letter spacing"
    AboutSpacing.Accelerator = "A"
    AboutSpacing.ControlTipText = "About Line/Letter Spacing Button"

    LineList.Clear
    For Each v In Lp_Pref_Line_Values()
        LineList.AddItem Lp_Pref_Number_Text(v)
    Next v
    LetterList.Clear
    For Each v In Lp_Pref_Letter_Values()
        LetterList.AddItem Lp_Pref_Number_Text(v)
    Next v
    MarkerList.Clear
    For Each v In Lp_Pref_Marker_Names()
        MarkerList.AddItem Lp_Pref_Marker_Caption(v)
    Next v

    If Lp_Pref_Chosen Then
        ' Reopened before 338's Okay: show what was chosen here, not the book's values again.
        iLine = Lp_Pref_Line_Index(Lp_Pref_Chosen_Line)
        iLetter = Lp_Pref_Letter_Index(Lp_Pref_Chosen_Letter)
        iMarker = Lp_Pref_Marker_Index(Lp_Pref_Chosen_Marker)
        If Lp_Pref_Chosen_AsDefault Then
            SaveAsDefault.Value = True
        Else
            ThisDocOnly.Value = True
        End If
    Else
        If Documents.Count > 0 Then Set doc = ActiveDocument
        iLine = Lp_Pref_Line_Index(Lp_Pref_Line_Multiple(doc))
        iLetter = Lp_Pref_Letter_Index(Lp_Pref_Letter_Spacing(doc))
        iMarker = Lp_Pref_Marker_Index(Lp_Pref_Attach_Marker(doc, Lp_Doc_Was_Already_LP))   ' a new book shows the saved default
        ThisDocOnly.Value = True
    End If
    If iLine < 0 Then iLine = Lp_Pref_Line_Index(Lp_Pref_Builtin_Line_Multiple())
    If iLetter < 0 Then iLetter = Lp_Pref_Letter_Index(Lp_Pref_Builtin_Letter_Spacing())
    If iMarker < 0 Then iMarker = Lp_Pref_Marker_Index(Lp_Pref_Builtin_Italics_Marker())
    LineList.ListIndex = iLine
    LetterList.ListIndex = iLetter
    MarkerList.ListIndex = iMarker
    ShowPoints

    ' Centered on Word, as 338 is (for two monitors).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

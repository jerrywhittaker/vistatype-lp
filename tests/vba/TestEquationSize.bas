Attribute VB_Name = "TestEquationSize"
'@uses Lp_Size_Equations
'@uses Lp_Eq_Xml_Resize
'@uses Lp_Size_Raised_Lowered
'
' Word equations, superscripts and subscripts at the book's size in large print (Jerry,
' 11/12/2026). First the XML rewrite on its own - a string in, a string out: script runs sized,
' deeper ones more, fractions of an in-line equation written across, nothing else touched. Then a
' hidden scratch book at 24 pt with an in-line fraction, an equation with an exponent on its own
' line, and a plain raised "5": the sizes the pass sets, the fraction type, nothing added or lost,
' and a second run changing nothing. Visible:=False, no Selection, closed unsaved, every error a
' failure (README.md).
'
' Version: 1.2  Date: 11/12/2026 - "3 " and a raised "2": the space is taken out (Jerry)
' Version: 1.1  Date: 11/12/2026 - two equations in one paragraph, both with exponents (the pass now
'                                 reads the book's XML once and cuts each equation out of it)
' Version: 1.0  Date: 11/12/2026
'
Option Explicit

Private Function R(ByVal sz As Long, ByVal t As String) As String
    R = "<m:r><w:rPr><w:sz w:val=""" & sz & """/><w:szCs w:val=""" & sz & """/></w:rPr><m:t>" & t & "</m:t></m:r>"
End Function

Public Function TestEquationSize_Suite() As TestSuite
    Dim Suite As New TestSuite
    Dim x As String, y As String
    Suite.Description = "Word equations, superscripts and subscripts at the book's size"

    With Suite.Test("an exponent is set to the size that prints full size; its base is left alone")
        x = "<m:sSup><m:e>" & R(53, "2") & "</m:e><m:sup>" & R(53, "3") & "</m:sup></m:sSup>"
        y = "<m:sSup><m:e>" & R(53, "2") & "</m:e><m:sup>" & R(73, "3") & "</m:sup></m:sSup>"
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, False), y
    End With

    With Suite.Test("a subscript and a lower limit likewise")
        x = "<m:sSub><m:e>" & R(53, "b") & "</m:e><m:sub>" & R(53, "1") & "</m:sub></m:sSub>"
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, False), Replace(x, R(53, "1"), R(73, "1"))
        x = "<m:limLow><m:e>" & R(53, "x") & "</m:e><m:lim>" & R(53, "_") & "</m:lim></m:limLow>"
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, False), Replace(x, R(53, "_"), R(73, "_"))
    End With

    With Suite.Test("an exponent on an exponent is set larger again (Word draws it at 60%)")
        x = "<m:sup>" & R(53, "a") & "<m:sSup><m:e>" & R(53, "b") & "</m:e><m:sup>" & R(53, "c") & "</m:sup></m:sSup></m:sup>"
        y = "<m:sup>" & R(73, "a") & "<m:sSup><m:e>" & R(73, "b") & "</m:e><m:sup>" & R(88, "c") & "</m:sup></m:sSup></m:sup>"
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, False), y
    End With

    With Suite.Test("an in-line equation's fraction is written across; one on its own line is not")
        x = "<m:f><m:num>" & R(53, "3") & "</m:num><m:den>" & R(53, "4") & "</m:den></m:f>"
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, True), Replace(x, "<m:f>", "<m:f><m:fPr><m:type m:val=""lin""/></m:fPr>")
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, False), x
        x = "<m:f><m:fPr><m:type m:val=""skw""/><m:ctrlPr/></m:fPr><m:num>" & R(53, "3") & "</m:num><m:den>" & R(53, "4") & "</m:den></m:f>"
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, True), Replace(x, "<m:fPr><m:type m:val=""skw""/>", "<m:fPr><m:type m:val=""lin""/>")
    End With

    With Suite.Test("sizes outside the scripts, and look-alike names, are not touched")
        x = "<w:style><w:rPr><w:sz w:val=""48""/></w:rPr></w:style><m:supHide m:val=""1""/>" & R(53, "x")
        .IsEqual Lp_Eq_Xml_Resize(x, 26.5, True), x
    End With

    TestEquationSize_In_A_Book Suite.Test("in a book at 24 pt: equations, fraction and raised text sized, twice the same")

    Set TestEquationSize_Suite = Suite
End Function

' The book: "Take [3/4] of the cake." / [2^3+10^5] / "10" + raised "5" / "[x^2] and [y^3]".
Private Sub TestEquationSize_In_A_Book(ByVal T As TestCase)
    Dim d As Document, r As Range
    Dim stage As String, txt As String, pc As Long, pass As Long

    On Error GoTo Failed
    stage = "making the scratch book"
    Set d = Documents.Add(Visible:=False)
    d.Styles(wdStyleNormal).Font.Size = 24
    d.Content.Text = "Take 3/4 of the cake." & vbCr & "2^3+10^5" & vbCr & "105" & vbCr & "x^2 and y^3" & vbCr & "read 3 2 as"
    d.Content.Font.Size = 24

    Set r = d.Paragraphs(1).Range
    r.SetRange r.Start + 5, r.Start + 8
    r.OMaths.Add r
    d.OMaths(1).BuildUp
    Set r = d.Paragraphs(2).Range
    r.MoveEnd wdCharacter, -1
    r.OMaths.Add r
    d.OMaths(2).BuildUp
    Set r = d.Paragraphs(3).Range
    r.SetRange r.Start + 2, r.Start + 3
    r.Font.Superscript = True
    Set r = d.Paragraphs(4).Range
    r.SetRange r.Start + 8, r.Start + 11
    r.OMaths.Add r
    d.OMaths(3).BuildUp
    Set r = d.Paragraphs(4).Range
    r.SetRange r.Start, r.Start + 3
    r.OMaths.Add r
    d.OMaths(3).BuildUp
    T.IsEqual d.OMaths.Count, 4, "the scratch book has 4 equations"
    Set r = d.Paragraphs(5).Range
    r.SetRange r.Start + 7, r.Start + 8
    r.Font.Superscript = True
    T.IsEqual d.OMaths(1).Type, wdOMathInline, "the fraction is in a line of text"
    pc = d.Paragraphs.Count

    For pass = 1 To 2
        stage = "sizing, pass " & pass
        Lp_Size_Equations d, 24
        T.IsEqual d.OMaths.Count, 4, "pass " & pass & ": still 4 equations"
        T.IsEqual d.OMaths(3).Functions(1).ScrSup.Sup.Range.Font.Size, 36.5, "pass " & pass & ": x's exponent at 36.5"
        T.IsEqual d.OMaths(4).Functions(1).ScrSup.Sup.Range.Font.Size, 36.5, "pass " & pass & ": y's exponent at 36.5"
        T.IsEqual d.OMaths(3).Range.Paragraphs(1).Range.Start, d.OMaths(4).Range.Paragraphs(1).Range.Start, _
                  "pass " & pass & ": x^2 and y^3 still share their paragraph"
        T.IsEqual d.Paragraphs.Count, pc, "pass " & pass & ": no paragraph added or lost"
        T.IsEqual d.OMaths(1).Functions(1).Type, wdOMathFunctionFrac, "pass " & pass & ": the fraction is still a fraction"
        T.IsEqual d.OMaths(1).Functions(1).Frac.Type, wdOMathFracLin, "pass " & pass & ": written across"
        T.IsEqual d.OMaths(1).Range.Font.Size, 26.5, "pass " & pass & ": the fraction at 26.5 (Cambria Math, 24 pt book)"
        T.IsEqual d.OMaths(2).Functions(1).ScrSup.E.Range.Font.Size, 26.5, "pass " & pass & ": 2 at 26.5"
        T.IsEqual d.OMaths(2).Functions(1).ScrSup.Sup.Range.Font.Size, 36.5, "pass " & pass & ": its exponent at 36.5"
        T.IsEqual d.Paragraphs(3).Range.Characters(3).Font.Size, 38.5, "pass " & pass & ": the raised 5 at 38.5"
        T.IsEqual d.Paragraphs(3).Range.Characters(1).Font.Size, 24, "pass " & pass & ": the 10 left at 24"
        T.IsEqual d.Paragraphs(5).Range.Text, "read 32 as" & vbCr, "pass " & pass & ": the space before the raised 2 is gone"
        T.IsEqual d.Paragraphs(5).Range.Characters(7).Font.Size, 38.5, "pass " & pass & ": and the 2 is at 38.5"
        T.IsEqual InStr(1, d.Paragraphs(1).Range.Text, "Take "), 1, "pass " & pass & ": the text before the fraction kept"
        T.IsEqual InStr(1, d.Paragraphs(1).Range.Text, " of the cake."), Len(d.Paragraphs(1).Range.Text) - 13, _
                  "pass " & pass & ": and after it - got " & d.Paragraphs(1).Range.Text
    Next pass

    d.Close SaveChanges:=False
    Exit Sub

Failed:
    T.Fail "VBA error " & Err.Number & " while " & stage & " - " & Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.Close SaveChanges:=False
    Err.Clear
End Sub

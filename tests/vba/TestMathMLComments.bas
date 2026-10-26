Attribute VB_Name = "TestMathMLComments"
'@uses Sh_MathML_Uncomment
'@uses Sh_MathML_With_Namespace
'@uses Sh_MathML_Fragments
'@uses Sh_MathML_Pattern
'
' Equations hidden in XML comments, each straight after the book's own picture of it (issue #40,
' Jerry, 10/25/2026, Bookshare's "Basic Math & Pre-Algebra For Dummies"). Sh_MathML_Uncomment
' brings each one out of its comment, and the picture before it becomes its altimg, so the
' picture goes in only where the equation cannot. The shapes below are the ones measured in
' that book. Also the stray xmlns1:m="..." four of its equations carry, which MSXML refused.
' Nothing here touches a document.
'
' Version: 1.1  Date: 10/25/2026 - every other comment is taken out of the import text too, so no
'                                math left in one is counted (review)
' Version: 1.0  Date: 10/25/2026
'
Option Explicit

Private Function Img(ByVal src As String) As String
    ' The book's picture, as the book writes it.
    Img = "<img src=""" & src & """ alt=""5 times 2"" class=""center"" data-ocr-confidence=""1.0"" " & _
          "data-math-confidence=""0.9999982"" id=""img_00001"" />"
End Function

Private Function Body() As String
    Body = "<m:mrow><m:mn>5</m:mn><m:mo>&#215;</m:mo><m:mn>2</m:mn></m:mrow></m:math>"
End Function

Private Function Decl() As String
    Decl = "xmlns:m=""http://www.w3.org/1998/Math/MathML"""
End Function

Public Function TestMathMLComments_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Equations hidden in comments (DAISY from EPUB)"

    Dim t As String

    With Suite.Test("the picture and the comment after it become one equation with that picture as its altimg")
        t = "<p>FIGURE 1-12</p>" & vbLf & "  " & Img("images/eq01036.png") & vbLf & "  " & _
            "<!--<m:math " & Decl() & ">" & Body() & "-->" & vbLf & "  <span>next</span>"
        .IsEqual Sh_MathML_Uncomment(t, "m:"), _
            "<p>FIGURE 1-12</p>" & vbLf & "  " & _
            "<m:math altimg=""images/eq01036.png"" " & Decl() & ">" & Body() & vbLf & "  <span>next</span>"
    End With

    With Suite.Test("white space of any kind, or none, between the picture and the comment, and inside it")
        .IsEqual Sh_MathML_Uncomment(Img("a.png") & "<!--<m:math>" & Body() & "-->", "m:"), _
                 "<m:math altimg=""a.png"">" & Body()
        .IsEqual Sh_MathML_Uncomment(Img("a.png") & vbCrLf & vbTab & " <!--" & vbCrLf & " <m:math>" & Body() & _
                                     " " & vbCrLf & "-->", "m:"), _
                 "<m:math altimg=""a.png"">" & Body()
        .IsEqual Sh_MathML_Uncomment("<img src='b c.png'/><!--<m:math>" & Body() & "-->", "m:"), _
                 "<m:math altimg=""b c.png"">" & Body()
    End With

    With Suite.Test("a comment holding math with no picture before it is only taken away")
        .IsEqual Sh_MathML_Uncomment("<p>x</p><!-- <m:math>" & Body() & " --><p>y</p>", "m:"), _
                 "<p>x</p><m:math>" & Body() & "<p>y</p>"
        ' text between the picture and the comment: not its picture
        t = Img("a.png") & " and <!--<m:math>" & Body() & "-->"
        .IsEqual Sh_MathML_Uncomment(t, "m:"), Img("a.png") & " and <m:math>" & Body()
    End With

    With Suite.Test("a comment holding anything else is taken out, and the picture before it stays")
        .IsEqual Sh_MathML_Uncomment(Img("a.png") & " <!-- a note --><p>x</p>", "m:"), Img("a.png") & " <p>x</p>"
        .IsEqual Sh_MathML_Uncomment("<p>a</p><!-- see <m:math>" & Body() & " --><p>b</p>", "m:"), "<p>a</p><p>b</p>"
        .IsEqual Sh_MathML_Uncomment("<!--<m:math>" & Body() & " and more-->", "m:"), ""
        .IsEqual Sh_MathML_Uncomment("<!--<m:math><m:mn>1</m:mn></m:math><m:math><m:mn>2</m:mn></m:math>-->", "m:"), ""
        ' math with another prefix is not this book's math
        .IsEqual Sh_MathML_Uncomment("<!--<mml:math><mml:mn>1</mml:mn></mml:math>-->", "m:"), ""
    End With

    With Suite.Test("math left in a comment is not counted: only the equations that reach Word are")
        t = Img("a.png") & "<!--<m:math>" & Body() & "-->" & _
            "<!--<m:math><m:mn>1</m:mn></m:math><m:math><m:mn>2</m:mn></m:math>-->" & _
            "<!-- see <m:math>" & Body() & " -->"
        ' before: all four math elements counted, three of them inside comments the import drops
        .IsEqual UBound(Sh_MathML_Fragments(t, "m:")), 3
        ' after: the one brought out, and no comment left for a placeholder to vanish with
        .IsEqual UBound(Sh_MathML_Fragments(Sh_MathML_Uncomment(t, "m:"), "m:")), 0
        .IsEqual InStr(1, Sh_MathML_Uncomment(t, "m:"), "<!--"), 0
    End With

    With Suite.Test("math that already has an altimg keeps it, and the picture before it stays")
        t = "<m:math altimg=""EQ1.jpg"">" & Body()
        .IsEqual Sh_MathML_Uncomment(Img("a.png") & " <!--" & t & "-->", "m:"), Img("a.png") & " " & t
        ' and live math, not in a comment, is not touched at all
        t = Img("a.png") & " " & t
        .IsEqual Sh_MathML_Uncomment(t, "m:"), t
    End With

    With Suite.Test("a picture on the web or in a data: URI, or with no src, is not made an altimg")
        .IsEqual Sh_MathML_Uncomment("<img src=""http://x.org/a.png""/><!--<m:math>" & Body() & "-->", "m:"), _
                 "<img src=""http://x.org/a.png""/><m:math>" & Body()
        .IsEqual Sh_MathML_Uncomment("<img src=""data:image/png;base64,AAA""/><!--<m:math>" & Body() & "-->", "m:"), _
                 "<img src=""data:image/png;base64,AAA""/><m:math>" & Body()
        .IsEqual Sh_MathML_Uncomment("<img alt=""x""/><!--<m:math>" & Body() & "-->", "m:"), _
                 "<img alt=""x""/><m:math>" & Body()
    End With

    With Suite.Test("a src with an entity is carried over exactly as written")
        .IsEqual Sh_MathML_Uncomment("<img src=""7&amp;8/eq%201.png""/><!--<m:math>" & Body() & "-->", "m:"), _
                 "<m:math altimg=""7&amp;8/eq%201.png"">" & Body()
    End With

    With Suite.Test("each pair in a run of them, and the text between kept")
        t = Img("a.png") & "<!--<m:math>" & Body() & "--><p>and</p>" & Img("b.png") & vbLf & "<!--<m:math>" & Body() & "-->"
        .IsEqual Sh_MathML_Uncomment(t, "m:"), _
                 "<m:math altimg=""a.png"">" & Body() & "<p>and</p><m:math altimg=""b.png"">" & Body()
    End With

    With Suite.Test("no comment at all: the text unchanged")
        .IsEqual Sh_MathML_Uncomment("<p>none</p>", "m:"), "<p>none</p>"
        .IsEqual Sh_MathML_Uncomment("", "m:"), ""
    End With

    With Suite.Test("a stray xmlns1:m declaration is taken off and the real one put on")
        .IsEqual Sh_MathML_With_Namespace("<m:math altimg=""a.png"" xmlns1:m=""http://www.w3.org/1998/Math/MathML"">" & Body(), "m:"), _
                 "<m:math " & Decl() & " altimg=""a.png"">" & Body()
        .IsEqual Sh_MathML_With_Namespace("<m:math xmlns1:m='http://www.w3.org/1998/Math/MathML'>" & Body(), "m:"), _
                 "<m:math " & Decl() & ">" & Body()
        ' beside a real declaration it goes, and the real one stays
        .IsEqual Sh_MathML_With_Namespace("<m:math " & Decl() & " xmlns12:m=""x"">" & Body(), "m:"), _
                 "<m:math " & Decl() & ">" & Body()
        ' and the result loads in MSXML, which refused it before
        Dim d As Object
        Set d = CreateObject("MSXML2.DOMDocument.6.0")
        d.async = False
        .IsEqual d.loadXML(Sh_MathML_With_Namespace("<m:math xmlns1:m=""http://www.w3.org/1998/Math/MathML"">" & _
                           "<m:mrow><m:mn>5</m:mn></m:mrow></m:math>", "m:")), True
        ' only the opening tag is looked at
        t = "<m:math " & Decl() & "><m:mtext>xmlns1:m=""y""</m:mtext></m:math>"
        .IsEqual Sh_MathML_With_Namespace(t, "m:"), t
    End With

    Set TestMathMLComments_Suite = Suite
End Function

Attribute VB_Name = "TestMathML"
'@uses Sh_MathML_Prefix
'@uses Sh_MathML_Placeholder
'@uses Sh_MathML_Fragments
'@uses Sh_MathML_To_Placeholders
'@uses Sh_MathML_Altimg
'@uses Sh_MathML_Picture_Path
'@uses Sh_MathML_With_Namespace
'@uses Sh_MathML_Unprefixed
'@uses Sh_Strip_Xml_Declaration
'@uses Sh_OMML_Package
'@uses Sh_Xml_Start_Tag_Length
'@uses Sh_Xml_Attribute
'@uses Sh_Xml_Unescape
'@uses Sh_Url_Unescape
'@uses Sh_MathML_One_Line
'@uses Sh_MathML_Placeholder_Positions
'@uses Dx_Math_Park_Mark
'@uses Dx_Math_Mark_Number
'@uses Dx_Math_Element_Length
'
' The string helpers behind the DAISY/NIMAS converter's equations (issue #37): finding a book's
' MathML and its prefix, parking each equation as a placeholder, the picture path from altimg,
' the namespace put back for MSXML, and the self-contained MathML text braille gets for DBT.
' Nothing here touches a document. The parts that do - Find, InsertXML, AddPicture - are for
' the headless conversion and for real Word.
'
' From the review of 10/18/2026: start tags read with their quoted values (a ">" in alttext), the
' altimg unescaped as XML and as a URL, the braille MathML kept character for character, and the
' braille cleanups' park - the mark and the rule for what one MathML element is.
'
' Version: 1.2  Date: 10/18/2026 - where the placeholders are, read from the text (the speed fix)
' Version: 1.1  Date: 10/18/2026 - the review's examples
' Version: 1.0  Date: 10/18/2026
'
Option Explicit

Private Function Park_Model(ByVal s As String, ByVal originals As Collection) As String
    ' The walk Dx_Math_Park makes through a book, done on a string: at each "<math", the element
    ' Dx_Math_Element_Length measures is stored and replaced by its mark. A MODEL of the walk,
    ' here in the test - the book version reads the same text through Word's Find.
    Dim pos As Long, at As Long, L As Long, out As String
    pos = 1
    Do
        at = InStr(pos, s, "<math")
        If at = 0 Then Exit Do
        L = Dx_Math_Element_Length(Mid$(s, at))
        If L = 0 Then
            out = out & Mid$(s, pos, at - pos + 5)
            pos = at + 5
        Else
            originals.Add Mid$(s, at, L)
            out = out & Mid$(s, pos, at - pos) & Dx_Math_Park_Mark(originals.count)
            pos = at + L
        End If
    Loop
    Park_Model = out & Mid$(s, pos)
End Function

Private Function Restore_Model(ByVal s As String, ByVal originals As Collection) As String
    Dim i As Long
    For i = 1 To originals.count
        s = Replace(s, Dx_Math_Park_Mark(i), originals(i))
    Next i
    Restore_Model = s
End Function

Private Function NIMAS_Sample() As String
    ' The shape of the Imagine IM A1 7&8 book: the prefix declared once on the root, inline math.
    NIMAS_Sample = "<dtbook xmlns=""http://www.daisy.org/z3986/2005/dtbook/"" " & _
        "xmlns:m=""http://www.w3.org/1998/Math/MathML"" version=""2005-3"">" & _
        "<p>Solve <m:math id=""m1"" display=""inline"" altimg=""./images/U07C01/EQp010-001.jpg"" alttext="""">" & _
        "<m:mi>x</m:mi><m:mo>+</m:mo><m:mn>2</m:mn></m:math> for x.</p>" & _
        "<p>Then <m:math id=""m2"" display=""inline"" altimg=""./images/U07C01/EQp010-002.jpg"" alttext="""">" & vbCrLf & _
        "  <m:msup><m:mi>y</m:mi><m:mn>2</m:mn></m:msup>" & vbCrLf & "</m:math>.</p></dtbook>"
End Function

Public Function TestMathML_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "MathML equations in a DAISY/NIMAS book"

    Dim x As String, f As Variant

    x = NIMAS_Sample()

    ' --- Finding the MathML, and how it is written.
    With Suite.Test("the prefix declared on the root is found")
        .IsEqual Sh_MathML_Prefix(x), "m:"
        .IsEqual Sh_MathML_Prefix(Replace(Replace(x, "m:", "mml:"), "xmlns:m=", "xmlns:mml=")), "mml:"
    End With

    With Suite.Test("MathML as the default namespace on a bare math element is found")
        .IsEqual Sh_MathML_Prefix("<p><math xmlns=""http://www.w3.org/1998/Math/MathML""><mi>x</mi></math></p>"), ""
    End With

    With Suite.Test("a book with no MathML says so")
        .IsEqual Sh_MathML_Prefix("<dtbook><p>No equations here. 2 + 2 = 4</p></dtbook>"), "#"
        .IsEqual Sh_MathML_Prefix(""), "#"
        ' A declaration with no equation after it is no MathML.
        .IsEqual Sh_MathML_Prefix("<dtbook xmlns:m=""http://www.w3.org/1998/Math/MathML""><p>x</p></dtbook>"), "#"
        ' <m:math> under some other namespace is not MathML.
        .IsEqual Sh_MathML_Prefix("<dtbook xmlns:m=""http://example.com/other""><m:math>x</m:math></dtbook>"), "#"
    End With

    ' --- Parking each equation as a placeholder.
    With Suite.Test("placeholders are fixed width, so none is the start of another")
        .IsEqual Sh_MathML_Placeholder(1), "VTEQ000001"
        .IsEqual Sh_MathML_Placeholder(12), "VTEQ000012"
        .IsEqual Sh_MathML_Placeholder(2139), "VTEQ002139"
    End With

    With Suite.Test("every equation is listed, in order, whole")
        f = Sh_MathML_Fragments(x, "m:")
        .IsEqual UBound(f), 1
        .IsEqual Left$(f(0), 15), "<m:math id=""m1"""
        .IsEqual Right$(f(0), 9), "</m:math>"
        .IsEqual Left$(f(1), 15), "<m:math id=""m2"""
        .IsEqual Right$(f(1), 9), "</m:math>"
        .IsOk InStr(f(1), "<m:msup>") > 0, "the second equation runs across lines and is still whole"
    End With

    With Suite.Test("no equations, an empty list")
        f = Sh_MathML_Fragments("<p>none</p>", "m:")
        .IsEqual UBound(f), -1
    End With

    With Suite.Test("each equation becomes its placeholder and nothing else changes")
        .IsEqual Sh_MathML_To_Placeholders(x, "m:"), _
            "<dtbook xmlns=""http://www.daisy.org/z3986/2005/dtbook/"" " & _
            "xmlns:m=""http://www.w3.org/1998/Math/MathML"" version=""2005-3"">" & _
            "<p>Solve VTEQ000001 for x.</p><p>Then VTEQ000002.</p></dtbook>"
        .IsEqual Sh_MathML_To_Placeholders("<p>none</p>", "m:"), "<p>none</p>"
    End With

    With Suite.Test("an empty math element is one equation, not the start of the next")
        f = Sh_MathML_Fragments("<p><m:math id=""e""/> and <m:math id=""f""><m:mi>z</m:mi></m:math></p>", "m:")
        .IsEqual UBound(f), 1
        .IsEqual f(0), "<m:math id=""e""/>"
        .IsEqual Sh_MathML_To_Placeholders("<p><m:math id=""e""/> and <m:math id=""f""><m:mi>z</m:mi></m:math></p>", "m:"), _
            "<p>VTEQ000001 and VTEQ000002</p>"
    End With

    With Suite.Test("<m:mathx> or the like is not taken for a math element")
        .IsEqual UBound(Sh_MathML_Fragments("<p><m:mathx>q</m:mathx></p>", "m:")), -1
    End With

    ' --- The picture an altimg names.
    With Suite.Test("altimg is read from the opening tag")
        f = Sh_MathML_Fragments(x, "m:")
        .IsEqual Sh_MathML_Altimg(f(0)), "./images/U07C01/EQp010-001.jpg"
        .IsEqual Sh_MathML_Altimg("<m:math alttext='a' altimg='pic.jpg'><m:mi>x</m:mi></m:math>"), "pic.jpg"
        .IsEqual Sh_MathML_Altimg("<m:math id=""m3""><m:mi>x</m:mi></m:math>"), ""
    End With

    With Suite.Test("altimg becomes a Windows path in the book folder")
        .IsEqual Sh_MathML_Picture_Path("C:\Books\Imagine\", "./images/U07C01/EQp010-001.jpg"), _
            "C:\Books\Imagine\images\U07C01\EQp010-001.jpg"
        .IsEqual Sh_MathML_Picture_Path("C:\Books\Imagine", "images/EQ 1.jpg"), "C:\Books\Imagine\images\EQ 1.jpg"
        .IsEqual Sh_MathML_Picture_Path("C:\Books\Imagine\", "images/EQ%201.jpg"), "C:\Books\Imagine\images\EQ 1.jpg"
        .IsEqual Sh_MathML_Picture_Path("C:\Books\Imagine\", "D:/pics/eq.jpg"), "D:\pics\eq.jpg"
        .IsEqual Sh_MathML_Picture_Path("C:\Books\Imagine\", ""), ""
    End With

    ' --- The namespace put back, for MSXML (large print, Word equations).
    With Suite.Test("the prefix's namespace goes back on the equation's own tag")
        .IsEqual Sh_MathML_With_Namespace("<m:math id=""m1""><m:mi>x</m:mi></m:math>", "m:"), _
            "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML"" id=""m1""><m:mi>x</m:mi></m:math>"
        .IsEqual Sh_MathML_With_Namespace("<mml:math><mml:mi>x</mml:mi></mml:math>", "mml:"), _
            "<mml:math xmlns:mml=""http://www.w3.org/1998/Math/MathML""><mml:mi>x</mml:mi></mml:math>"
        .IsEqual Sh_MathML_With_Namespace("<math><mi>x</mi></math>", ""), _
            "<math xmlns=""http://www.w3.org/1998/Math/MathML""><mi>x</mi></math>"
    End With

    With Suite.Test("an equation that already declares it is left alone")
        .IsEqual Sh_MathML_With_Namespace("<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mi>x</m:mi></m:math>", "m:"), _
            "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mi>x</m:mi></m:math>"
    End With

    ' --- The MathML text braille gets, for DBT.
    With Suite.Test("braille: the prefix comes off the tags and the namespace goes on the root")
        f = Sh_MathML_Fragments(x, "m:")
        .IsEqual Sh_MathML_Unprefixed(f(0), "m:"), _
            "<math xmlns=""http://www.w3.org/1998/Math/MathML"" id=""m1"" display=""inline"" " & _
            "altimg=""./images/U07C01/EQp010-001.jpg"" alttext=""""><mi>x</mi><mo>+</mo><mn>2</mn></math>"
    End With

    With Suite.Test("braille: text inside the tags is not touched, entities stay as written")
        .IsEqual Sh_MathML_Unprefixed("<m:math><m:mtext>m:s and a:b</m:mtext><m:mi>&#x2113;</m:mi></m:math>", "m:"), _
            "<math xmlns=""http://www.w3.org/1998/Math/MathML""><mtext>m:s and a:b</mtext><mi>&#x2113;</mi></math>"
    End With

    With Suite.Test("braille: an old prefix declaration on the tag goes, and one namespace is left")
        .IsEqual Sh_MathML_Unprefixed("<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML"" id=""a""><m:mi>x</m:mi></m:math>", "m:"), _
            "<math xmlns=""http://www.w3.org/1998/Math/MathML"" id=""a""><mi>x</mi></math>"
        .IsEqual Sh_MathML_Unprefixed("<math xmlns=""http://www.w3.org/1998/Math/MathML""><mi>x</mi></math>", ""), _
            "<math xmlns=""http://www.w3.org/1998/Math/MathML""><mi>x</mi></math>"
    End With

    ' --- The OMML package for InsertXML.
    With Suite.Test("the XML declaration MSXML puts on its answer comes off")
        .IsEqual Sh_Strip_Xml_Declaration("<?xml version=""1.0"" encoding=""UTF-16""?>" & vbCrLf & "<m:oMath/>"), "<m:oMath/>"
        .IsEqual Sh_Strip_Xml_Declaration("<m:oMath/>"), "<m:oMath/>"
    End With

    With Suite.Test("the equation sits inside one paragraph of a Flat-OPC package")
        Dim pkg As String
        pkg = Sh_OMML_Package("<m:oMath>Q</m:oMath>")
        .IsOk Left$(pkg, 13) = "<pkg:package ", "starts with the package"
        .IsOk InStr(pkg, "<w:body><w:p><m:oMath>Q</m:oMath></w:p></w:body>") > 0, "the equation is inside one w:p"
        .IsOk InStr(pkg, "pkg:name=""/_rels/.rels""") > 0, "carries the relationships part"
        .IsOk Right$(pkg, 14) = "</pkg:package>", "ends with the package"
    End With

    ' --- The review, 10/18/2026: start tags with a ">" inside a quoted value.
    With Suite.Test("a start tag ends at its own >, not one inside a quoted value")
        .IsEqual Sh_Xml_Start_Tag_Length("<m:math alttext=""x > 2"">x</m:math>"), 24
        .IsEqual Sh_Xml_Start_Tag_Length("<a b='>' c=""'"">x"), 15
        .IsEqual Sh_Xml_Start_Tag_Length("<a b=""x>"), 0
        .IsEqual Sh_Xml_Start_Tag_Length("no tag"), 0
    End With

    With Suite.Test("an attribute is read as an attribute, never from inside another value")
        .IsEqual Sh_Xml_Attribute("<m:math alttext=""x > 2"" altimg=""EQ1.jpg"">", "altimg"), "EQ1.jpg"
        .IsEqual Sh_Xml_Attribute("<m:math alttext='altimg=""no""' altimg=""yes.jpg"">", "altimg"), "yes.jpg"
        .IsEqual Sh_Xml_Attribute("<m:math alttext = 'a' >", "alttext"), "a"
        .IsEqual Sh_Xml_Attribute("<m:math id=""m1"">", "altimg"), ""
    End With

    With Suite.Test("altimg is found after a raw > in alttext")
        .IsEqual Sh_MathML_Altimg("<m:math alttext=""x > 2"" altimg=""./images/EQ1.jpg""><m:mi>x</m:mi></m:math>"), _
            "./images/EQ1.jpg"
    End With

    With Suite.Test("an empty math element with a > in a value does not swallow the next one")
        f = Sh_MathML_Fragments("<p><m:math alttext=""a>b""/> and <m:math id=""f""><m:mi>z</m:mi></m:math></p>", "m:")
        .IsEqual UBound(f), 1
        .IsEqual f(0), "<m:math alttext=""a>b""/>"
        .IsEqual f(1), "<m:math id=""f""><m:mi>z</m:mi></m:math>"
    End With

    ' --- altimg as XML, then as a URL.
    With Suite.Test("XML entities in an attribute become their characters")
        .IsEqual Sh_Xml_Unescape("7&amp;8/EQ&#x20;1&#46;jpg"), "7&8/EQ 1.jpg"
        .IsEqual Sh_Xml_Unescape("&lt;&gt;&quot;&apos;"), "<>""'"
        .IsEqual Sh_Xml_Unescape("&amp;lt;"), "&lt;"
        .IsEqual Sh_Xml_Unescape("AT&T &bogus; &#;"), "AT&T &bogus; &#;"
        .IsEqual Sh_Xml_Unescape("&#x1D465;"), ChrW(&HD835) & ChrW(&HDC65)
    End With

    With Suite.Test("percent escapes become their characters, read as UTF-8")
        .IsEqual Sh_Url_Unescape("EQ%201%C3%A9.jpg"), "EQ 1" & ChrW(233) & ".jpg"
        .IsEqual Sh_Url_Unescape("100% and %zz"), "100% and %zz"
        .IsEqual Sh_Url_Unescape("a%E9b"), "a" & ChrW(233) & "b"
        .IsEqual Sh_Url_Unescape("plain.jpg"), "plain.jpg"
    End With

    With Suite.Test("the picture path reads entities and escapes")
        .IsEqual Sh_MathML_Picture_Path("C:\Books\Imagine IM  A1 7&8\", "./images/7&amp;8/EQ%201.jpg"), _
            "C:\Books\Imagine IM  A1 7&8\images\7&8\EQ 1.jpg"
    End With

    ' --- The braille MathML keeps every character.
    With Suite.Test("braille MathML: line breaks and tabs become spaces, nothing else changes")
        Dim mathX As String
        mathX = ChrW(&HD835) & ChrW(&HDC65)            ' U+1D465, mathematical italic x
        .IsEqual Sh_MathML_One_Line("<math>" & vbCrLf & vbTab & "<mi>" & mathX & "</mi>" & ChrW(&HE000) & _
                                    ChrW(&H8000) & vbLf & "</math>"), _
            "<math>  <mi>" & mathX & "</mi>" & ChrW(&HE000) & ChrW(&H8000) & " </math>"
    End With

    ' --- The braille cleanups' park.
    With Suite.Test("the park mark: fixed width, brackets no pass names")
        .IsEqual Dx_Math_Park_Mark(12), ChrW(&H27E6) & "VTMATH000012" & ChrW(&H27E7)
        .IsEqual Len(Dx_Math_Park_Mark(1)), Len(Dx_Math_Park_Mark(999999))
    End With

    With Suite.Test("a mark is read back to its number, and nothing else is")
        .IsEqual Dx_Math_Mark_Number(Dx_Math_Park_Mark(12)), 12
        .IsEqual Dx_Math_Mark_Number(Dx_Math_Park_Mark(999999)), 999999
        .IsEqual Dx_Math_Mark_Number(ChrW(&H27E6) & "VTMATH00001x" & ChrW(&H27E7)), 0
        .IsEqual Dx_Math_Mark_Number(ChrW(&H27E6) & "VTMATH000012"), 0
        .IsEqual Dx_Math_Mark_Number("VTEQ000012"), 0
    End With

    With Suite.Test("one MathML element is measured whole, and only a real one")
        .IsEqual Dx_Math_Element_Length("<math alttext=""a > b""><mi>x</mi></math> and more"), 39
        .IsEqual Dx_Math_Element_Length("<math/> x"), 7
        .IsEqual Dx_Math_Element_Length("<math xmlns=""http://www.w3.org/1998/Math/MathML"" /> rest"), 51
        .IsEqual Dx_Math_Element_Length("<math><mi>x</mi></math >"), 24
        .IsEqual Dx_Math_Element_Length("<math>" & vbCr & "<mi>x</mi>" & vbCr & "</math>"), 25
        .IsEqual Dx_Math_Element_Length("<mathematics>x</mathematics>"), 0
        .IsEqual Dx_Math_Element_Length("<math><mi>x</mi>"), 0
        .IsEqual Dx_Math_Element_Length("<math><mi>x</mi></mathx>"), 0
    End With

    With Suite.Test("a book's MathML parks as marks and comes back exactly")
        Dim book As String, parked As String
        Dim kept As New Collection
        book = "Solve <math xmlns=""http://www.w3.org/1998/Math/MathML"" id=""m1"" alttext='x > 2'>" & _
               "<mi>x</mi><mo>" & ChrW(183) & "</mo><mn>2</mn></math>"" for x." & vbCr & _
               "Then <math id=""m2""/> and <mathematics> and <math><mi>y</mi></math>."
        parked = Park_Model(book, kept)
        .IsEqual kept.count, 3
        .IsEqual parked, "Solve " & Dx_Math_Park_Mark(1) & """ for x." & vbCr & "Then " & Dx_Math_Park_Mark(2) & _
                         " and <mathematics> and " & Dx_Math_Park_Mark(3) & "."
        .IsEqual Restore_Model(parked, kept), book
    End With

    ' --- Where the placeholders are, read once from the book's text (the speed fix).
    With Suite.Test("placeholder places are read from the text, counted from 0")
        f = Sh_MathML_Placeholder_Positions("ab VTEQ000001 and VTEQ000002.", 2)
        .IsEqual LBound(f), 1
        .IsEqual UBound(f), 2
        .IsEqual f(1), 3
        .IsEqual f(2), 18
        f = Sh_MathML_Placeholder_Positions("VTEQ000001", 1)
        .IsEqual f(1), 0
    End With

    With Suite.Test("each table cell end (CR and BEL) is one place in the book, not two")
        f = Sh_MathML_Placeholder_Positions("x" & Chr$(13) & Chr$(7) & "VTEQ000001", 1)
        .IsEqual f(1), 2
        ' Two cells and a row end before the second one; a BEL AFTER a placeholder moves nothing.
        f = Sh_MathML_Placeholder_Positions("VTEQ000001" & Chr$(13) & Chr$(7) & "y" & Chr$(13) & Chr$(7) & _
                                            Chr$(13) & Chr$(7) & "VTEQ000002" & Chr$(13) & Chr$(7), 2)
        .IsEqual f(1), 0
        .IsEqual f(2), 14
    End With

    With Suite.Test("a placeholder that is missing, or out of order, is -1")
        f = Sh_MathML_Placeholder_Positions("VTEQ000002 VTEQ000001", 2)
        .IsEqual f(1), 11
        .IsEqual f(2), -1
        f = Sh_MathML_Placeholder_Positions("VTEQ000001 VTEQ000003", 3)
        .IsEqual f(1), 0
        .IsEqual f(2), -1
        .IsEqual f(3), 11
        f = Sh_MathML_Placeholder_Positions("", 1)
        .IsEqual f(1), -1
    End With

    With Suite.Test("no placeholders asked for, an empty list")
        f = Sh_MathML_Placeholder_Positions("VTEQ000001", 0)
        .IsEqual UBound(f), -1
    End With

    Set TestMathML_Suite = Suite
End Function

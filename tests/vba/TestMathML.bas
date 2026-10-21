Attribute VB_Name = "TestMathML"
'@uses Sh_MathML_Prefix
'@uses Sh_MathML_Placeholder
'@uses Sh_MathML_Fragments
'@uses Sh_MathML_To_Placeholders
'@uses Sh_MathML_Altimg
'@uses Sh_MathML_Picture_Path
'@uses Sh_MathML_With_Namespace
'@uses Sh_Strip_Xml_Declaration
'@uses Sh_OMML_Package
'@uses Sh_Xml_Start_Tag_Length
'@uses Sh_Xml_Attribute
'@uses Sh_Xml_Unescape
'@uses Sh_Url_Unescape
'@uses Sh_MathML_Placeholder_Positions
'@uses Sh_MathML_Wrap_Fences
'@uses Sh_MathML_Wrap_Fences_In
'
' The string helpers behind the DAISY/NIMAS converter's equations (issue #37): finding a book's
' MathML and its prefix, parking each equation as a placeholder, the picture path from altimg,
' and the namespace put back for MSXML. Large print and braille both get Word equations (OMML)
' or, in large print only, the pictures. Nothing here touches a document. The parts that do -
' Find, InsertXML, AddPicture - are for the headless conversion and for real Word.
'
' From the review of 10/18/2026: start tags read with their quoted values (a ">" in alttext), and
' the altimg unescaped as XML and as a URL.
'
' Version: 1.4  Date: 10/20/2026 - brackets MML2OMML.XSL used to split, (-5) into "(|5)" (issue #38):
'                                the MathML is wrapped, and Word's own stylesheet is run on it
' Version: 1.3  Date: 10/20/2026 - the braille MathML text tests and the braille cleanup park
'                                tests are gone with the code they tested: braille gets Word
'                                equations (OMML) now (Jerry)
' Version: 1.2  Date: 10/18/2026 - where the placeholders are, read from the text (the speed fix)
' Version: 1.1  Date: 10/18/2026 - the review's examples
' Version: 1.0  Date: 10/18/2026
'
Option Explicit

Private Function NIMAS_Sample() As String
    ' The shape of the Imagine IM A1 7&8 book: the prefix declared once on the root, inline math.
    NIMAS_Sample = "<dtbook xmlns=""http://www.daisy.org/z3986/2005/dtbook/"" " & _
        "xmlns:m=""http://www.w3.org/1998/Math/MathML"" version=""2005-3"">" & _
        "<p>Solve <m:math id=""m1"" display=""inline"" altimg=""./images/U07C01/EQp010-001.jpg"" alttext="""">" & _
        "<m:mi>x</m:mi><m:mo>+</m:mo><m:mn>2</m:mn></m:math> for x.</p>" & _
        "<p>Then <m:math id=""m2"" display=""inline"" altimg=""./images/U07C01/EQp010-002.jpg"" alttext="""">" & vbCrLf & _
        "  <m:msup><m:mi>y</m:mi><m:mn>2</m:mn></m:msup>" & vbCrLf & "</m:math>.</p></dtbook>"
End Function

Private Function Eq(ByVal body As String) As String
    ' One equation as Sh_MathML_Put_Equation receives it: the namespace back on its own tag.
    Eq = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML"">" & body & "</m:math>"
End Function

Private Function Mo(ByVal ch As String) As String
    Mo = "<m:mo>" & ch & "</m:mo>"
End Function

Private Function Minus() As String
    Minus = ChrW(&H2212)
End Function

Private Function Omml_Parts(ByVal mathml As String) As String
    ' Word's own MML2OMML.XSL run on mathml, as Sh_MathML_Put_Equation runs it, and the brackets
    ' it made, each as its number of parts: "1" for (-5) as one part, "2" for (-5) split in
    ' two, "1 2" for two brackets. "ERROR" when the stylesheet or the MathML cannot be read.
    Dim src As Object, xsl As Object, out As Object, d As Object
    Dim omml As String, r As String
    On Error GoTo Failed
    Set xsl = CreateObject("MSXML2.DOMDocument.6.0")
    xsl.async = False
    If Not xsl.Load(Application.Path & "\MML2OMML.XSL") Then GoTo Failed
    Set src = CreateObject("MSXML2.DOMDocument.6.0")
    src.async = False
    If Not src.loadXML(mathml) Then GoTo Failed
    omml = src.transformNode(xsl)
    Set out = CreateObject("MSXML2.DOMDocument.6.0")
    out.async = False
    If Not out.loadXML(omml) Then GoTo Failed
    out.setProperty "SelectionNamespaces", "xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'"
    For Each d In out.selectNodes("//m:d")
        r = r & " " & CStr(d.selectNodes("m:e").length)
    Next d
    Omml_Parts = Trim$(r)
    Exit Function
Failed:
    Omml_Parts = "ERROR"
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

    ' --- Issue #38: brackets Word's MML2OMML.XSL used to split at the one operator inside.
    Dim m5 As String, pb As String, wrapped As String
    m5 = Eq("<m:mrow><m:mo stretchy=""false"">(</m:mo>" & Mo(Minus()) & "<m:mn>5</m:mn><m:mo stretchy=""false"">)</m:mo></m:mrow>")
    pb = Eq("<m:mrow>" & Mo("(") & "<m:mi>P</m:mi>" & Mo("+") & "<m:mi>b</m:mi>" & Mo(")") & "</m:mrow>")

    With Suite.Test("(-5): what is inside the brackets becomes one mrow")
        .IsEqual Sh_MathML_Wrap_Fences(m5), _
            Eq("<m:mrow><m:mo stretchy=""false"">(</m:mo><m:mrow>" & Mo(Minus()) & "<m:mn>5</m:mn></m:mrow>" & _
               "<m:mo stretchy=""false"">)</m:mo></m:mrow>")
    End With

    With Suite.Test("(P+b), [x+y], |x+y|, {x+y}: the same")
        .IsEqual Sh_MathML_Wrap_Fences(pb), _
            Eq("<m:mrow>" & Mo("(") & "<m:mrow><m:mi>P</m:mi>" & Mo("+") & "<m:mi>b</m:mi></m:mrow>" & Mo(")") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(Eq("<m:mrow>" & Mo("[") & "<m:mi>x</m:mi>" & Mo("+") & "<m:mi>y</m:mi>" & Mo("]") & "</m:mrow>")), _
            Eq("<m:mrow>" & Mo("[") & "<m:mrow><m:mi>x</m:mi>" & Mo("+") & "<m:mi>y</m:mi></m:mrow>" & Mo("]") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(Eq("<m:mrow>" & Mo("|") & "<m:mi>x</m:mi>" & Mo("+") & "<m:mi>y</m:mi>" & Mo("|") & "</m:mrow>")), _
            Eq("<m:mrow>" & Mo("|") & "<m:mrow><m:mi>x</m:mi>" & Mo("+") & "<m:mi>y</m:mi></m:mrow>" & Mo("|") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(Eq("<m:mrow>" & Mo("{") & "<m:mi>x</m:mi>" & Mo("+") & "<m:mi>y</m:mi>" & Mo("}") & "</m:mrow>")), _
            Eq("<m:mrow>" & Mo("{") & "<m:mrow><m:mi>x</m:mi>" & Mo("+") & "<m:mi>y</m:mi></m:mrow>" & Mo("}") & "</m:mrow>")
    End With

    With Suite.Test("a list keeps its commas: (3, 4) and {1, 2} come back exactly as they came")
        x = Eq("<m:mrow>" & Mo("(") & "<m:mn>3</m:mn>" & Mo(",") & "<m:mn>4</m:mn>" & Mo(")") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(x), x
        x = Eq("<m:mrow>" & Mo("{") & "<m:mn>1</m:mn>" & Mo(",") & "<m:mn>2</m:mn>" & Mo("}") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(x), x
    End With

    With Suite.Test("nested brackets: ((-5)+1) - the inner and the outer are each one part")
        .IsEqual Sh_MathML_Wrap_Fences(Eq("<m:mrow>" & Mo("(") & "<m:mrow>" & Mo("(") & Mo(Minus()) & "<m:mn>5</m:mn>" & Mo(")") & _
                                          "</m:mrow>" & Mo("+") & "<m:mn>1</m:mn>" & Mo(")") & "</m:mrow>")), _
            Eq("<m:mrow>" & Mo("(") & "<m:mrow><m:mrow>" & Mo("(") & "<m:mrow>" & Mo(Minus()) & "<m:mn>5</m:mn></m:mrow>" & Mo(")") & _
               "</m:mrow>" & Mo("+") & "<m:mn>1</m:mn></m:mrow>" & Mo(")") & "</m:mrow>")
    End With

    With Suite.Test("left alone: brackets that do not match, mixed operators, brackets mid-row")
        x = Eq("<m:mrow>" & Mo("(") & Mo(Minus()) & "<m:mn>5</m:mn>" & Mo("]") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(x), x
        x = Eq("<m:mrow>" & Mo("(") & "<m:mi>x</m:mi>" & Mo("+") & "<m:mi>y</m:mi>" & Mo(Minus()) & "<m:mi>z</m:mi>" & Mo(")") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(x), x
        x = Eq("<m:mrow><m:mn>2</m:mn>" & Mo("(") & Mo(Minus()) & "<m:mn>5</m:mn>" & Mo(")") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(x), x
        x = Eq("<m:mrow>" & Mo("(") & "<m:mi>x</m:mi>" & Mo(")") & "</m:mrow>")
        .IsEqual Sh_MathML_Wrap_Fences(x), x
    End With

    With Suite.Test("MathML with no prefix, and MathML that cannot be read")
        .IsEqual Sh_MathML_Wrap_Fences("<math xmlns=""http://www.w3.org/1998/Math/MathML""><mrow><mo>(</mo><mo>" & Minus() & _
                                       "</mo><mn>5</mn><mo>)</mo></mrow></math>"), _
            "<math xmlns=""http://www.w3.org/1998/Math/MathML""><mrow><mo>(</mo><mrow><mo>" & Minus() & _
            "</mo><mn>5</mn></mrow><mo>)</mo></mrow></math>"
        .IsEqual Sh_MathML_Wrap_Fences("<m:math><m:mi>x</m:mi></m:math>"), "<m:math><m:mi>x</m:mi></m:math>"
        .IsEqual Sh_MathML_Wrap_Fences(""), ""
    End With

    With Suite.Test("Word's own MML2OMML.XSL: split before, one part after; a list still two")
        .IsEqual Omml_Parts(m5), "2", "(-5) as it is in the book - got {0}, expected {1}"
        .IsEqual Omml_Parts(Sh_MathML_Wrap_Fences(m5)), "1", "(-5) wrapped - got {0}, expected {1}"
        .IsEqual Omml_Parts(pb), "2", "(P+b) as it is in the book - got {0}, expected {1}"
        .IsEqual Omml_Parts(Sh_MathML_Wrap_Fences(pb)), "1", "(P+b) wrapped - got {0}, expected {1}"
        x = Eq("<m:mrow>" & Mo("(") & "<m:mn>3</m:mn>" & Mo(",") & "<m:mn>4</m:mn>" & Mo(")") & "</m:mrow>")
        .IsEqual Omml_Parts(Sh_MathML_Wrap_Fences(x)), "2", "(3, 4) is still a list of two - got {0}, expected {1}"
        wrapped = Sh_MathML_Wrap_Fences(Eq("<m:mrow>" & Mo("(") & "<m:mrow>" & Mo("(") & Mo(Minus()) & "<m:mn>5</m:mn>" & Mo(")") & _
                                           "</m:mrow>" & Mo("+") & "<m:mn>1</m:mn>" & Mo(")") & "</m:mrow>"))
        .IsEqual Omml_Parts(wrapped), "1 1", "((-5)+1) - got {0}, expected {1}"
    End With

    Set TestMathML_Suite = Suite
End Function

Attribute VB_Name = "TestMTEF"
'@uses Sh_MTEF_From_MathML
'@uses Sh_MTEF_Not_Handled
'@uses Sh_MTEF_Ole_File
'@uses Sh_Base64
'@uses Sh_MTEF_Preamble_Hex
'@uses Sh_MTEF_Text_Bytes
'@uses Sh_MTEF_Hex_Text
'@uses Sh_MathML_With_Namespace
'@uses Sh_MathType_Package
'@uses Sh_Picture_Size_Pt
'@uses Sh_Points_Text
'@uses Sh_MathType_Is_Installed
'
' MathType equations made from MathML (Steps 1 and 2 of the MathType plan, 10/20/2026): the
' MTEF that Sh_MTEF_From_MathML writes, the OLE compound file Sh_MTEF_Ole_File wraps it in, and
' Sh_Base64. Nothing here touches a document or MathType.
'
' Every expected value below is MathType's own output, not a guess. The Section D fixtures are
' the MTEF of "Section D with MathType.docx" on VistaBuild with its COLOR records taken out; the
' others are what MathType's Convert Equations made, headlessly on VistaBuild, from the same
' MathML put in as Word equations. Only the records after the 193-byte preamble are compared -
' the preamble is checked once.
'
' Version: 1.3  Date: 11/5/2026 - what Basic Math sent to Word equations (Jerry): no-break spaces,
'                                mspace, normal mi text, % and $, infinity and aleph, the Euclid
'                                fonts, borrowing marks and long division (issue #50). Expected
'                                bytes from MathType: the book after its Convert Equations, and its
'                                own MathML import where that route loses the thing or never meets it
' Version: 1.2  Date: 10/20/2026 - Step 3's pure pieces: the Flat-OPC package a MathType equation
'                                goes in with, and the picture sizes read from a file's bytes
' Version: 1.1  Date: 10/20/2026 - letters in a one-column table stay bold italic (Jerry,
'                                10/20/2026), not MathType's plain italic
' Version: 1.0  Date: 10/20/2026
'
Option Explicit

Private Const MML As String = "xmlns:m=""http://www.w3.org/1998/Math/MathML"""

Private Function HexOf(ByVal b As Variant) As String
    ' A Byte array as upper case hex, two digits a byte; "" for anything else.
    Dim i As Long, s As String
    If Not IsArray(b) Then Exit Function
    s = String$(2 * (UBound(b) - LBound(b) + 1), "0")
    For i = LBound(b) To UBound(b)
        Mid$(s, 2 * (i - LBound(b)) + 1, 2) = Right$("0" & Hex$(b(i)), 2)
    Next i
    HexOf = s
End Function

Private Function Body(ByVal b As Variant) As String
    ' The MTEF after MathType's 193-byte preamble, as hex - "EMPTY" when there is none, and
    ' "PREAMBLE?" when the first 193 bytes are not the preamble.
    Dim h As String
    If IsEmpty(b) Then
        Body = "EMPTY"
        Exit Function
    End If
    h = HexOf(b)
    If Left$(h, 386) <> Sh_MTEF_Preamble_Hex() Then
        Body = "PREAMBLE?"
        Exit Function
    End If
    Body = Mid$(h, 387)
End Function

Private Function FromBase64(ByVal s As String) As Variant
    ' Base64 back to a Byte array, through MSXML, to prove Sh_Base64 round trips.
    Dim doc As Object, el As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = doc.createElement("b")
    el.DataType = "bin.base64"
    el.Text = s
    FromBase64 = el.nodeTypedValue
End Function

Private Sub BasicMathTests(ByVal Suite As TestSuite)
    ' 11/5/2026: in subs of their own, so TestMTEF_Suite stays under VBA's size limit for one
    ' procedure ("Procedure too large" is a compile error, and so a hang headlessly).
    BasicMathTests1 Suite
    BasicMathTests2 Suite
End Sub

Private Sub BasicMathTests1(ByVal Suite As TestSuite)
    ' No-break spaces, mspace, normal mi text, % $, infinity, aleph, U+2019.
    Dim x As String

    ' --- 11/5/2026 (Jerry): what Basic Math Pre-Algebra for Dummies sent to Word equations. The
    ' book equations below are as the book has them, ids left out; their bytes are MathType's own
    ' after Convert Equations, COLOR records out. tools/compare_mtef.py makes the same comparison
    ' over the whole book.
    With Suite.Test("a no-break space is TEXT in an mi and FUNCTION in an mo (Basic Math)")
        x = "<m:math " & MML & ">" & "<m:mrow><m:mo>&#215;</m:mo><m:mi mathvariant=""normal"">&#160;</m:mi><m:mn>0.01</m:mn>" & _
            "</m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020486D700B4020081A00002008830000200822E00020088300002008831000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mrow><m:mn>5</m:mn><m:mo>&#160;</m:mo><m:mo>&#215;</m:mo><m:mn>2</m:mn></m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0200883500020082A000020486D700B402008832000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("an mspace is a MathType space: thick and 0.25em are EF03, 1em EF05, 0.5em two")
        ' MathType's own MathML import, measured on VistaBuild 11/5/2026 (its Convert Equations
        ' route drops every mspace, so there is nothing to copy from that).
        x = "<m:math " & MML & ">" & "<m:mn>1</m:mn><m:mspace width=""thickmathspace""/><m:mn>2</m:mn>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088310002009803EF02008832000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mn>1</m:mn><m:mspace width=""0.25em""/><m:mn>2</m:mn>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088310002009803EF02008832000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mn>1</m:mn><m:mspace width=""1em""/><m:mn>2</m:mn>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088310002009805EF02008832000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mn>1</m:mn><m:mspace width=""0.5em""/><m:mn>2</m:mn>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088310002009804EF02009803EF02008832000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("22.82 feet keeps its space (Jerry, 11/5/2026)")
        ' MathType's own MathML import; through Convert Equations it read 22.82feet.
        x = "<m:math " & MML & ">" & "<m:mn>22.82</m:mn><m:mspace width=""thickmathspace""/><m:mi mathvariant=""normal"">feet<" & _
            "/m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088320002008832000200822E000200883800020088320002009803EF020081660002008165000200" & _
            "81650002008174000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("a normal mi is typed as MathType does: letters TEXT, + and = SYMBOL, ? FUNCTION")
        x = "<m:math " & MML & ">" & "<m:mrow><m:mi mathvariant=""normal"">Bobo&#160;+&#160;Nunu&#160;=&#160;?</m:mi></m:mro" & _
            "w>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "02008142000200816F0002008162000200816F00020081A0000204862B002B020081A0000200814E0002" & _
            "008175000200816E000200817500020081A0000204863D003D020081A0000200823F000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("digits and commas in a normal mi, and an ellipsis in an mn (Basic Math)")
        x = "<m:math " & MML & ">" & "<m:mrow><m:mi mathvariant=""normal"">1,&#160;10,&#160;100,&#160;1,000,&#160;10,000,&#1" & _
            "60;100,000,&#160;1,000,000,&#160;</m:mi><m:mn>&#8230;</m:mn></m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "02008831000200822C00020081A000020088310002008830000200822C00020081A00002008831000200" & _
            "88300002008830000200822C00020081A00002008831000200822C000200883000020088300002008830" & _
            "000200822C00020081A000020088310002008830000200822C0002008830000200883000020088300002" & _
            "00822C00020081A0000200883100020088300002008830000200822C0002008830000200883000020088" & _
            "30000200822C00020081A00002008831000200822C000200883000020088300002008830000200822C00" & _
            "0200883000020088300002008830000200822C00020081A0000204862620BC0000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("an en dash in a normal mi is the minus")
        x = "<m:math " & MML & ">" & "<m:mi mathvariant=""normal"">&#8211;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "02048612222D0000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("em spaces in a normal mi are MathType's em space EF05")
        x = "<m:math " & MML & ">" & "<m:mrow><m:mtable><m:mtr><m:mtd><m:mrow><m:mn>34.8</m:mn><m:mi mathvariant=""normal"">" & _
            "&#8195;</m:mi><m:mi mathvariant=""normal"">&#8195;</m:mi><m:mi mathvariant=""normal"">&#" & _
            "8195;</m:mi><m:mi mathvariant=""normal"">&#8195;</m:mi><m:mi mathvariant=""normal"">&#81" & _
            "95;</m:mi><m:mi mathvariant=""normal"">&#8195;</m:mi><m:mi mathvariant=""normal"">&#8195" & _
            ";</m:mi><m:mi mathvariant=""normal"">&#8195;</m:mi></m:mrow></m:mtd><m:mtd><m:mrow><m:" & _
            "mn>34.80</m:mn><m:mi mathvariant=""normal"">&#8195;</m:mi><m:mi mathvariant=""normal"">&" & _
            "#8195;</m:mi><m:mi mathvariant=""normal"">&#8195;</m:mi><m:mi mathvariant=""normal"">&#8" & _
            "195;</m:mi><m:mi mathvariant=""normal"">&#8195;</m:mi><m:mi mathvariant=""normal"">&#819" & _
            "5;</m:mi><m:mi mathvariant=""normal"">&#8195;</m:mi><m:mi mathvariant=""normal"">&#8195;" & _
            "</m:mi></m:mrow></m:mtd><m:mtd><m:mrow><m:mn>34.8000</m:mn></m:mrow></m:mtd></m:mtr>" & _
            "</m:mtable></m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0500010101010300000100020088330002008834000200822E00020088380002009805EF02009805EF02" & _
            "009805EF02009805EF02009805EF02009805EF02009805EF02009805EF00010002008833000200883400" & _
            "0200822E000200883800020088300002009805EF02009805EF02009805EF02009805EF02009805EF0200" & _
            "9805EF02009805EF02009805EF000100020088330002008834000200822E000200883800020088300002" & _
            "00883000020088300000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("% and $ in an mi are FUNCTION")
        x = "<m:math " & MML & ">" & "<m:mrow><m:mn>4</m:mn><m:mi>&#37;</m:mi><m:mo>=</m:mo><m:mn>0.04</m:mn></m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088340002008225000204863D003D02008830000200822E00020088300002008834000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mrow><m:mi>&#36;</m:mi><m:mfrac><m:mn>5</m:mn><m:mn>2</m:mn></m:mfrac></m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020082240003000B000001000200883500000100020088320000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("infinity and aleph in an mi are on the Symbol font")
        x = "<m:math " & MML & ">" & "<m:mi>&#8734;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0204861E22A50000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#8501;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0204863521C00000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("U+2019 in an mo is FUNCTION")
        x = "<m:math " & MML & ">" & "<m:mrow><m:mi mathvariant=""normal"">This&#160;year</m:mi><m:mo>&#8217;</m:mo><m:mi ma" & _
            "thvariant=""normal"">s&#160;salary</m:mi><m:mo>=</m:mo><m:mi>&#36;</m:mi><m:mn>42</m:m" & _
            "n><m:mo>,</m:mo><m:mn>000</m:mn></m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0200815400020081680002008169000200817300020081A0000200817900020081650002008161000200" & _
            "81720002008219200200817300020081A000020081730002008161000200816C00020081610002008172" & _
            "0002008179000204863D003D0200822400020088340002008832000200822C0002008830000200883000" & _
            "02008830000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
End Sub

Private Sub BasicMathTests2(ByVal Suite As TestSuite)
    ' The Euclid fonts, borrowing marks, long division (issue #50), and what is still refused.
    Dim x As String

    With Suite.Test("asymptotic to is on Euclid Math One, N Z Q R C on Euclid Math Two")
        ' Each brings its encoding, font and style records with it, as MathType wrote them.
        x = "<m:math " & MML & ">" & "<m:mo>&#8781;</m:mo>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468310011064575636C6964204D617468204F6E650008050002047F4D22A90000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#8469;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468320011064575636C6964204D6174682054776F0008050002047F15214E0000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#8484;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468320011064575636C6964204D6174682054776F0008050002047F24215A0000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#8474;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468320011064575636C6964204D6174682054776F0008050002047F1A21510000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#8477;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468320011064575636C6964204D6174682054776F0008050002047F1D21520000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#8450;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468320011064575636C6964204D6174682054776F0008050002047F0221430000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("bold italic and the Euclid fonts are numbered in the order written")
        ' MathType's own MathML import, measured on VistaBuild 11/5/2026: no equation in Basic Math
        ' has both, so these come from MathType itself, not from that book.
        x = "<m:math " & MML & ">" & "<m:mi mathvariant=""bold-italic"">x</m:mi><m:mo>&#x2208;</m:mo><m:mi>&#x211D;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F78000204860822CE134575636C69644D617468320011064575636C6964204D6174682054" & _
            "776F0008050002047E1D21520000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#x2115;</m:mi><m:mo>,</m:mo><m:mi mathvariant=""bold-italic"">x</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468320011064575636C6964204D6174682054776F0008050002047F15214E0200" & _
            "822C0008010302007E78000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi mathvariant=""bold-italic"">x</m:mi><m:mo>&#x2208;</m:mo><m:mi>&#x2124;</m:mi><m" & _
            ":mo>,</m:mo><m:mi mathvariant=""bold-italic"">y</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F78000204860822CE134575636C69644D617468320011064575636C6964204D6174682054" & _
            "776F0008050002047E24215A0200822C0002007F79000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mi>&#x211D;</m:mi><m:mo>&#x224D;</m:mo><m:mi>&#x2102;</m:mi>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "134575636C69644D617468320011064575636C6964204D6174682054776F0008050002047F1D21521345" & _
            "75636C69644D617468310011074575636C6964204D617468204F6E650008060002047E4D22A902047F02" & _
            "21430000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("borrowing marks: scripts before the base are template 29 variation 1")
        ' The first as MathType's Convert Equations wrote the 7 in Basic Math (cut from its equation
        ' 136 there); the other two from MathType's own MathML import.
        x = "<m:math " & MML & ">" & "<m:mmultiscripts><m:mn>8</m:mn><m:mprescripts/><m:mn>1</m:mn></m:mmultiscripts>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03001D01000B0100020088310000010000000A02008838000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mmultiscripts><m:mn>8</m:mn><m:mprescripts/><m:none/><m:mn>1</m:mn></m:mmultiscri" & _
            "pts>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03001D01000B0100000100020088310000000A02008838000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:mmultiscripts><m:mn>8</m:mn><m:mprescripts/><m:mn>1</m:mn><m:mn>2</m:mn></m:mmult" & _
            "iscripts>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03001D01000B01000200883100000100020088320000000A02008838000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("long division is template 26, and keeps its dividend (issue #50)")
        ' MathType's own MathML import, measured on VistaBuild 11/5/2026. As a Word equation the
        ' dividend is lost (issue #50), so Convert Equations had nothing to copy.
        x = "<m:math " & MML & ">" & "<m:mrow><m:mn>5</m:mn><m:menclose notation=""longdiv""><m:mrow><m:mn>860</m:mn></m:mro" & _
            "w></m:menclose></m:mrow>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088350003001A00000100020088380002008836000200883000000101000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
        x = "<m:math " & MML & ">" & "<m:menclose notation=""longdiv""><m:mn>860</m:mn></m:menclose>" & "</m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03001A00000100020088380002008836000200883000000101000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With
    With Suite.Test("what is still refused: other mspace widths and mmultiscripts, aleph in an mo")
        x = "<m:math " & MML & "><m:mn>1</m:mn><m:mspace width=""3pt""/><m:mn>2</m:mn></m:math>"
        .IsOk IsEmpty(Sh_MTEF_From_MathML(x))
        .IsEqual Sh_MTEF_Not_Handled(x), "mspace width=""3pt"" is not handled"
        x = "<m:math " & MML & "><m:mmultiscripts><m:mi>x</m:mi><m:mn>1</m:mn><m:none/></m:mmultiscripts></m:math>"
        .IsEqual Sh_MTEF_Not_Handled(x), "mmultiscripts other than base, mprescripts and one or two scripts is not handled"
        x = "<m:math " & MML & "><m:mo>&#x2135;</m:mo></m:math>"
        .IsEqual Sh_MTEF_Not_Handled(x), "the character U+2135 in mo is not handled"
        x = "<m:math " & MML & "><m:mi mathvariant=""normal"">a&#x2009;b</m:mi></m:math>"
        .IsEqual Sh_MTEF_Not_Handled(x), "the character U+2009 in mi (normal) is not handled"
    End With
End Sub

Public Function TestMTEF_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "MathType equations (MTEF) from MathML"

    Dim x As String, b As Variant, ole As Variant, h As String, big() As Byte, i As Long

    ' --- The preamble: MathType's own 193 bytes, the same in every equation.
    With Suite.Test("the preamble is MathType's 193 bytes, and the number 1 follows it")
        .IsEqual Len(Sh_MTEF_Preamble_Hex()), 386
        .IsEqual Left$(Sh_MTEF_Preamble_Hex(), 22), "050100070C44534D543700"
        b = Sh_MTEF_From_MathML("<m:math " & MML & "><m:mn>1</m:mn></m:math>")
        .IsEqual UBound(b) - LBound(b) + 1, 200
        .IsEqual Body(b), "02008831000000"
    End With

    With Suite.Test("a square root, plus-minus, a TEXT run of underscores (equation 1)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1477"" display=""inline"" altimg=""./images/U08C04/EQp346-001.jpg"" altt" & _
            "ext=""""><m:mrow><m:mn>1</m:mn><m:mo>&#x00B1;</m:mo><m:msqrt><m:mrow><m:mn>49</m:m" & _
            "n></m:mrow></m:msqrt><m:mtext>&#x005F;&#x005F;&#x005F;&#x005F;&#x005F;&#x005F;&#" & _
            "x005F;&#x005F;&#x005F;</m:mtext></m:mrow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0200883100020486B100B103000A0000010002008834000200883900000B0101000A0200815F0002" & _
            "00815F000200815F000200815F000200815F000200815F000200815F000200815F000200815F0000" & _
            "00"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a fraction (equation 2)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1478"" display=""inline"" altimg=""./images/U08C04/EQp346-002.jpg"" altt" & _
            "ext=""""><m:mrow><m:mfrac><m:mrow><m:mn>8</m:mn><m:mo>&#x00B1;</m:mo><m:mn>2</m:mn" & _
            "></m:mrow><m:mn>5</m:mn></m:mfrac><m:mtext>&#x005F;&#x005F;&#x005F;&#x005F;&#x00" & _
            "5F;&#x005F;&#x005F;&#x005F;&#x005F;</m:mtext></m:mrow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03000B000001000200883800020486B100B10200883200000100020088350000000200815F000200" & _
            "815F000200815F000200815F000200815F000200815F000200815F000200815F000200815F000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a superscript, bold italic letters, the minus, a decimal point (equation 5)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1481"" display=""inline"" altimg=""./images/U08C04/EQp347-001.jpg"" altt" & _
            "ext=""""><m:mrow><m:msup><m:mi mathvariant=""bold-italic"">x</m:mi><m:mn>2</m:mn></m" & _
            ":msup><m:mo>&#x2212;</m:mo><m:mn>2</m:mn><m:mi mathvariant=""bold-italic"">x</m:mi" & _
            "><m:mo>&#x2212;</m:mo><m:mn>1.25</m:mn><m:mo>&#x003D;</m:mo><m:mn>0</m:mn></m:mr" & _
            "ow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F780003001C00000B01010100020088320000000A02048612222D020088320002007F" & _
            "780002048612222D02008831000200822E00020088320002008835000204863D003D020088300000" & _
            "00"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("not equal on the Symbol font (equation 37)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1513"" display=""inline"" altimg=""./images/U08C04/EQp352-003.jpg"" altt" & _
            "ext=""""><m:mrow><m:mi mathvariant=""bold-italic"">a</m:mi><m:mo>&#x2260;</m:mo><m:m" & _
            "n>0</m:mn></m:mrow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F61000204866022B902008830000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("brackets in the middle of a row, and a whole mrow of bars (equation 78)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1554"" display=""inline"" altimg=""./images/U08C04/EQp355-007.jpg"" altt" & _
            "ext=""""><m:mrow><m:mi mathvariant=""bold-italic"">v</m:mi><m:mrow><m:mo>(</m:mo><m:" & _
            "mi mathvariant=""bold-italic"">x</m:mi><m:mo>)</m:mo></m:mrow><m:mo>&#x003D;</m:mo" & _
            "><m:mrow><m:mo>&#x007C;</m:mo> <m:mrow><m:mi mathvariant=""bold-italic"">x</m:mi><" & _
            "m:mo>&#x002B;</m:mo><m:mn>6</m:mn></m:mrow> <m:mo>&#x007C;</m:mo></m:mrow></m:mr" & _
            "ow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F76000300010300010002007F78000002009628000200962900000204863D003D0300" & _
            "040300010002007F78000204862B002B02008836000002009607EC02009608EC000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a fraction inside brackets, raised to a power (equation 114)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1590"" display=""inline"" altimg=""./images/U08C04/EQp365-001.jpg"" altt" & _
            "ext=""""><m:mrow><m:mi mathvariant=""bold-italic"">f</m:mi><m:mo stretchy=""false"">(<" & _
            "/m:mo><m:mi mathvariant=""bold-italic"">w</m:mi><m:mo stretchy=""false"">)</m:mo><m:" & _
            "mo>&#x003D;</m:mo><m:mn>100</m:mn><m:mo>,</m:mo><m:mn>000</m:mn><m:mo>&#x22C5;</" & _
            "m:mo><m:msup><m:mrow><m:mrow><m:mo>(</m:mo><m:mrow><m:mfrac><m:mn>9</m:mn><m:mro" & _
            "w><m:mn>10</m:mn></m:mrow></m:mfrac></m:mrow><m:mo>)</m:mo></m:mrow></m:mrow><m:" & _
            "mi mathvariant=""bold-italic"">w</m:mi></m:msup></m:mrow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F66000300010300010002007F77000002009628000200962900000204863D003D0200" & _
            "883100020088300002008830000200822C00020088300002008830000200883000020486C522D703" & _
            "00010300010003000B00000100020088390000010002008831000200883000000000020096280002" & _
            "009629000003001C00000B0101010002007F770000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("thin and em spaces, text, boxes in a fraction (equation 292)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1768"" display=""inline"" altimg=""./images/U08C04/EQp394-008.jpg"" altt" & _
            "ext=""Equation 1, if the root of 2 equals a over b, then 2 equals blank over blan" & _
            "k.""><m:mrow><m:mtext>If</m:mtext><m:mtext>&#x2009;</m:mtext><m:msqrt><m:mn>2</m:" & _
            "mn></m:msqrt><m:mo>&#x003D;</m:mo><m:mfrac><m:mi mathvariant=""bold-italic"">a</m:" & _
            "mi><m:mi mathvariant=""bold-italic"">b</m:mi></m:mfrac><m:mo>,</m:mo><m:mtext>&#x2" & _
            "009;</m:mtext><m:mtext>then</m:mtext><m:mtext>&#x2009;</m:mtext><m:mn>2</m:mn><m" & _
            ":mo>&#x003D;</m:mo><m:mfrac><m:mrow><m:menclose notation=""box""><m:mrow><m:mtext>" & _
            "&#x2003;</m:mtext><m:mtext>&#x2003;</m:mtext></m:mrow></m:menclose></m:mrow><m:m" & _
            "row><m:menclose notation=""box""><m:mrow><m:mtext>&#x2003;</m:mtext><m:mtext>&#x20" & _
            "03;</m:mtext></m:mrow></m:menclose></m:mrow></m:mfrac></m:mrow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0200814900020081660002009802EF03000A000001000200883200000B0101000A0204863D003D03" & _
            "000B0000010008010302007F610000010002007F620000000200822C0002009802EF020081740002" & _
            "0081680002008165000200816E0002009802EF02008832000204863D003D03000B00000100030025" & _
            "1E00010002009805EF02009805EF00000001000300251E00010002009805EF02009805EF00000000" & _
            "0000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("an em space and a bracketed list with commas (equation 378)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1854"" display=""inline"" altimg=""./images/U08C04/EQp403-002.jpg"" altt" & _
            "ext=""""><m:mrow><m:mn>3</m:mn><m:mo>&#x002B;</m:mo><m:mn>4</m:mn><m:msqrt><m:mn>5" & _
            "</m:mn></m:msqrt><m:mtext>&#x2003;</m:mtext><m:mo stretchy=""false"">(</m:mo><m:mi" & _
            " mathvariant=""bold-italic"">a</m:mi><m:mo>&#x003D;</m:mo><m:mn>3</m:mn><m:mo>,</m" & _
            ":mo><m:mi mathvariant=""bold-italic"">b</m:mi><m:mo>&#x003D;</m:mo><m:mn>4</m:mn><" & _
            "m:mo stretchy=""false"">)</m:mo></m:mrow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "02008833000204862B002B020088340003000A000001000200883500000B0101000A02009805EF03" & _
            "00010300010008010302007F61000204863D003D02008833000200822C0002007F62000204863D00" & _
            "3D02008834000002009628000200962900000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("underscores in an mo are FUNCTION characters (equation 454)")
        x = Sh_MathML_With_Namespace("<m:math id=""m1930"" display=""inline"" altimg=""./images/U08C04/EQp410-005.jpg"" altt" & _
            "ext=""""><m:mrow><m:mn>10</m:mn><m:mo>&#x00B1;</m:mo><m:mn>0.3</m:mn><m:mo>&#x005F" & _
            ";&#x005F;&#x005F;&#x005F;&#x005F;&#x005F;&#x005F;&#x005F;&#x005F;&#x005F;&#x005F" & _
            ";&#x005F;&#x005F;&#x005F;</m:mo></m:mrow></m:math>", "m:")
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "02008831000200883000020486B100B102008830000200822E0002008833000200825F000200825F" & _
            "000200825F000200825F000200825F000200825F000200825F000200825F000200825F000200825F" & _
            "000200825F000200825F000200825F000200825F000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a subscript: template 27, the superscript slot empty")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:msub><m:mi mathvariant=""" & _
            "bold-italic"">x</m:mi><m:mn>2</m:mn></m:msub></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F780003001B00000B01000200883200000101000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a subscript and superscript: template 29")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:msubsup><m:mi mathvarian" & _
            "t=""bold-italic"">x</m:mi><m:mn>2</m:mn><m:mn>3</m:mn></m:msubsup></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F780003001D00000B01000200883200000100020088330000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("an nth root: the index one size down")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mroot><m:mn>8</m:mn><m:m" & _
            "n>3</m:mn></m:mroot></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03000A010001000200883800000B0100020088330000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a superscript on a superscript is SUB2")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:msup><m:mi mathvariant=""" & _
            "bold-italic"">x</m:mi><m:msup><m:mn>2</m:mn><m:mn>3</m:mn></m:msup></m:msup></m:m" & _
            "ath>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F780003001C00000B01010100020088320003001C00000C0101010002008833000000" & _
            "00000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("FULL comes back after a script, before the next character")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:msub><m:mi mathvariant=""" & _
            "bold-italic"">x</m:mi><m:mn>2</m:mn></m:msub><m:mo>+</m:mo><m:mi mathvariant=""bol" & _
            "d-italic"">y</m:mi></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F780003001B00000B01000200883200000101000A0204862B002B02007F79000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a bar over one letter is an embellishment")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mover><m:mi mathvariant=" & _
            """bold-italic"">x</m:mi><m:mo>&#x00AF;</m:mo></m:mover></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302017F7800060011000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a bar over two letters is the overbar template")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mover><m:mrow><m:mi math" & _
            "variant=""bold-italic"">A</m:mi><m:mi mathvariant=""bold-italic"">B</m:mi></m:mrow><" & _
            "m:mo>&#x00AF;</m:mo></m:mover></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03000D0000010008010302007F410002007F420000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a prime on a digit is an embellishment")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:msup><m:mn>5</m:mn><m:mo" & _
            ">&#x2032;</m:mo></m:msup></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0201883500060005000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("mfenced with square brackets")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mfenced open=""["" close=""" & _
            "]""><m:mi mathvariant=""bold-italic"">x</m:mi></m:mfenced></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0300030300010008010302007F7800000200965B000200965D00000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("braces in the middle of a row")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mi mathvariant=""bold-ita" & _
            "lic"">a</m:mi><m:mo>{</m:mo><m:mi mathvariant=""bold-italic"">x</m:mi><m:mo>}</m:mo" & _
            "><m:mi mathvariant=""bold-italic"">y</m:mi></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F61000300020300010002007F7800000200967B000200967D000002007F79000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("Greek on the Symbol font")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mi>&#x03B8;</m:mi></m:ma" & _
            "th>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020484B803710000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a plain mi of more than one letter is TEXT")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mi>sin</m:mi><m:mi>x</m:" & _
            "mi></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020081730002008169000200816E0002008378000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a bevelled fraction has a slash")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mfrac bevelled=""true""><m" & _
            ":mn>1</m:mn><m:mn>2</m:mn></m:mfrac></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03000B020001000200883100000100020088320000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a fraction with no line is a two-row matrix")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mfrac linethickness=""0"">" & _
            "<m:mn>5</m:mn><m:mn>2</m:mn></m:mfrac></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "05000101010201000001000200883500000100020088320000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a two-by-two table")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mtable><m:mtr><m:mtd><m:" & _
            "mn>1</m:mn></m:mtd><m:mtd><m:mn>2</m:mn></m:mtd></m:mtr><m:mtr><m:mtd><m:mn>3</m" & _
            ":mn></m:mtd><m:mtd><m:mn>4</m:mn></m:mtd></m:mtr></m:mtable></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "05000101010202000001000200883100000100020088320000010002008833000001000200883400" & _
            "00000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("letters directly in a one-column table cell are VARIABLE")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mtable><m:mtr><m:mtd><m:" & _
            "mi>x</m:mi><m:mo>=</m:mo><m:mn>1</m:mn></m:mtd></m:mtr><m:mtr><m:mtd><m:mi>y</m:" & _
            "mi></m:mtd></m:mtr></m:mtable></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "050001010102010000010002008378000204863D003D0200883100000100020083790000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("a short table row is filled out with an empty cell")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mtable><m:mtr><m:mtd><m:" & _
            "mn>1</m:mn></m:mtd><m:mtd><m:mn>2</m:mn></m:mtd></m:mtr><m:mtr><m:mtd><m:mn>3</m" & _
            ":mn></m:mtd></m:mtr></m:mtable></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "050001010102020000010002008831000001000200883200000100020088330000010000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("menclose top is the overbar template")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:menclose notation=""top"">" & _
            "<m:mi mathvariant=""bold-italic"">x</m:mi></m:menclose></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "03000D0000010008010302007F780000000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("times and divide on the Symbol font")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mn>2</m:mn><m:mo>&#x00D7" & _
            ";</m:mo><m:mn>3</m:mn><m:mo>&#x00F7;</m:mo><m:mn>4</m:mn></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0200883200020486D700B40200883300020486F700B802008834000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("minus-plus on MT Extra")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mn>2</m:mn><m:mo>&#x2213" & _
            ";</m:mo><m:mn>3</m:mn></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "020088320002048B13226D02008833000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("mtext spaces are collapsed, and a space is a no-break space")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mtext> a  b </m:mtext></" & _
            "m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "0200816100020081A00002008162000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("an empty mo writes nothing")
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mi mathvariant=""bold-ita" & _
            "lic"">x</m:mi><m:mo></m:mo><m:mi mathvariant=""bold-italic"">y</m:mi></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "08010302007F780002007F79000000"
        .IsEqual Sh_MTEF_Not_Handled(x), ""
    End With

    With Suite.Test("letters directly in a one-column table cell stay bold italic (Jerry, 10/20/2026)")
        ' MathType's own conversion makes them plain italic (02 00 83 ..); this book wants them
        ' like every other letter (02 00 7F ..).
        x = "<m:math xmlns:m=""http://www.w3.org/1998/Math/MathML""><m:mtable><m:mtr><m:mtd><m:" & _
            "mi mathvariant=""bold-italic"">x</m:mi><m:mo>=</m:mo><m:mi mathvariant=""bold-itali" & _
            "c"">y</m:mi></m:mtd></m:mtr><m:mtr><m:mtd><m:mi mathvariant=""bold-italic"">y</m:mi" & _
            "></m:mtd></m:mtr></m:mtable></m:math>"
        .IsEqual Body(Sh_MTEF_From_MathML(x)), "050001010102010000010008010302007F78000204863D003D02007F790000010002007F79000000" & _
            "0000"
    End With

    ' --- Step 3: the package a MathType equation goes in with, and the picture's size.
    With Suite.Test("the MathType package carries the equation, the picture and its size")
        h = Sh_MathType_Package("OLEDATA", "PICDATA", "jpg", 89.04, 13.2)
        .IsOk InStr(1, h, "<w:object w:dxaOrig=""1781"" w:dyaOrig=""264"">") > 0
        .IsOk InStr(1, h, "style=""width:89.04pt;height:13.2pt""") > 0
        .IsOk InStr(1, h, "ProgID=""Equation.DSMT4""") > 0
        .IsOk InStr(1, h, "<pkg:binaryData>OLEDATA</pkg:binaryData>") > 0
        .IsOk InStr(1, h, "pkg:name=""/word/media/image1.jpg"" pkg:contentType=""image/jpeg""") > 0
        .IsOk InStr(1, h, "<pkg:binaryData>PICDATA</pkg:binaryData>") > 0
        .IsOk InStr(1, Sh_MathType_Package("A", "B", "png", 12, 12), "pkg:contentType=""image/png""") > 0
    End With

    With Suite.Test("the MathType package is well-formed XML")
        Dim xd As Object
        Set xd = CreateObject("MSXML2.DOMDocument.6.0")
        xd.async = False
        .IsOk xd.loadXML(Sh_MathType_Package("T0xF", "UElD", "png", 12.5, 7))
        .IsEqual xd.parseError.errorCode, 0
    End With

    With Suite.Test("a size is written with a point, whatever the machine's own setting")
        .IsEqual Sh_Points_Text(89.04), "89.04"
        .IsEqual Sh_Points_Text(0.5), "0.5"
        .IsEqual Sh_Points_Text(12), "12"
    End With

    With Suite.Test("a picture's size in points comes from its own bytes")
        ' A 371 x 55 JPEG at 300 dpi (JFIF), the shape of the book's EQ pictures.
        b = Sh_Picture_Size_Pt(Sh_MTEF_Text_Bytes(Sh_MTEF_Hex_Text( _
            "FFD8FFE000104A46494600010201012C012C0000FFC000110800370173030122000211010311" & "01FFD9")))
        .IsEqual Join(Array(b(0), b(1), b(2)), "|"), "89.04|13.2|jpg"
        ' The same as a PNG with a pHYs of 11,811 pixels a meter (300 dpi), and without one (96).
        b = Sh_Picture_Size_Pt(Sh_MTEF_Text_Bytes(Sh_MTEF_Hex_Text( _
            "89504E470D0A1A0A0000000D494844520000017300000037080600000000000000" & _
            "0000000970485973" & "00002E2300002E230100000000")))
        .IsEqual Join(Array(b(0), b(1), b(2)), "|"), "89.04|13.2|png"
        b = Sh_Picture_Size_Pt(Sh_MTEF_Text_Bytes(Sh_MTEF_Hex_Text( _
            "89504E470D0A1A0A0000000D494844520000017300000037080600000000000000" & _
            "0000000049444154" & "00000000")))
        .IsEqual Join(Array(b(0), b(1), b(2)), "|"), "278.25|41.25|png"
        ' A GIF has no density: 96 dpi.
        b = Sh_Picture_Size_Pt(Sh_MTEF_Text_Bytes(Sh_MTEF_Hex_Text( _
            "474946383961730137000000000000000000000000000000")))
        .IsEqual Join(Array(b(0), b(1), b(2)), "|"), "278.25|41.25|gif"
        .IsOk IsEmpty(Sh_Picture_Size_Pt(Sh_MTEF_Text_Bytes("not a picture at all, no")))
        .IsOk IsEmpty(Sh_Picture_Size_Pt(Empty))
    End With

    With Suite.Test("asking whether MathType is installed never raises")
        .IsEqual VarType(Sh_MathType_Is_Installed()), vbBoolean
    End With

    BasicMathTests Suite

    ' --- What is refused gives Empty, and says why.
    With Suite.Test("an element not handled gives Empty and is named")
        x = "<m:math " & MML & "><m:mi mathvariant=""bold-italic"">x</m:mi><m:mphantom><m:mi>y</m:mi></m:mphantom></m:math>"
        .IsOk IsEmpty(Sh_MTEF_From_MathML(x))
        .IsEqual Sh_MTEF_Not_Handled(x), "mphantom is not handled"
    End With

    With Suite.Test("a character not handled gives Empty and is named")
        x = "<m:math " & MML & "><m:mi mathvariant=""double-struck"">R</m:mi></m:math>"
        .IsOk IsEmpty(Sh_MTEF_From_MathML(x))
        .IsEqual Sh_MTEF_Not_Handled(x), "mi mathvariant=""double-struck"" is not handled"
        x = "<m:math " & MML & "><m:mo>&#x230A;</m:mo></m:math>"
        .IsOk IsEmpty(Sh_MTEF_From_MathML(x))
        .IsEqual Sh_MTEF_Not_Handled(x), "the character U+230A in mo is not handled"
    End With

    With Suite.Test("mfenced with two arguments, and brackets that close each other wrongly, are refused")
        x = "<m:math " & MML & "><m:mfenced><m:mn>3</m:mn><m:mn>4</m:mn></m:mfenced></m:math>"
        .IsEqual Sh_MTEF_Not_Handled(x), "mfenced with more than one argument is not handled"
        x = "<m:math " & MML & "><m:mo>[</m:mo><m:mn>3</m:mn><m:mo>)</m:mo></m:math>"
        .IsEqual Sh_MTEF_Not_Handled(x), "brackets that do not match"
    End With

    With Suite.Test("MathML that cannot be read, or is not math, gives Empty")
        .IsOk IsEmpty(Sh_MTEF_From_MathML("<m:math " & MML & "><m:mi>x</m:math>"))
        .IsEqual Sh_MTEF_Not_Handled("<m:math " & MML & "><m:mi>x</m:math>"), "the MathML could not be read"
        .IsEqual Sh_MTEF_Not_Handled("<math><mi>x</mi></math>"), "it is not a MathML math element"
    End With

    ' --- The OLE compound file.
    With Suite.Test("the compound file has the signature, the size and the Equation Native header")
        b = Sh_MTEF_From_MathML("<m:math " & MML & "><m:mn>1</m:mn></m:math>")
        ole = Sh_MTEF_Ole_File(b)
        h = HexOf(ole)
        .IsEqual UBound(ole) - LBound(ole) + 1, 3072
        .IsEqual Left$(h, 16), "D0CF11E0A1B11AE1"
        ' The root entry (sector 1, at 1024) carries MathType's class id at byte 80.
        .IsEqual Mid$(h, 2 * (1024 + 80) + 1, 32), "03CE020000000000C000000000000046"
        ' Equation Native is the fifth 64-byte piece of the mini stream, which starts at sector 4:
        ' 512 + 4 * 512 + 4 * 64 = 2816. Its header: 28, version 2.0, C213, the MTEF's length.
        .IsEqual Mid$(h, 2 * 2816 + 1, 24), "1C000000020013C2C8000000"
        .IsEqual Mid$(h, 2 * (2816 + 28) + 1, 400), HexOf(b)
    End With

    With Suite.Test("a big MTEF goes in ordinary sectors; too big, or not bytes, gives Empty")
        ReDim big(0 To 4999)
        For i = 0 To 4999
            big(i) = (i * 7) Mod 256
        Next i
        ole = Sh_MTEF_Ole_File(big)
        .IsEqual UBound(ole) - LBound(ole) + 1, 8192
        ReDim big(0 To 69999)
        .IsOk IsEmpty(Sh_MTEF_Ole_File(big))
        .IsOk IsEmpty(Sh_MTEF_Ole_File("not bytes"))
        .IsOk IsEmpty(Sh_MTEF_Ole_File(Empty))
    End With

    ' --- Base64.
    With Suite.Test("Base64 is one line, and comes back to the same bytes")
        .IsEqual Sh_Base64(Sh_MTEF_Text_Bytes("Man")), "TWFu"
        .IsEqual Sh_Base64(Sh_MTEF_Text_Bytes(Sh_MTEF_Hex_Text("FBFF"))), "+/8="
        .IsEqual Sh_Base64(Empty), ""
        ole = Sh_MTEF_Ole_File(Sh_MTEF_From_MathML("<m:math " & MML & "><m:mn>1</m:mn></m:math>"))
        h = Sh_Base64(ole)
        .IsEqual Len(h), 4096
        .IsEqual InStr(1, h, vbLf), 0
        .IsEqual InStr(1, h, vbCr), 0
        .IsEqual HexOf(FromBase64(h)), HexOf(ole)
    End With

    Set TestMTEF_Suite = Suite
End Function

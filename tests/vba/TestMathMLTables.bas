Attribute VB_Name = "TestMathMLTables"
'@uses Sh_MathML_Untable
'@uses Sh_MathML_Uncomment
'@uses Sh_MathML_Fragments
'
' Tables of equations taken apart before the import (issue #43, Jerry, 10/27/2026). Any table
' that holds an equation becomes one paragraph per cell, words included, row by row and left
' to right; empty cells give nothing, a caption comes first, and the table is gone. A table
' with no equation is not touched. The first shape is Bookshare's "Basic Math & Pre-Algebra
' For Dummies" one-row "Dummies-Mini-Table". Nothing here touches a document.
'
' Version: 1.1  Date: 10/27/2026 - a table of pictures without "eq" in their names, each with
'                                math in a comment, stays a table (issue #46)
' Version: 1.0  Date: 10/27/2026
'
Option Explicit

Private Function E1() As String
    E1 = "<m:math><m:mrow><m:mn>2</m:mn><m:mo>&#215;</m:mo><m:mn>4</m:mn><m:mo>=</m:mo><m:mn>8</m:mn></m:mrow></m:math>"
End Function

Private Function E2() As String
    E2 = "<m:math altimg=""images/eq01012.png""><m:mn>15</m:mn></m:math>"
End Function

Private Function P(ByVal inner As String) As String
    P = "<p>" & inner & "</p>" & vbLf
End Function

Public Function TestMathMLTables_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Tables of equations taken apart (DAISY/NIMAS)"

    Dim t As String

    With Suite.Test("a one-row table of equations becomes one paragraph per equation, and the table goes")
        t = "<p>a</p><figure><table class=""Dummies-Mini-Table"" id=""table_1""><tbody id=""tb""><tr>" & _
            "<td class=""Mini-TB-Body""><p class=""Mini-TB-Body"">" & E1() & "</p></td>" & _
            "<td class=""Mini-TB-Body""><p class=""Mini-TB-Body"">" & E2() & "</p></td>" & _
            "</tr></tbody></table></figure><p>b</p>"
        .IsEqual Sh_MathML_Untable(t, "m:"), _
            "<p>a</p><figure><p class=""Mini-TB-Body"">" & E1() & "</p>" & vbLf & _
            "<p class=""Mini-TB-Body"">" & E2() & "</p>" & vbLf & "</figure><p>b</p>"
    End With

    With Suite.Test("words come out too, row by row, left to right; words right of an equation follow below it")
        t = "<table><tr><td><p>Threes:</p></td><td><p>" & E1() & "</p></td></tr>" & _
            "<tr><td>" & E2() & "</td><td>is the answer</td></tr></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), _
            "<p>Threes:</p>" & vbLf & P(E1()) & P(E2()) & P("is the answer")
    End With

    With Suite.Test("empty cells give nothing")
        t = "<table><tr><td> <p>" & E1() & "</p> </td><td>" & vbLf & " </td><td><p>&#160;</p></td><td/></tr></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), P(E1())
    End With

    With Suite.Test("a picture alone in a cell is not empty")
        t = "<table><tr><td>" & E1() & "</td><td><img src=""a.png"" /></td></tr></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), P(E1()) & P("<img src=""a.png"" />")
    End With

    With Suite.Test("the caption comes first, then the heading cells, then the body")
        t = "<table><caption>Table 1</caption><thead><tr><th>x</th><th>y</th></tr></thead>" & _
            "<tbody><tr><td>1</td><td>" & E1() & "</td></tr></tbody></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), P("Table 1") & P("x") & P("y") & P("1") & P(E1())
        t = "<table><tr><td>" & E1() & "</td></tr><caption>Cap</caption></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), P("Cap") & P(E1())
    End With

    With Suite.Test("a table with no equation is not touched, beside one that is")
        t = "<table class=""t""><tbody><tr><td>a</td><td>b</td></tr></tbody></table>"
        .IsEqual Sh_MathML_Untable(t & "<p>" & E1() & "</p>", "m:"), t & "<p>" & E1() & "</p>"
        .IsEqual Sh_MathML_Untable(t & "<table><tr><td>" & E1() & "</td></tr></table>", "m:"), t & P(E1())
    End With

    With Suite.Test("nested: an equation table inside a cell is taken apart, and so is the table around it")
        t = "<table><tr><td>w</td><td><table><tr><td>" & E1() & "</td><td>" & E2() & "</td></tr></table></td></tr></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), P("w") & "<p>" & E1() & "</p>" & vbLf & "<p>" & E2() & "</p>" & vbLf
    End With

    With Suite.Test("nested: a table with no equation inside one taken apart stays whole")
        t = "<table><tr><td>" & E1() & "</td><td><table><tr><td>x</td></tr></table></td></tr></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), P(E1()) & "<table><tr><td>x</td></tr></table>" & vbLf
    End With

    With Suite.Test("a page number between two rows is kept in its place")
        t = "<table><tr><td>" & E1() & "</td></tr><p>$pg5</p><tr><td>" & E2() & "</td></tr></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), P(E1()) & "<p>$pg5</p>" & vbLf & P(E2())
    End With

    With Suite.Test("MathML's own tables are not tables, and <thead> is not <th>")
        t = "<m:math><m:mtable><m:mtr><m:mtd><m:mn>1</m:mn></m:mtd></m:mtr></m:mtable></m:math>"
        .IsEqual Sh_MathML_Untable("<table><thead><tr><th>h</th></tr></thead><tr><td>" & t & "</td></tr></table>", "m:"), _
                 P("h") & P(t)
        t = "<math><mtable><mtr><mtd><mn>1</mn></mtd></mtr></mtable></math>"
        .IsEqual Sh_MathML_Untable("<table><tr><td>" & t & "</td></tr></table>", ""), P(t)
    End With

    With Suite.Test("tags in capitals, and a cell left open, still count")
        .IsEqual Sh_MathML_Untable("<TABLE><TR><TD>" & E1() & "</TD></TR></TABLE>", "m:"), P(E1())
        .IsEqual Sh_MathML_Untable("<table><tr><td>a<td>" & E1() & "</tr></table>", "m:"), P("a") & P(E1())
    End With

    With Suite.Test("math with another prefix, or none at all: the text unchanged")
        t = "<table><tr><td><mml:math><mml:mn>1</mml:mn></mml:math></td></tr></table>"
        .IsEqual Sh_MathML_Untable(t, "m:"), t
        .IsEqual Sh_MathML_Untable("<p>none</p>", "m:"), "<p>none</p>"
        .IsEqual Sh_MathML_Untable("", "m:"), ""
    End With

    With Suite.Test("a table that never closes leaves the text as it was")
        t = "<p>a</p><table><tr><td>" & E1() & "</td></tr>"
        .IsEqual Sh_MathML_Untable(t, "m:"), t
    End With

    With Suite.Test("after the uncommenting, every equation is still there to be counted")
        t = "<table class=""Dummies-Mini-Table""><tbody><tr>" & _
            "<td><p><img src=""images/eq01011.png"" alt=""2 times 4 equals 8"" />  <!--" & E1() & "-->  </p></td>" & _
            "<td><p><img src=""images/eq01012.png"" alt=""3 times 5 equals 15"" />  <!--<m:math><m:mn>15</m:mn></m:math>-->  </p></td>" & _
            "<td><p> </p></td></tr></tbody></table>"
        t = Sh_MathML_Uncomment(t, "m:")
        .IsEqual UBound(Sh_MathML_Fragments(t, "m:")), 1
        .IsEqual UBound(Sh_MathML_Fragments(Sh_MathML_Untable(t, "m:"), "m:")), 1
        .IsEqual InStr(1, Sh_MathML_Untable(t, "m:"), "<table"), 0
        .IsEqual InStr(1, Sh_MathML_Untable(t, "m:"), "<td"), 0
    End With

    With Suite.Test("a table of pictures without ""eq"" in their names, math in comments, stays a table (issue #46)")
        t = "<table><tbody><tr>" & _
            "<td><p><img src=""images/fig01.png"" alt=""a"" />  <!--" & E1() & "-->  </p></td>" & _
            "<td><p><img src=""images/eq/fig02.png"" alt=""b"" /><!--<m:math><m:mn>15</m:mn></m:math>--></p></td>" & _
            "</tr></tbody></table>"
        t = Sh_MathML_Uncomment(t, "m:")
        .IsEqual t, "<table><tbody><tr>" & _
            "<td><p><img src=""images/fig01.png"" alt=""a"" />    </p></td>" & _
            "<td><p><img src=""images/eq/fig02.png"" alt=""b"" /></p></td>" & _
            "</tr></tbody></table>"
        .IsEqual UBound(Sh_MathML_Fragments(t, "m:")), -1
        .IsEqual Sh_MathML_Untable(t, "m:"), t
    End With

    Set TestMathMLTables_Suite = Suite
End Function

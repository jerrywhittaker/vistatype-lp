Attribute VB_Name = "TestTocBox"
'@uses Lp_Tocb_Drop_Last
'@uses Lp_Tocb_Held_End
'
' The two decisions behind the TOC box that stays open (354), Jerry, 9/22/2026: which paragraphs
' the held range covers, and where it ends once a job has rewritten the TOC. Numbers in and out,
' because anything reaching a Range hangs this runner (tests/vba/README.md).
'
' Version: 1.0  Date: 9/22/2026
'
Option Explicit

Public Function TestTocBox_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "The TOC box that stays open"

    ' --- a last paragraph the selection only reaches into is left out
    With Suite.Test("an over-drag into the next paragraph is dropped")
        ' the last paragraph runs to 500; the selection stopped at 420, part way into it
        .IsOk Lp_Tocb_Drop_Last(5, 420, 500)
    End With

    With Suite.Test("a last line selected to its end is kept")
        ' to the end of its text, which is one short of its paragraph mark
        .NotOk Lp_Tocb_Drop_Last(5, 499, 500)
        ' and with the mark itself
        .NotOk Lp_Tocb_Drop_Last(5, 500, 500)
    End With

    With Suite.Test("a single paragraph is always kept")
        .NotOk Lp_Tocb_Drop_Last(1, 420, 500)
    End With

    ' --- the range taken again after a job
    With Suite.Test("the end follows a TOC that grew or shrank")
        ' 300 characters after the TOC in a book now 1,200 long: the TOC ends at 900
        .IsEqual Lp_Tocb_Held_End(1200, 300, 100), 900
        ' the same book after a job took 150 characters out of the TOC
        .IsEqual Lp_Tocb_Held_End(1050, 300, 100), 750
    End With

    With Suite.Test("a TOC at the end of the book ends at the end of the book")
        .IsEqual Lp_Tocb_Held_End(1200, 0, 100), 1200
    End With

    With Suite.Test("a TOC that has gone comes back empty, never backwards")
        .IsEqual Lp_Tocb_Held_End(350, 300, 100), 100
    End With

    Set TestTocBox_Suite = Suite
End Function

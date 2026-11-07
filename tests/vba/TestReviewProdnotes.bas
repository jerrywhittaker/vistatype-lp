Attribute VB_Name = "TestReviewProdnotes"
'@uses Dx_Rvp_Step
'@uses Dx_Rvp_Label
'@uses Dx_Rvp_Tally
'@uses Dx_Rvp_Joins
'@uses Dx_Rvp_Bare
'@uses Dx_Rvp_Collect
'@uses Dx_Rvp_Live
'@uses Dx_Rvp_Delete_Run
'@uses Dx_Rvp_Restyle_Run
'@uses Dx_Rvp_Is_BANA_Name
'
' Review Prodnotes (418), Jerry, 11/7/2026: which prodnote Next and Previous go to, the lines the
' box shows, and the rule that groups Prodnote paragraphs into one prodnote - and then, in a hidden
' scratch book of its own, that three prodnotes (one of them two paragraphs long) are read as
' three, that Delete takes both paragraphs of the long one, that a prodnote ending the book leaves
' a Normal paragraph behind, and that Convert gives the TranscriberNote style.
'
' THE SCRATCH BOOK IS THE ONE EXCEPTION here to "nothing that reaches a Document" (README.md):
' it is made with Visible:=False, it never touches the Selection, it makes its own two styles so
' no template is attached, and it is closed unsaved whatever happens. Every error is caught and
' reported as a failure, because an unhandled one is a modal dialog in an invisible Word.
'
' Version: 1.1  Date: 11/7/2026 - which template names count as BANA, for Print Layout and Convert
' Version: 1.0  Date: 11/7/2026
'
Option Explicit

Public Function TestReviewProdnotes_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Review Prodnotes goes through the prodnotes one at a time"

    With Suite.Test("Next goes to the prodnote after, and round from the last to the first")
        .IsEqual Dx_Rvp_Step(1, 3, True), 2
        .IsEqual Dx_Rvp_Step(2, 3, True), 3
        .IsEqual Dx_Rvp_Step(3, 3, True), 1
    End With

    With Suite.Test("Previous goes to the prodnote before, and round from the first to the last")
        .IsEqual Dx_Rvp_Step(3, 3, False), 2
        .IsEqual Dx_Rvp_Step(1, 3, False), 3
    End With

    With Suite.Test("the box says which prodnote of how many, and what was done")
        .IsEqual Dx_Rvp_Label(3, 41), "Prodnote 3 of 41"
        .IsEqual Dx_Rvp_Label(1, 1200), "Prodnote 1 of 1,200"
        .IsEqual Dx_Rvp_Tally(2, 1), "Deleted: 2     Converted to TN Style: 1"
    End With

    With Suite.Test("a Prodnote paragraph joins the one before only in the same place")
        .IsEqual Dx_Rvp_Joins(True, False, False, False), True     ' body text, one after another
        .IsEqual Dx_Rvp_Joins(True, False, True, True), True       ' the same table cell
        .IsEqual Dx_Rvp_Joins(False, False, False, False), False   ' the one before was not one
        .IsEqual Dx_Rvp_Joins(True, True, True, True), False       ' the one before ended its cell
        .IsEqual Dx_Rvp_Joins(True, False, False, True), False     ' into a table
        .IsEqual Dx_Rvp_Joins(True, False, True, False), False     ' out of a table
    End With

    With Suite.Test("an empty paragraph is empty, whatever marks it ends in")
        .IsEqual Dx_Rvp_Bare("  " & vbCr), ""
        .IsEqual Dx_Rvp_Bare(vbCr & Chr$(7)), ""
        .IsEqual Dx_Rvp_Bare(" A picture of a cat." & vbCr), "A picture of a cat."
    End With

    With Suite.Test("a template counts as BANA when its name begins BANA Braille")
        .IsEqual Dx_Rvp_Is_BANA_Name("BANA Braille 2024.dotx"), True
        .IsEqual Dx_Rvp_Is_BANA_Name("bana braille 2017.dot"), True
        .IsEqual Dx_Rvp_Is_BANA_Name("LargePrintTemplate.dotx"), False
        .IsEqual Dx_Rvp_Is_BANA_Name("Normal.dotm"), False
        .IsEqual Dx_Rvp_Is_BANA_Name(""), False
    End With

    TestReviewProdnotes_In_A_Book Suite.Test("in a scratch book: three prodnotes, then Delete and Convert")

    Set TestReviewProdnotes_Suite = Suite
End Function

' The book: Chapter 1 / Prodnote A / Some text. / Prodnote B, two paragraphs / More text. /
' Prodnote C, the last paragraph of the book.
Private Sub TestReviewProdnotes_In_A_Book(ByVal T As TestCase)
    Dim d As Document
    Dim stProd As Style, stTN As Style
    Dim hits As Collection
    Dim n As Long, before As Long
    Dim stage As String

    On Error GoTo Failed
    stage = "making the scratch book"
    Set d = Documents.Add(Visible:=False)
    Set stProd = d.Styles.Add("Prodnote", wdStyleTypeParagraph)
    Set stTN = d.Styles.Add("TranscriberNote", wdStyleTypeParagraph)
    d.Content.Text = "Chapter 1" & vbCr & "Prodnote A" & vbCr & "Some text." & vbCr _
                   & "Prodnote B, first part" & vbCr & "Prodnote B, second part" & vbCr _
                   & "More text." & vbCr & "Prodnote C"
    For n = 1 To d.Paragraphs.Count
        If n = 2 Or n = 4 Or n = 5 Or n = 7 Then d.Paragraphs(n).Style = stProd
    Next n
    T.IsEqual d.Paragraphs.Count, 7, "the scratch book has 7 paragraphs"

    stage = "finding the prodnotes"
    Set hits = Dx_Rvp_Collect(d)
    T.IsEqual hits.Count, 3, "three prodnotes"
    T.IsEqual hits(1).Paragraphs.Count, 1, "A is one paragraph"
    T.IsEqual hits(2).Paragraphs.Count, 2, "B is two paragraphs"
    T.IsEqual hits(3).Paragraphs.Count, 1, "C is one paragraph"
    T.IsEqual Dx_Rvp_Live(hits(2)), True, "B is live before the delete"

    stage = "deleting B, two paragraphs"
    before = d.Paragraphs.Count
    Dx_Rvp_Delete_Run hits(2)
    T.IsEqual d.Paragraphs.Count, before - 2, "both of B's paragraphs are gone"
    T.IsEqual InStr(1, d.Content.Text, "Prodnote B"), 0, "none of B's text is left"
    T.IsEqual Dx_Rvp_Live(hits(2)), False, "B is not live after the delete"
    T.IsEqual Dx_Rvp_Live(hits(3)), True, "C is still live"
    T.IsEqual d.Paragraphs(4).Range.Text, "More text." & vbCr, "the text after B is untouched"

    stage = "deleting C, the end of the book"
    Dx_Rvp_Delete_Run hits(3)
    T.IsEqual InStr(1, d.Content.Text, "Prodnote C"), 0, "C's text is gone"
    T.IsEqual Dx_Rvp_Bare(d.Paragraphs.Last.Range.Text), "", "the book's last mark is left empty"
    T.IsEqual d.Paragraphs.Last.Style.NameLocal, d.Styles(wdStyleNormal).NameLocal, _
              "and set back to Normal"

    stage = "converting A"
    Dx_Rvp_Restyle_Run hits(1), stTN
    T.IsEqual d.Paragraphs(2).Style.NameLocal, "TranscriberNote", "A is a TranscriberNote"
    T.IsEqual d.Paragraphs(2).Range.Text, "Prodnote A" & vbCr, "A's text is kept"
    T.IsEqual Dx_Rvp_Live(hits(1)), False, "A is not a prodnote any more"

    d.Close SaveChanges:=False
    Exit Sub

Failed:
    T.Fail "VBA error " & Err.Number & " while " & stage & " - " & Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.Close SaveChanges:=False
    Err.Clear
End Sub

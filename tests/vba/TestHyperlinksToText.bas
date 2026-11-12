Attribute VB_Name = "TestHyperlinksToText"
'@uses Dx_Hyperlinks_To_Text
'
' Braille Full File Cleanup stopped with 4680 on a web link attached to a picture (issue #54,
' Jerry, 11/12/2026). In a hidden scratch book: a link on text becomes its address; a link on a
' floating shape has its address written on a line of its own after the paragraph the shape is
' attached to, and the link comes off the shape; a link within the book is left alone; and a
' second run changes nothing. Made with Visible:=False, never touches the Selection, closed
' unsaved whatever happens, every error caught as a failure (README.md).
'
' Version: 1.0  Date: 11/12/2026
'
Option Explicit

Public Function TestHyperlinksToText_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Web links become text, on words and on pictures (issue #54)"
    TestHyperlinksToText_In_A_Book Suite.Test("a link on a picture is written out after its paragraph, without 4680")
    Set TestHyperlinksToText_Suite = Suite
End Function

' The book: Follow us: / See the index. / The end.
' "Follow us:" carries two shapes, each linked to a web address; "index" is a link within the
' book; "The end" has a web link on its words.
Private Sub TestHyperlinksToText_In_A_Book(ByVal T As TestCase)
    Dim d As Document
    Dim shp1 As Shape, shp2 As Shape
    Dim r As Range
    Dim stage As String, txt As String

    On Error GoTo Failed
    stage = "making the scratch book"
    Set d = Documents.Add(Visible:=False)
    d.Content.Text = "Follow us:" & vbCr & "See the index." & vbCr & "The end."
    d.Bookmarks.Add "Ix", d.Paragraphs(3).Range
    Set r = d.Paragraphs(1).Range
    Set shp1 = d.Shapes.AddShape(msoShapeRectangle, 10, 10, 20, 20, r)
    Set shp2 = d.Shapes.AddShape(msoShapeRectangle, 40, 10, 20, 20, r)
    d.Hyperlinks.Add Anchor:=shp1, Address:="https://www.facebook.com/fordummies"
    d.Hyperlinks.Add Anchor:=shp2, Address:="https://www.youtube.com/user/fordummies"
    Set r = d.Paragraphs(2).Range
    r.SetRange r.Start + 8, r.Start + 13
    d.Hyperlinks.Add Anchor:=r, Address:="", SubAddress:="Ix"
    Set r = d.Paragraphs(3).Range
    r.SetRange r.Start + 4, r.Start + 7
    d.Hyperlinks.Add Anchor:=r, Address:="https://www.dummies.com/go"
    T.IsEqual d.Hyperlinks.Count, 4, "the scratch book has 4 links"
    T.IsEqual d.Hyperlinks(1).Type + d.Hyperlinks(2).Type + d.Hyperlinks(3).Type + d.Hyperlinks(4).Type, _
              2, "two of them are on shapes (msoHyperlinkShape = 1)"

    stage = "turning the links into text"
    Dx_Hyperlinks_To_Text d
    txt = d.Content.Text
    T.IsEqual txt, "Follow us:" & vbCr & "https://www.facebook.com/fordummies" & vbCr _
                 & "https://www.youtube.com/user/fordummies" & vbCr & "See the index." & vbCr _
                 & "The https://www.dummies.com/go." & vbCr, _
              "each picture's address on its own line, in order, and the text link replaced - got: " _
              & Replace(txt, vbCr, " | ")
    T.IsEqual d.Hyperlinks.Count, 1, "only the link within the book is left"
    T.IsEqual d.Hyperlinks(1).SubAddress, "Ix", "and it is the link within the book"
    T.IsEqual d.Shapes.Count, 2, "both pictures are kept"

    stage = "running it a second time"
    Dx_Hyperlinks_To_Text d
    T.IsEqual d.Content.Text, txt, "a second run writes nothing again"

    d.Close SaveChanges:=False
    Exit Sub

Failed:
    T.Fail "VBA error " & Err.Number & " while " & stage & " - " & Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.Close SaveChanges:=False
    Err.Clear
End Sub

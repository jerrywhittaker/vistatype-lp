Attribute VB_Name = "TestMinPictureSize"
'@uses Lp_Min_Pic_Factor
'@uses Lp_Enlarge_Tiny_Pictures
'@uses Lp_Same_Pic_Is_Picture
'
' No picture smaller than the text, Jerry, 11/1/2026: the LP template attach enlarges a picture
' whose shorter side is under the Normal style's font size until it is that size, keeping its
' shape. Lp_Min_Pic_Factor is that decision, handed numbers rather than a picture, because
' anything reaching a picture hangs this runner (tests/vba/README.md).
'
' A picture that grew on every attach, Jerry, 11/8/2026: ten pictures in A.docx just under 18
' point grew their long side a little each time the template was attached again. Word reports a
' size cut down to whole twips (17.99 reads 17.95), and the short side was rounded down to 17.99,
' so it was never big enough. In a hidden scratch book of its own, with its own picture file in
' %TEMP%, Lp_Enlarge_Tiny_Pictures is run twice on a 17.99 x 29.36 picture and once more on one
' at exactly 18: the first must reach 18 and then stay put, the second must never be touched.
' The scratch book is the same exception TestReviewProdnotes makes (README.md): Visible:=False,
' no Selection, closed unsaved, every error caught as a failure.
'
' Version: 1.1  Date: 11/8/2026 - a picture just under the text size is enlarged once, then left alone
' Version: 1.0  Date: 11/1/2026
'
Option Explicit

Public Function TestMinPictureSize_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "No picture smaller than the text"

    With Suite.Test("a dot is enlarged to the text size")
        .IsEqual Lp_Min_Pic_Factor(2, 2, 18), 9
    End With

    With Suite.Test("a sliver is enlarged by its shorter side")
        .IsEqual Lp_Min_Pic_Factor(2, 100, 18), 9
        .IsEqual Lp_Min_Pic_Factor(100, 2, 18), 9
    End With

    With Suite.Test("a picture big enough is left alone, never shrunk")
        .IsEqual Lp_Min_Pic_Factor(200, 100, 18), 1
        .IsEqual Lp_Min_Pic_Factor(18, 18, 18), 1
        .IsEqual Lp_Min_Pic_Factor(18, 400, 18), 1
    End With

    With Suite.Test("sizes of zero or less change nothing")
        .IsEqual Lp_Min_Pic_Factor(0, 10, 18), 1
        .IsEqual Lp_Min_Pic_Factor(10, 0, 18), 1
        .IsEqual Lp_Min_Pic_Factor(10, 10, 0), 1
    End With

    TestMinPictureSize_Twice Suite.Test("a picture just under the text size is enlarged once, then left alone")

    Set TestMinPictureSize_Suite = Suite
End Function

' A plain 24 x 40 pixel bitmap, 24 bits a pixel, gray. Only something AddPicture will take.
Private Sub TestMinPictureSize_Write_Bmp(ByVal path As String)
    Dim b() As Byte
    Dim n As Long, k As Long, f As Integer

    n = 54 + 72 * 40                ' header, then 40 rows of 24 pixels x 3 bytes (72, no padding)
    ReDim b(0 To n - 1)
    b(0) = 66: b(1) = 77            ' "BM"
    TestMinPictureSize_Put_Long b, 2, n
    TestMinPictureSize_Put_Long b, 10, 54
    TestMinPictureSize_Put_Long b, 14, 40
    TestMinPictureSize_Put_Long b, 18, 24
    TestMinPictureSize_Put_Long b, 22, 40
    b(26) = 1                       ' one plane
    b(28) = 24                      ' bits a pixel
    TestMinPictureSize_Put_Long b, 38, 3780     ' 96 dpi, as pixels a meter
    TestMinPictureSize_Put_Long b, 42, 3780
    For k = 54 To n - 1
        b(k) = 128
    Next k

    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, b
    Close #f
End Sub

Private Sub TestMinPictureSize_Put_Long(ByRef b() As Byte, ByVal at As Long, ByVal v As Long)
    b(at) = v And &HFF&
    b(at + 1) = (v \ &H100&) And &HFF&
    b(at + 2) = (v \ &H10000) And &HFF&
    b(at + 3) = (v \ &H1000000) And &HFF&
End Sub

' The book: two paragraphs, a picture at the start of each. The first is 17.99 x 29.36 - the size
' of UF613 in Jerry's A.docx - and the second exactly 18 x 30.
Private Sub TestMinPictureSize_Twice(ByVal T As TestCase)
    Dim d As Document
    Dim r As Range
    Dim small As InlineShape, exact As InlineShape
    Dim path As String, stage As String
    Dim w1 As Single, h1 As Single

    On Error GoTo Failed
    stage = "writing the picture file"
    path = Environ$("TEMP") & "\VistaType-minpic-test.bmp"
    TestMinPictureSize_Write_Bmp path

    stage = "making the scratch book"
    Set d = Documents.Add(Visible:=False)
    d.Styles(wdStyleNormal).Font.Size = 18
    d.Content.Text = "One" & vbCr & "Two"
    Set r = d.Paragraphs(2).Range
    r.Collapse wdCollapseStart
    Set exact = d.InlineShapes.AddPicture(FileName:=path, LinkToFile:=False, _
                                          SaveWithDocument:=True, Range:=r)
    Set r = d.Paragraphs(1).Range
    r.Collapse wdCollapseStart
    Set small = d.InlineShapes.AddPicture(FileName:=path, LinkToFile:=False, _
                                          SaveWithDocument:=True, Range:=r)
    T.IsEqual d.InlineShapes.Count, 2, "two pictures in the scratch book"
    T.IsOk Lp_Same_Pic_Is_Picture(small), "the scratch picture counts as a picture"

    stage = "sizing the two pictures"
    small.LockAspectRatio = msoFalse
    small.Width = 17.99
    small.Height = 29.36
    exact.LockAspectRatio = msoFalse
    exact.Width = 18
    exact.Height = 30
    T.IsOk small.Width < 18, "the small picture starts under 18 point (" & small.Width & ")"
    T.IsEqual exact.Width, CSng(18), "the other picture starts at exactly 18 point wide"
    T.IsEqual exact.Height, CSng(30), "and 30 point high"

    stage = "the first attach"
    Lp_Enlarge_Tiny_Pictures d
    w1 = small.Width
    h1 = small.Height
    T.IsOk w1 >= 18, "after the first attach the short side is 18 or more (" & w1 & ")"
    T.IsOk h1 > 29.3 And h1 < 29.5, "and the long side kept its shape (" & h1 & ")"
    T.IsEqual exact.Width, CSng(18), "the picture at exactly 18 is not touched (width)"
    T.IsEqual exact.Height, CSng(30), "the picture at exactly 18 is not touched (height)"

    stage = "the second attach"
    Lp_Enlarge_Tiny_Pictures d
    T.IsEqual small.Width, w1, "the second attach leaves the width alone (" & small.Width & ")"
    T.IsEqual small.Height, h1, "the second attach leaves the height alone (" & small.Height & ")"
    T.IsEqual exact.Width, CSng(18), "the picture at exactly 18 is still not touched (width)"
    T.IsEqual exact.Height, CSng(30), "the picture at exactly 18 is still not touched (height)"

    d.Close SaveChanges:=False
    Set d = Nothing
    Kill path
    Exit Sub

Failed:
    T.Fail "VBA error " & Err.Number & " while " & stage & " - " & Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.Close SaveChanges:=False
    Kill path
    Err.Clear
End Sub

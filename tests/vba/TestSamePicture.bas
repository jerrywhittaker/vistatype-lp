Attribute VB_Name = "TestSamePicture"
'@uses Lp_Rst_In_Range
'@uses Lp_Same_Pic_Mark_May_Go
'@uses Lp_Same_Pic_Alt_Matches
'@uses Lp_Same_Pic_Ratio_Close
'@uses Lp_Same_Pic_Size_Differs
'@uses Lp_Same_Pic_Align_Value
'
' The decisions behind Resize Pictures Used Throughout, and Resize Pictures Used Throughout a
' Selected Range (Jerry, 9/22/2026). Each one is handed numbers and flags rather than a picture
' or a Range, because anything reaching a Range hangs this runner (tests/vba/README.md) - which
' is exactly why the decisions were split out of the procedures that touch the document.
'
' The four marks that must NEVER be replaced were each proved in Word on the build box, 9/22/2026,
' on a document built from this source: the end of a table cell, the last mark in the document, a
' mark with a table starting on the next line, and - checked by reading only, because Word would
' not let automation build the layout - a mark in front of a table nested in the picture's own
' cell.
'
' Version: 1.0  Date: 9/22/2026
'
Option Explicit

Public Function TestSamePicture_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Resizing one picture wherever it repeats"

    ' --- is this picture inside the range the transcriber picked?
    With Suite.Test("a picture inside the range counts")
        .IsOk Lp_Rst_In_Range(120, 121, wdMainTextStory, 100, 200, wdMainTextStory)
        ' sitting exactly on the start, and exactly on the end
        .IsOk Lp_Rst_In_Range(100, 101, wdMainTextStory, 100, 200, wdMainTextStory)
        .IsOk Lp_Rst_In_Range(199, 200, wdMainTextStory, 100, 200, wdMainTextStory)
    End With

    With Suite.Test("a picture outside the range does not")
        ' one character in front of it, and one past it
        .NotOk Lp_Rst_In_Range(99, 100, wdMainTextStory, 100, 200, wdMainTextStory)
        .NotOk Lp_Rst_In_Range(200, 201, wdMainTextStory, 100, 200, wdMainTextStory)
        ' hanging over the end
        .NotOk Lp_Rst_In_Range(195, 205, wdMainTextStory, 100, 200, wdMainTextStory)
    End With

    With Suite.Test("a picture in another story is never in range")
        ' the same numbers, in a header: a position means nothing across stories
        .NotOk Lp_Rst_In_Range(120, 121, wdPrimaryHeaderStory, 100, 200, wdMainTextStory)
    End With

    With Suite.Test("a range of nothing holds nothing")
        .NotOk Lp_Rst_In_Range(100, 101, wdMainTextStory, 100, 100, wdMainTextStory)
        ' and a range handed over backwards is refused rather than guessed at
        .NotOk Lp_Rst_In_Range(120, 121, wdMainTextStory, 200, 100, wdMainTextStory)
    End With

    ' --- may the paragraph mark after this picture become a space?
    With Suite.Test("a plain paragraph mark after a picture may go")
        .IsOk Lp_Same_Pic_Mark_May_Go(vbCr, 50, 900, False, False, False)
        ' inside a table cell, with more of the cell below it
        .IsOk Lp_Same_Pic_Mark_May_Go(vbCr, 50, 900, True, True, True)
    End With

    With Suite.Test("anything that is not a lone paragraph mark is left alone")
        .NotOk Lp_Same_Pic_Mark_May_Go(" ", 50, 900, False, False, False)
        .NotOk Lp_Same_Pic_Mark_May_Go("T", 50, 900, False, False, False)
        ' the end of a table cell comes back as two characters, Chr(13) & Chr(7)
        .NotOk Lp_Same_Pic_Mark_May_Go(vbCr & Chr$(7), 50, 900, True, False, False)
        ' a page or section break is Chr(12), not a paragraph mark
        .NotOk Lp_Same_Pic_Mark_May_Go(Chr$(12), 50, 900, False, False, False)
    End With

    With Suite.Test("the last mark in the story is left alone")
        ' Word puts it straight back, so taking it is not a change that can be made
        .NotOk Lp_Same_Pic_Mark_May_Go(vbCr, 900, 900, False, False, False)
        .NotOk Lp_Same_Pic_Mark_May_Go(vbCr, 901, 900, False, False, False)
    End With

    With Suite.Test("a mark in front of a table is left alone")
        ' the picture is outside the table that starts on the next line
        .NotOk Lp_Same_Pic_Mark_May_Go(vbCr, 50, 900, False, True, False)
        ' and a table NESTED in the picture's own cell: both read as in a table
        .NotOk Lp_Same_Pic_Mark_May_Go(vbCr, 50, 900, True, True, False)
    End With

    ' --- is this the same picture?
    With Suite.Test("the same description matches, whatever its case or spacing")
        .IsOk Lp_Same_Pic_Alt_Matches("Lesson icon", "Lesson icon")
        .IsOk Lp_Same_Pic_Alt_Matches("lesson ICON", "Lesson icon")
        .IsOk Lp_Same_Pic_Alt_Matches("  Lesson icon  ", "Lesson icon")
    End With

    With Suite.Test("a different description does not match")
        .NotOk Lp_Same_Pic_Alt_Matches("Think about it", "Lesson icon")
    End With

    With Suite.Test("no description matches nothing at all")
        ' a book that leaves pictures undescribed must not count them all as one picture
        .NotOk Lp_Same_Pic_Alt_Matches("", "Lesson icon")
        .NotOk Lp_Same_Pic_Alt_Matches("Lesson icon", "")
        .NotOk Lp_Same_Pic_Alt_Matches("", "")
        .NotOk Lp_Same_Pic_Alt_Matches("   ", "Lesson icon")
    End With

    With Suite.Test("two pictures of the same shape match, at any size")
        ' 2 inches by 1, and the same artwork shown at a quarter of that
        .IsOk Lp_Same_Pic_Ratio_Close(144, 72, 36, 18)
        ' 4 per cent out: a copy that has been nudged or cropped a little is still the same one
        .IsOk Lp_Same_Pic_Ratio_Close(144, 72, 144, 75)
    End With

    With Suite.Test("pictures of a different shape do not match")
        ' 6 per cent out, and a square against a wide one
        .NotOk Lp_Same_Pic_Ratio_Close(144, 72, 144, 77)
        .NotOk Lp_Same_Pic_Ratio_Close(144, 72, 72, 72)
    End With

    With Suite.Test("a picture with no height is refused rather than divided by")
        .NotOk Lp_Same_Pic_Ratio_Close(144, 0, 144, 72)
        .NotOk Lp_Same_Pic_Ratio_Close(144, 72, 144, 0)
        .NotOk Lp_Same_Pic_Ratio_Close(0, 72, 144, 72)
    End With

    ' --- does this copy need resizing?
    With Suite.Test("a copy already the right size is left alone")
        .NotOk Lp_Same_Pic_Size_Differs(144, 72, 144, 72)
        ' within half a point either way: Word keeps sizes to several decimal places
        .NotOk Lp_Same_Pic_Size_Differs(144.4, 72.4, 144, 72)
        .NotOk Lp_Same_Pic_Size_Differs(143.6, 71.6, 144, 72)
    End With

    With Suite.Test("a copy half a point out or more is resized")
        .IsOk Lp_Same_Pic_Size_Differs(144.5, 72, 144, 72)
        .IsOk Lp_Same_Pic_Size_Differs(144, 72.5, 144, 72)
        .IsOk Lp_Same_Pic_Size_Differs(72, 36, 144, 72)
    End With

    ' --- what the position buttons mean
    With Suite.Test("the position choice becomes a Word alignment")
        .IsEqual Lp_Same_Pic_Align_Value(1), wdAlignParagraphLeft
        .IsEqual Lp_Same_Pic_Align_Value(2), wdAlignParagraphCenter
    End With

    With Suite.Test("leave as is, and anything unrecognized, touches nothing")
        ' -1 is "do not touch it". A number nobody recognizes must not left align a book.
        .IsEqual Lp_Same_Pic_Align_Value(0), -1
        .IsEqual Lp_Same_Pic_Align_Value(9), -1
        .IsEqual Lp_Same_Pic_Align_Value(-3), -1
    End With

    Set TestSamePicture_Suite = Suite
End Function

Attribute VB_Name = "TestTocBoldEntries"
'@uses Lp_TOC_Bold_Marks_Headings
'@uses Lp_TOC_Break_Ends_Entry
'@uses Lp_TOC_Count_Text
'
' A contents page where nearly every entry is bold, and an entry whose page number is followed by
' a manual line break. Jerry, 9/28/2026, on "TOC with Black Boxes.docx": 63 wholly bold entries
' were read as section headings, and four "Chapter Practice 79<line break>..." entries got no
' tab, because the page number was not at the end of the paragraph. The counts and lines below
' are that file's, and "Table of Contents with Bold.docx"'s.
'
' Numbers and strings only: counting the lines off a Range would hang this runner
' (tests/vba/README.md).
'
' Version: 1.1  Date: 9/28/2026 - the break must come before a bullet; the count's reading of a line.
' Version: 1.0  Date: 9/28/2026
'
Option Explicit

Public Function TestTocBoldEntries_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Contents pages that are bold throughout, and line breaks after a page number"

    With Suite.Test("bold does not mark a heading when most entries are bold")
        ' The shape of "TOC with Black Boxes.docx" once its squares are off: the lines ending in
        ' a page number are bold nearly without exception. Jerry: "yes, bold doesn't mean
        ' heading there".
        .NotOk Lp_TOC_Bold_Marks_Headings(156, 0)
        .NotOk Lp_TOC_Bold_Marks_Headings(20, 20)
    End With

    With Suite.Test("bold marks a heading when most entries are plain")
        ' "Table of Contents with Bold.docx", 9/7/2026: "Section 1" and "Section 2" are bold and
        ' end in a number; the entries are plain. Those two must stay headings.
        .IsOk Lp_TOC_Bold_Marks_Headings(2, 40)
        .IsOk Lp_TOC_Bold_Marks_Headings(19, 20)
    End With

    With Suite.Test("nothing to count keeps the rule as it was")
        .IsOk Lp_TOC_Bold_Marks_Headings(0, 0)
    End With

    With Suite.Test("the count reads a tabbed, leadered or trailing-space line as ending in its number")
        ' Found in review, 9/28/2026: none of these were counted, so a page of plain tabbed
        ' entries with two bold "Section 1" headings counted 2 bold and 0 plain.
        .IsEqual Lp_TOC_Count_Text("Chapter 5" & vbTab & "36" & vbCr), "Chapter 5 36"
        .IsEqual Lp_TOC_Count_Text("Chapter 5.....36" & vbCr), "Chapter 5 36"
        .IsEqual Lp_TOC_Count_Text("Chapter 5 . . . . 36"), "Chapter 5 36"
        .IsEqual Lp_TOC_Count_Text("Title 45 " & ChrW(160) & vbCr), "Title 45"
        .IsEqual Lp_TOC_Count_Text("Chapter 5" & ChrW(&H2026) & "36"), "Chapter 5 36"
    End With

    With Suite.Test("the count leaves a single period alone")
        .IsEqual Lp_TOC_Count_Text("Lesson 1.1 Place Value 3" & vbCr), "Lesson 1.1 Place Value 3"
        .IsEqual Lp_TOC_Count_Text("Ch. 4 Fractions 45"), "Ch. 4 Fractions 45"
        .IsEqual Lp_TOC_Count_Text(""), ""
    End With

    With Suite.Test("a line break after an entry's page number, before a bullet, ends the entry")
        ' The fault: "TOC with Black Boxes.docx", the legend after "Chapter Practice 79".
        .IsOk Lp_TOC_Break_Ends_Entry("Chapter Practice 79 ", ChrW(&H25A0) & " Major Topic")
        .IsOk Lp_TOC_Break_Ends_Entry("Chapter Practice 79", ChrW(&H25A0) & " Major Topic")
        .IsOk Lp_TOC_Break_Ends_Entry("Chapter Practice" & ChrW(160) & "79" & ChrW(160), " " & ChrW(&H2022) & " Next")
    End With

    With Suite.Test("a legend line is not an entry")
        .NotOk Lp_TOC_Break_Ends_Entry(ChrW(&H25A0) & " Major Topic", ChrW(&H25A0) & " Supporting Topic")
        .NotOk Lp_TOC_Break_Ends_Entry("Major Topic", ChrW(&H25A0) & " Additional Topic")
    End With

    With Suite.Test("a heading wrapped with a line break is not split")
        ' Found in review, 9/28/2026: the first version split these, and "Unit 1" became an entry.
        .NotOk Lp_TOC_Break_Ends_Entry("Unit 1", "Place Value and Numbers")
        .NotOk Lp_TOC_Break_Ends_Entry("Part II", "Measurement")
    End With

    With Suite.Test("one entry wrapped with a line break is not split")
        .NotOk Lp_TOC_Break_Ends_Entry("Chapter 3", "Fractions 45")
        .NotOk Lp_TOC_Break_Ends_Entry("Chapter 3", "Fractions 45" & vbCr)
    End With

    With Suite.Test("a break with nothing after it is not split")
        .NotOk Lp_TOC_Break_Ends_Entry("Chapter Practice 79", "")
        .NotOk Lp_TOC_Break_Ends_Entry("Chapter Practice 79", "   ")
    End With

    Set TestTocBoldEntries_Suite = Suite
End Function

Attribute VB_Name = "TestTocPageNumbers"
'@uses Lp_TOC_Line_Has_Page_Number
'
' Does a contents line already carry its own page number? Asked of the line above a paragraph
' that is nothing but a number, to decide whether that number should be pulled up onto it.
'
' The fault these were written for, reported by Jerry 9/20/2026 against "Problem TOC 2.docx":
' "Lesson 2 Add and Subtract 10 or 100" is on page 143, and the entry came out carrying 100.
' A bare space in front of a trailing number used to count as a page number, so a title that
' merely ended in a digit was read as a line that already had one.
'
' Version: 1.0  Date: 9/20/2026
'
Option Explicit

Public Function TestTocPageNumbers_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Contents lines and their page numbers"

    With Suite.Test("a bare space in front of the number is not evidence")
        ' The decision this whole change turns on. Without these, the suite would still be
        ' green if a space were let back into the pattern.
        .NotOk Lp_TOC_Line_Has_Page_Number("Chapter One   7" & vbCr)
        .NotOk Lp_TOC_Line_Has_Page_Number("Chapter One 7" & vbCr)
        .NotOk Lp_TOC_Line_Has_Page_Number("Foreword xiv" & vbCr)
    End With

    With Suite.Test("a title ending in a number is not a page number")
        ' The reported line, exactly as it stands in the file.
        .NotOk Lp_TOC_Line_Has_Page_Number("Lesson 2 Add and Subtract 10 or 100" & vbCr)
        .NotOk Lp_TOC_Line_Has_Page_Number("Module 6" & vbCr)
        .NotOk Lp_TOC_Line_Has_Page_Number("Chapter 12" & vbCr)
        .NotOk Lp_TOC_Line_Has_Page_Number("Lesson 5 Place Value to 1,000" & vbCr)
    End With

    With Suite.Test("dot leaders in front of the number are a page number")
        .IsOk Lp_TOC_Line_Has_Page_Number("Lesson 2 Add and Subtract...143" & vbCr)
        .IsOk Lp_TOC_Line_Has_Page_Number("Lesson 2 Add and Subtract . . . 143" & vbCr)
        .IsOk Lp_TOC_Line_Has_Page_Number("Chapter One ....... 7" & vbCr)
    End With

    With Suite.Test("two dots are required, so a numbered sentence is not one")
        .NotOk Lp_TOC_Line_Has_Page_Number("Lesson 2. 100" & vbCr)
        .NotOk Lp_TOC_Line_Has_Page_Number("See Appendix B. 12" & vbCr)
    End With

    With Suite.Test("a roman page number behind leaders counts")
        .IsOk Lp_TOC_Line_Has_Page_Number("Foreword . . . xiv" & vbCr)
        .IsOk Lp_TOC_Line_Has_Page_Number("Foreword . . . XIV" & vbCr)
    End With

    With Suite.Test("a lettered page number behind leaders counts")
        ' G1 and E1 are the shapes the lone-number test accepts.
        .IsOk Lp_TOC_Line_Has_Page_Number("Glossary . . . G1" & vbCr)
    End With

    With Suite.Test("leaders with nothing after them are not a page number")
        .NotOk Lp_TOC_Line_Has_Page_Number("Lesson 2 Add and Subtract . . ." & vbCr)
        .NotOk Lp_TOC_Line_Has_Page_Number("Lesson 2 Add and Subtract.." & vbCr)
    End With

    With Suite.Test("a line with no paragraph mark is answered the same way")
        .IsOk Lp_TOC_Line_Has_Page_Number("Chapter One ... 7")
        .NotOk Lp_TOC_Line_Has_Page_Number("Chapter 12")
    End With

    With Suite.Test("trailing space or non-breaking space does not hide the number")
        .IsOk Lp_TOC_Line_Has_Page_Number("Chapter One ... 7  " & vbCr)
        .IsOk Lp_TOC_Line_Has_Page_Number("Chapter One ... 7" & ChrW(160) & vbCr)
    End With

    With Suite.Test("a non-breaking space between the leaders and the number is allowed")
        .IsOk Lp_TOC_Line_Has_Page_Number("Chapter One ..." & ChrW(160) & "7" & vbCr)
    End With

    With Suite.Test("an empty line has no page number")
        .NotOk Lp_TOC_Line_Has_Page_Number("")
        .NotOk Lp_TOC_Line_Has_Page_Number(vbCr)
    End With

    Set TestTocPageNumbers_Suite = Suite
End Function

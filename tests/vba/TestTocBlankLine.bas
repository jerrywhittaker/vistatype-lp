Attribute VB_Name = "TestTocBlankLine"
'@uses Lp_TOC_Line_Gets_Blank_Before
'@uses Lp_TOC_Line_Ends_In_Page_Number
'
' Which lines of a contents page get a blank line in front of them. Jerry, 9/21/2026: "the
' blank line goes before any line that has no associated page number."
'
' Until 3.0.476 only a wholly bold line got one. On "TOC Table 2 Formatting.docx" the bullet in
' front of each title is never bold, so none of its unnumbered reading titles did. The lines
' below are shaped like that file's, after Format the TOC has taken the bullets off.
'
' Version: 1.0  Date: 9/21/2026
'
Option Explicit

Public Function TestTocBlankLine_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Contents lines that get a blank line in front"

    With Suite.Test("an entry behind a non-breaking space ends in its page number")
        ' The file's own shape. Read with the non-breaking space as a space, or every entry in
        ' it would be given a blank line.
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("1.2" & ChrW(160) & "Perception Is Everything" & ChrW(160) & "6" & vbCr)
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("Unit 1 Opener" & ChrW(160) & "2" & vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before("1.2" & ChrW(160) & "Perception Is Everything" & ChrW(160) & "6" & vbCr)
    End With

    With Suite.Test("an entry behind an ordinary space ends in its page number")
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("Chapter 5 36" & vbCr)
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("Chapter 5 36")
        .NotOk Lp_TOC_Line_Gets_Blank_Before("Chapter 5 36" & vbCr)
    End With

    With Suite.Test("roman and lettered page numbers count")
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("Foreword xiv" & vbCr)
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("Foreword XIV" & vbCr)
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("Glossary G1" & vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before("Foreword" & ChrW(160) & "xiv" & vbCr)
    End With

    With Suite.Test("a reading title with no page number gets the blank line")
        ' The fault. About 98 of these in the file.
        .NotOk Lp_TOC_Line_Ends_In_Page_Number("Poetry:" & ChrW(160) & """The New Colossus,"" Emma Lazarus" & vbCr)
        .IsOk Lp_TOC_Line_Gets_Blank_Before("Poetry:" & ChrW(160) & """The New Colossus,"" Emma Lazarus" & vbCr)
        .IsOk Lp_TOC_Line_Gets_Blank_Before("Short Story:" & ChrW(160) & """The Gift of the Magi""" & vbCr)
    End With

    With Suite.Test("other lines with no page number get it too")
        .IsOk Lp_TOC_Line_Gets_Blank_Before("ACTIVITY Unit 1:" & vbCr)
        .IsOk Lp_TOC_Line_Gets_Blank_Before("Resources" & vbCr)
        .IsOk Lp_TOC_Line_Gets_Blank_Before("Table of Contents" & vbCr)
        .IsOk Lp_TOC_Line_Gets_Blank_Before("Resources")
    End With

    With Suite.Test("a $pg line does not")
        .NotOk Lp_TOC_Line_Gets_Blank_Before("$pg" & vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before("$pg 12" & vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before("$pg vii" & vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before(" $pg 12" & vbCr)
    End With

    With Suite.Test("an empty line does not")
        ' The file has one empty bulleted paragraph; with its bullet gone it is this.
        .NotOk Lp_TOC_Line_Gets_Blank_Before("")
        .NotOk Lp_TOC_Line_Gets_Blank_Before(vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before("   " & vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before("" & ChrW(160) & "" & vbCr)
        .NotOk Lp_TOC_Line_Ends_In_Page_Number("")
        .NotOk Lp_TOC_Line_Ends_In_Page_Number(vbCr)
    End With

    With Suite.Test("a title ending in a roman-letter word reads as numbered, as it always has")
        ' Not a choice made here: the macro's pattern runs with IgnoreCase on, and has since
        ' 9/1/2025. Pinned so the blank-line decision and the tab can never disagree.
        .IsOk Lp_TOC_Line_Ends_In_Page_Number("Part II" & vbCr)
        .NotOk Lp_TOC_Line_Gets_Blank_Before("Part II" & vbCr)
    End With

    Set TestTocBlankLine_Suite = Suite
End Function

Attribute VB_Name = "TestTocLinkEntry"
'@uses Lp_TOC_Line_Is_Link_Entry
'
' A contents line whose page number is BLUE is an entry, however bold it is. Jerry, 9/21/2026,
' on "TOC Table 2 Formatting.docx": 77 wholly bold lines with blue, underlined page numbers were
' read as section headings, so they got no tab and no TOC 1. The lines below are that file's.
'
' The color is handed in as a number: reading it off a Range would hang this runner
' (tests/vba/README.md).
'
' Version: 1.0  Date: 9/21/2026
'
Option Explicit

Public Function TestTocLinkEntry_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Contents lines with a blue page number"

    With Suite.Test("a blue page number makes the line an entry")
        ' The fault. Bold from end to end, no leader, no tab - read as a heading until 3.0.478.
        .IsOk Lp_TOC_Line_Is_Link_Entry("1.2 Perception Is Everything 6" & vbCr, wdColorBlue)
        .IsOk Lp_TOC_Line_Is_Link_Entry("A Letter to the Student xxi" & vbCr, wdColorBlue)
        .IsOk Lp_TOC_Line_Is_Link_Entry("2.13 Money, Power, and Class in Pygmalion 193" & vbCr, wdColorBlue)
        .IsOk Lp_TOC_Line_Is_Link_Entry("1.1" & ChrW(160) & "Previewing the Unit" & ChrW(160) & "4" & vbCr, wdColorBlue)
    End With

    With Suite.Test("a bold heading in black that ends in a number is not")
        ' "Table of Contents with Bold.docx", 9/7/2026: these stay headings.
        .NotOk Lp_TOC_Line_Is_Link_Entry("Section 1" & vbCr, wdColorAutomatic)
        .NotOk Lp_TOC_Line_Is_Link_Entry("Section 2" & vbCr, wdColorBlack)
    End With

    With Suite.Test("blue with no page number is not")
        .NotOk Lp_TOC_Line_Is_Link_Entry("ACTIVITY Unit 1: Perception Is Everything" & vbCr, wdColorBlue)
        .NotOk Lp_TOC_Line_Is_Link_Entry("", wdColorBlue)
        .NotOk Lp_TOC_Line_Is_Link_Entry(vbCr, wdColorBlue)
    End With

    With Suite.Test("another color is not")
        ' The same file marks its $pg lines in red.
        .NotOk Lp_TOC_Line_Is_Link_Entry("$pg vi" & vbCr, wdColorRed)
        .NotOk Lp_TOC_Line_Is_Link_Entry("1.2 Perception Is Everything 6" & vbCr, wdColorRed)
    End With

    Set TestTocLinkEntry_Suite = Suite
End Function

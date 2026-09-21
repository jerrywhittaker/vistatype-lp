Attribute VB_Name = "TestTocShortTitle"
'@uses Lp_TOC_Tab_Stops_At_Indent
'
' A contents entry whose title ends before the TOC style's hanging indent sends its tab only as
' far as the indent, so the page number shows with no dot leader. Jerry, 9/21/2026, on "TOC Table
' 1 Formatting.docx", build 3.0.480: "Index<tab>18" in TOC 1 showed as "Index 18". The cure is a
' second tab, and this is the decision of when to add one.
'
' The positions are handed in as numbers, in points from the left edge of the text: reading them
' off a Range would hang this runner (tests/vba/README.md). The TOC 1 numbers are the ones
' MEASURED on Jerry's file on the build box, 9/21/2026 - indent 54, hanging by 54. Where in the
' TOC the line sits (last line or mid-TOC) cannot change these numbers, and both were proved on
' the build box against his file; see docs/Reported-Errors.md.
'
' Version: 1.0  Date: 9/21/2026
'
Option Explicit

Public Function TestTocShortTitle_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Contents entries whose title is shorter than the hanging indent"

    With Suite.Test("a short title's tab stops at the indent and gets a second tab")
        ' "Index<tab>18", measured: the tab starts at 51.05, the 18 is drawn at 54.4.
        .IsOk Lp_TOC_Tab_Stops_At_Indent(51.05, 54.4, 54, -54, True, "1")
        ' "Quiz" and "Map 5" end sooner still; the number is at the indent either way.
        .IsOk Lp_TOC_Tab_Stops_At_Indent(38, 54, 54, -54, True, "7")
        .IsOk Lp_TOC_Tab_Stops_At_Indent(46.5, 54.2, 54, -54, True, "1")
        ' TOC 2 hangs by the same 0.75" from a 1" indent: its first line starts at 18 pt.
        .IsOk Lp_TOC_Tab_Stops_At_Indent(60, 72.3, 72, -54, True, "4")
    End With

    With Suite.Test("a title long enough to reach the leader gets nothing")
        ' "Glossary<tab>G1" and "Chapter 1<tab>1", measured the same day.
        .NotOk Lp_TOC_Tab_Stops_At_Indent(77, 425.3, 54, -54, True, "G")
        .NotOk Lp_TOC_Tab_Stops_At_Indent(88.75, 438.7, 54, -54, True, "1")
        ' "Contributors to this book<tab>iv"
        .NotOk Lp_TOC_Tab_Stops_At_Indent(227.7, 434.5, 54, -54, True, "i")
    End With

    With Suite.Test("a title that wraps onto a second line gets nothing")
        ' Its tab is on the second line, which starts AT the indent, so it cannot start short.
        .NotOk Lp_TOC_Tab_Stops_At_Indent(54, 440, 54, -54, True, "1")
        .NotOk Lp_TOC_Tab_Stops_At_Indent(70, 440, 54, -54, True, "1")
        ' The number pushed onto the next line on its own starts at the indent too.
        .NotOk Lp_TOC_Tab_Stops_At_Indent(51, 54, 54, -54, False, "1")
    End With

    With Suite.Test("a second tab already there is not stacked")
        .NotOk Lp_TOC_Tab_Stops_At_Indent(51.05, 54.4, 54, -54, True, vbTab)
        .NotOk Lp_TOC_Tab_Stops_At_Indent(51.05, 54.4, 54, -54, True, vbCr)
        .NotOk Lp_TOC_Tab_Stops_At_Indent(51.05, 54.4, 54, -54, True, "")
    End With

    With Suite.Test("no layout, or no hanging indent, never adds a tab")
        ' Information answers -1 when Word has no layout to read.
        .NotOk Lp_TOC_Tab_Stops_At_Indent(-1, -1, 54, -54, True, "1")
        .NotOk Lp_TOC_Tab_Stops_At_Indent(51.05, -1, 54, -54, True, "1")
        ' Not hanging - TOC 6 to TOC 9 in the template, or a style the transcriber changed.
        .NotOk Lp_TOC_Tab_Stops_At_Indent(51.05, 54.4, 54, 0, True, "1")
        .NotOk Lp_TOC_Tab_Stops_At_Indent(20, 22, 0, 0, True, "1")
    End With

    Set TestTocShortTitle_Suite = Suite
End Function

Attribute VB_Name = "VtRunTests"
'
' The entry point. Run-VbaTests.ps1 injects this module, substitutes the result path below,
' and calls VtRunAllTests.
'
' The path is a placeholder rather than something built in VBA on purpose: Environ$("APPDATA")
' inside an automation Word points at the SYSTEM profile, not at the user's, which cost a run
' on 9/1/2026. Every path is built in PowerShell and substituted into the text.
'
' Nothing here may raise. An unhandled VBA error puts up a modal dialog behind an invisible
' Word and hangs the SSH command until it times out, leaving a Word to kill and lock files to
' sweep. The trap writes RESULT ERROR to the file instead, so the run always ends and always
' says why.
'
' Version: 1.0  Date: 9/20/2026
'
Option Explicit

Private Const VT_RESULT_PATH As String = "__VT_RESULT_PATH__"

Public Sub VtRunAllTests()
    Dim Reporter As New VtFileReporter
    Dim Suite As TestSuite

    On Error GoTo Failed

    Reporter.StartFile VT_RESULT_PATH

    Reporter.Trace "building Roman page numbers"
    Set Suite = TestRomanNumerals_Suite()
    Reporter.ListenTo Suite

    Reporter.Trace "building Reference page numbers"
    Set Suite = TestPageNumbers_Suite()
    Reporter.ListenTo Suite

    Reporter.Trace "building Contents lines and their page numbers"
    Set Suite = TestTocPageNumbers_Suite()
    Reporter.ListenTo Suite

    Reporter.Trace "building Contents lines that get a blank line in front"
    Set Suite = TestTocBlankLine_Suite()
    Reporter.ListenTo Suite

    Reporter.Trace "building Contents lines with a blue page number"
    Set Suite = TestTocLinkEntry_Suite()
    Reporter.ListenTo Suite

    Reporter.Trace "building Contents entries shorter than the hanging indent"
    Set Suite = TestTocShortTitle_Suite()
    Reporter.ListenTo Suite

    Reporter.Trace "building Resizing one picture wherever it repeats"
    Set Suite = TestSamePicture_Suite()
    Reporter.ListenTo Suite

    Reporter.Trace "building Shared text helpers"
    Set Suite = TestTextHelpers_Suite()
    Reporter.ListenTo Suite

    Reporter.Finish
    Exit Sub

Failed:
    Reporter.Abort "VBA error " & Err.Number & " - " & Err.Description
End Sub

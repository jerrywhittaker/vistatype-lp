Attribute VB_Name = "TestLocatePicture"
'@uses Lp_Locp_Step
'@uses Lp_Locp_Label
'
' Locate This Picture (404), Jerry, 11/1/2026: which copy Next and Previous go to, and the line
' the box shows. Numbers in and out, because anything reaching a Range hangs this runner
' (tests/vba/README.md).
'
' Version: 1.0  Date: 11/1/2026
'
Option Explicit

Public Function TestLocatePicture_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Locate This Picture steps through the copies"

    With Suite.Test("Next goes to the copy after, and round from the last to the first")
        .IsEqual Lp_Locp_Step(1, 3, True), 2
        .IsEqual Lp_Locp_Step(2, 3, True), 3
        .IsEqual Lp_Locp_Step(3, 3, True), 1
    End With

    With Suite.Test("Previous goes to the copy before, and round from the first to the last")
        .IsEqual Lp_Locp_Step(3, 3, False), 2
        .IsEqual Lp_Locp_Step(2, 3, False), 1
        .IsEqual Lp_Locp_Step(1, 3, False), 3
    End With

    With Suite.Test("with nothing on screen yet, Next is the first and Previous the last")
        .IsEqual Lp_Locp_Step(0, 3, True), 1
        .IsEqual Lp_Locp_Step(0, 3, False), 3
    End With

    With Suite.Test("one copy: every step stays on it")
        .IsEqual Lp_Locp_Step(1, 1, True), 1
        .IsEqual Lp_Locp_Step(1, 1, False), 1
    End With

    With Suite.Test("no copies: nowhere to go")
        .IsEqual Lp_Locp_Step(0, 0, True), 0
        .IsEqual Lp_Locp_Step(1, 0, False), 0
    End With

    With Suite.Test("the box says which copy of how many")
        .IsEqual Lp_Locp_Label(2, 5), "Picture 2 of 5"
        .IsEqual Lp_Locp_Label(1, 1200), "Picture 1 of 1,200"
    End With

    Set TestLocatePicture_Suite = Suite
End Function

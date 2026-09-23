Attribute VB_Name = "TestModelessBoxes"
'@uses Sh_Box_Remove
'@uses Sh_Box_Push
'@uses Sh_Box_Top
'@uses Lp_Box_Sel_Is_A_Spot
'
' The list of boxes that stay open, which decides who has F6 and who watches the cursor - Jerry's
' rule, 9/23/2026 (docs/UI-Conventions.md). The newest box up is the one; closing it hands both to
' the box opened before. Strings in and out, because anything reaching Word hangs this runner.
'
' Version: 1.0  Date: 9/23/2026
'
Option Explicit

Public Function TestModelessBoxes_Suite() As TestSuite
    Dim Suite As New TestSuite
    Dim s As String
    Suite.Description = "Boxes that stay open share F6"

    With Suite.Test("no box up: nobody has F6")
        .IsEqual Sh_Box_Top(""), ""
    End With

    With Suite.Test("the newest box has F6")
        s = Sh_Box_Push("", "tocb")
        .IsEqual Sh_Box_Top(s), "tocb"
        s = Sh_Box_Push(s, "rst")
        .IsEqual Sh_Box_Top(s), "rst"
    End With

    With Suite.Test("closing the newest hands F6 to the one before")
        s = Sh_Box_Push(Sh_Box_Push("", "tocb"), "rst")
        s = Sh_Box_Remove(s, "rst")
        .IsEqual Sh_Box_Top(s), "tocb"
    End With

    With Suite.Test("closing an older box leaves the newest with F6")
        s = Sh_Box_Push(Sh_Box_Push(Sh_Box_Push("", "pg"), "tocb"), "rst")
        s = Sh_Box_Remove(s, "tocb")
        .IsEqual Sh_Box_Top(s), "rst"
        s = Sh_Box_Remove(s, "rst")
        .IsEqual Sh_Box_Top(s), "pg"
    End With

    With Suite.Test("the last box closing leaves nothing, so F6 goes back to Word")
        s = Sh_Box_Remove(Sh_Box_Push("", "pg"), "pg")
        .IsEqual s, ""
        .IsEqual Sh_Box_Top(s), ""
    End With

    With Suite.Test("a box opened again moves to the top, never listed twice")
        s = Sh_Box_Push(Sh_Box_Push(Sh_Box_Push("", "tocb"), "rst"), "tocb")
        .IsEqual Sh_Box_Top(s), "tocb"
        s = Sh_Box_Remove(s, "tocb")
        .IsEqual Sh_Box_Top(s), "rst"
        s = Sh_Box_Remove(s, "rst")
        .IsEqual s, ""
    End With

    With Suite.Test("closing a box that is not up changes nothing")
        s = Sh_Box_Push("", "tocb")
        .IsEqual Sh_Box_Remove(s, "rst"), s
        .IsEqual Sh_Box_Remove("", "rst"), ""
    End With

    ' --- what the out-of-range messages (389, 390) are about
    With Suite.Test("a placed cursor or a selected picture is a spot")
        .IsOk Lp_Box_Sel_Is_A_Spot(wdSelectionIP)
        .IsOk Lp_Box_Sel_Is_A_Spot(wdSelectionInlineShape)
        .IsOk Lp_Box_Sel_Is_A_Spot(wdSelectionShape)
    End With

    With Suite.Test("a stretch of text being selected is not")
        .NotOk Lp_Box_Sel_Is_A_Spot(wdSelectionNormal)
        .NotOk Lp_Box_Sel_Is_A_Spot(wdSelectionColumn)
    End With

    Set TestModelessBoxes_Suite = Suite
End Function

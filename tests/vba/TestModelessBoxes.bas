Attribute VB_Name = "TestModelessBoxes"
'@uses Sh_Box_Remove
'@uses Sh_Box_Push
'@uses Sh_Box_Top
'@uses Sh_Box_AltKeys
'@uses Sh_Box_AltName
'@uses Lp_Box_Sel_Is_A_Spot
'
' The list of boxes that stay open, which decides who has F6 and who watches the cursor - Jerry's
' rule, 9/23/2026 (docs/UI-Conventions.md). The newest box up is the one; closing it hands both to
' the box opened before. Strings in and out, because anything reaching Word hangs this runner.
'
' Version: 1.2  Date: 10/30/2026 - the other five boxes' Alt keys, and Sh_Box_AltName for the
'                                two $pg menus
' Version: 1.1  Date: 10/30/2026 - the Alt keys each box binds in Word while it is the newest
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

    ' --- the box's own Alt keys, bound in Word while it is the newest (10/30/2026). Which letters
    ' are on each form is held by tests/test_box_alt_keys_in_book.py; this is the function itself.
    With Suite.Test("357 and 375 bind their own Alt keys")
        .IsEqual Sh_Box_AltKeys("fil"), "1234567890FD"
        .IsEqual Sh_Box_AltKeys("rst"), "AUDFPLCV"
    End With

    With Suite.Test("the TOC box, 348, the $pg menus and 402 bind their own Alt keys")
        .IsEqual Sh_Box_AltKeys("tocb"), "CAROD"
        .IsEqual Sh_Box_AltKeys("hv"), "ODSALT"
        .IsEqual Sh_Box_AltKeys("pg2"), "LDW"
        .IsEqual Sh_Box_AltKeys("pg4"), "RDW"
        .IsEqual Sh_Box_AltKeys("msg"), "OC"
    End With

    With Suite.Test("no box, or a $pg validation with no menu up, binds no Alt keys")
        .IsEqual Sh_Box_AltKeys(""), ""
        .IsEqual Sh_Box_AltKeys("pg0"), ""
        .IsEqual Sh_Box_AltKeys("pg"), ""
    End With

    With Suite.Test("the $pg menus' keys follow the menu on screen; other boxes keep their name")
        .IsEqual Sh_Box_AltName("pg", 2), "pg2"
        .IsEqual Sh_Box_AltName("pg", 4), "pg4"
        .IsEqual Sh_Box_AltName("pg", 0), "pg0"
        .IsEqual Sh_Box_AltName("tocb", 2), "tocb"
        .IsEqual Sh_Box_AltName("msg", 4), "msg"
        .IsEqual Sh_Box_AltName("", 0), ""
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

Attribute VB_Name = "TestResizePictures"
'@uses Lp_Fit_Factor
'@uses Lp_Fit_Right_Room
'
' Resize Selected Pictures (350), Jerry, 10/14/2026: no picture may end up past any margin, and
' one that would is shrunk to fit, keeping its shape. Lp_Fit_Factor is that decision, handed
' numbers rather than a picture, because anything reaching a picture hangs this runner
' (tests/vba/README.md). The margins it is fitted to, and the shrinking itself, are Word's and
' can only be tried in Word.
'
' Version: 1.1  Date: 10/14/2026 - the room to the right margin from where the picture starts
' Version: 1.0  Date: 10/14/2026
'
Option Explicit

Public Function TestResizePictures_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Resized pictures stay inside the margins"

    With Suite.Test("a picture that fits is left the size it is")
        .IsEqual Lp_Fit_Factor(200, 100, 468, 648), 1
        ' exactly the width, and exactly the height, still fit
        .IsEqual Lp_Fit_Factor(468, 100, 468, 648), 1
        .IsEqual Lp_Fit_Factor(100, 648, 468, 648), 1
    End With

    With Suite.Test("too wide: shrunk to the width between the margins")
        .IsEqual Lp_Fit_Factor(800, 100, 400, 648), 0.5
    End With

    With Suite.Test("too tall: shrunk to the height between the margins")
        .IsEqual Lp_Fit_Factor(100, 800, 468, 600), 0.75
    End With

    With Suite.Test("too wide AND too tall: the side furthest over decides")
        ' width wants 0.5, height wants 0.75 - 0.5 makes both fit
        .IsEqual Lp_Fit_Factor(800, 800, 400, 600), 0.5
        ' and the other way round
        .IsEqual Lp_Fit_Factor(800, 800, 600, 400), 0.5
    End With

    With Suite.Test("only the height is over, but the width is close: still the height decides")
        .IsEqual Lp_Fit_Factor(460, 1296, 468, 648), 0.5
    End With

    With Suite.Test("a size of zero or less changes nothing")
        .IsEqual Lp_Fit_Factor(0, 100, 468, 648), 1
        .IsEqual Lp_Fit_Factor(100, 0, 468, 648), 1
        .IsEqual Lp_Fit_Factor(800, 100, 0, 648), 1
        .IsEqual Lp_Fit_Factor(800, 100, 468, -10), 1
    End With

    ' --- the room to the right margin. Letter (612), margins 72, gutter 36: the positions are the
    ' ones measured on the build box, 10/14/2026, for a 72 pt picture.
    With Suite.Test("plain page: a picture at the left margin has the whole text width")
        ' right edge at 144, so it starts at 72; the margin is at 540
        .IsEqual Lp_Fit_Right_Room(144, 72, 612, 72, 72, 0, False, False, False), 468
    End With

    With Suite.Test("indented, after text, or in a second column: less room")
        ' indented 72 - right edge 216
        .IsEqual Lp_Fit_Right_Room(216, 72, 612, 72, 72, 0, False, False, False), 396
        ' table column 2 - right edge 360
        .IsEqual Lp_Fit_Right_Room(360, 72, 612, 72, 72, 0, False, False, False), 252
    End With

    With Suite.Test("mirror margins: the inside margin and gutter are on the right of an even page")
        ' odd page: text starts at 108 (72 + gutter 36), right margin at 540
        .IsEqual Lp_Fit_Right_Room(180, 72, 612, 72, 72, 36, False, True, False), 432
        ' even page: text starts at 72, right margin at 612 - 72 - 36 = 504
        .IsEqual Lp_Fit_Right_Room(144, 72, 612, 72, 72, 36, False, True, True), 432
    End With

    With Suite.Test("a gutter at the top takes nothing off the right")
        .IsEqual Lp_Fit_Right_Room(144, 72, 612, 72, 72, 36, True, True, True), 468
        .IsEqual Lp_Fit_Right_Room(144, 72, 612, 72, 72, 36, True, False, False), 468
    End With

    With Suite.Test("no position from Word: -1, and the caller keeps the width between the margins")
        .IsEqual Lp_Fit_Right_Room(-1, 72, 612, 72, 72, 0, False, False, False), -1
    End With

    Set TestResizePictures_Suite = Suite
End Function

Attribute VB_Name = "TestResizePictures"
'@uses Lp_Fit_Factor
'
' Resize Selected Pictures (350), Jerry, 10/14/2026: no picture may end up past any margin, and
' one that would is shrunk to fit, keeping its shape. Lp_Fit_Factor is that decision, handed
' numbers rather than a picture, because anything reaching a picture hangs this runner
' (tests/vba/README.md). The margins it is fitted to, and the shrinking itself, are Word's and
' can only be tried in Word.
'
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

    Set TestResizePictures_Suite = Suite
End Function

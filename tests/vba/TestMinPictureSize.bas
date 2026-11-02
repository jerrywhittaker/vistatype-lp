Attribute VB_Name = "TestMinPictureSize"
'@uses Lp_Min_Pic_Factor
'
' No picture smaller than the text, Jerry, 11/1/2026: the LP template attach enlarges a picture
' whose shorter side is under the Normal style's font size until it is that size, keeping its
' shape. Lp_Min_Pic_Factor is that decision, handed numbers rather than a picture, because
' anything reaching a picture hangs this runner (tests/vba/README.md).
'
' Version: 1.0  Date: 11/1/2026
'
Option Explicit

Public Function TestMinPictureSize_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "No picture smaller than the text"

    With Suite.Test("a dot is enlarged to the text size")
        .IsEqual Lp_Min_Pic_Factor(2, 2, 18), 9
    End With

    With Suite.Test("a sliver is enlarged by its shorter side")
        .IsEqual Lp_Min_Pic_Factor(2, 100, 18), 9
        .IsEqual Lp_Min_Pic_Factor(100, 2, 18), 9
    End With

    With Suite.Test("a picture big enough is left alone, never shrunk")
        .IsEqual Lp_Min_Pic_Factor(200, 100, 18), 1
        .IsEqual Lp_Min_Pic_Factor(18, 18, 18), 1
        .IsEqual Lp_Min_Pic_Factor(18, 400, 18), 1
    End With

    With Suite.Test("sizes of zero or less change nothing")
        .IsEqual Lp_Min_Pic_Factor(0, 10, 18), 1
        .IsEqual Lp_Min_Pic_Factor(10, 0, 18), 1
        .IsEqual Lp_Min_Pic_Factor(10, 10, 0), 1
    End With

    Set TestMinPictureSize_Suite = Suite
End Function

Attribute VB_Name = "TestErrorLog"
'@uses Sh_Log_Newest_First
'@uses Sh_One_Log_Line
'
' The error log, issue #24. A description ending in CR LF NUL split its entry, and the next
' rewrite left about 2,000 NULs in the file. MEASURED 10/7/2026 on a copy of the damaged log:
' reading it back with FileSystemObject's ReadAll returned junk, the Write raised after the file
' had been emptied, and the WHOLE log was lost. These run the real Sh_Log_Newest_First on a
' scratch file in %TEMP% - never on the real log.
'
' Version: 1.0  Date: 10/7/2026
'
Option Explicit

Private Function TestErrorLog_ReadFile(ByVal path As String) As String
    Dim f As Integer
    Dim s As String
    f = FreeFile
    Open path For Binary Access Read As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    TestErrorLog_ReadFile = s
End Function

Private Sub TestErrorLog_WriteFile(ByVal path As String, ByVal s As String)
    Dim f As Integer
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , s
    Close #f
End Sub

Public Function TestErrorLog_Suite() As TestSuite
    Dim Suite As New TestSuite
    Dim path As String
    Dim damaged As String
    Dim newLine As String
    Dim after As String
    Dim lines As Variant

    Suite.Description = "The error log"
    path = Environ$("TEMP") & "\VistaType-Errors-test24.log"

    ' The shape of the real damaged log: an entry split by CR LF NUL, then a run of NULs where
    ' the next entry's details were.
    damaged = "2026-10-06 20:05:22  v3.0.539  Dx_Manual_Tag_with_Dollar_pg  err -2147221040 ""OpenClipboard Failed" _
            & vbCrLf & String$(2000, vbNullChar) & """  Word 16.0  doc ""Document1""" & vbCrLf _
            & "2026-10-06 20:04:45  v3.0.539  Dx_Manual_Tag_with_Dollar_pg  err -2147221040 ""x""" & vbCrLf
    newLine = Sh_One_Log_Line("2026-10-07 12:00:00  v3.0.544  Test  err 5 ""OpenClipboard Failed" _
                              & vbCrLf & vbNullChar & """  step ""s""")

    With Suite.Test("a log already holding NULs is mended, not wiped")
        TestErrorLog_WriteFile path, damaged
        Sh_Log_Newest_First path, newLine
        after = TestErrorLog_ReadFile(path)
        .IsOk Len(after) > 0, "the log was emptied"
        .IsEqual InStr(after, vbNullChar), 0, "NULs left in the log"
        lines = Split(after, vbCrLf)
        .IsEqual lines(0), newLine, "the new entry is not the first line"
        .IsOk InStr(after, "2026-10-06 20:04:45") > 0, "an earlier entry was lost"
    End With

    With Suite.Test("the new entry is one line with its two-space separators")
        .IsEqual newLine, "2026-10-07 12:00:00  v3.0.544  Test  err 5 ""OpenClipboard Failed   ""  step ""s"""
    End With

    On Error Resume Next
    Kill path
    On Error GoTo 0

    Set TestErrorLog_Suite = Suite
End Function

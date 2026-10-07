Attribute VB_Name = "TestErrorLog"
'@uses Sh_Log_Newest_First
'@uses Sh_One_Log_Line
'@uses Sh_Read_Log_Text
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

' Reads the log back the way a person's editor would: UTF-16 if it starts with the Unicode
' byte-order mark, ANSI if not. Its own code, NOT the add-in's, so it cannot share a mistake.
Private Function TestErrorLog_ReadFile(ByVal path As String) As String
    Dim f As Integer
    Dim n As Long
    Dim b() As Byte
    Dim s As String
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    If n = 0 Then
        s = ""
    ElseIf n >= 2 And b(0) = &HFF And b(1) = &HFE Then
        s = b
        s = Mid$(s, 2)
    Else
        s = StrConv(b, vbUnicode)
    End If
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
    Dim f As Integer
    Dim bom(0 To 1) As Byte
    Dim uni() As Byte

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
        lines = Split(after & vbCrLf, vbCrLf)     ' never empty, so lines(0) is safe on an emptied log
        .IsEqual lines(0), newLine, "the new entry is not the first line"
        .IsOk InStr(after, "2026-10-06 20:04:45") > 0, "an earlier entry was lost"
    End With

    With Suite.Test("the new entry is one line with its two-space separators")
        .IsEqual newLine, "2026-10-07 12:00:00  v3.0.544  Test  err 5 ""OpenClipboard Failed   ""  step ""s"""
    End With

    ' Issue #25. The log was rewritten as ANSI text, emptied first, so a character ANSI cannot
    ' hold - a Greek document name - made the write fail with the file already empty, and every
    ' earlier entry was lost. The Greek is written as "?", so the log stays plain text.
    With Suite.Test("a Greek letter in an entry does not wipe the log")
        TestErrorLog_WriteFile path, "2026-10-06 20:04:45  v3.0.539  Older  err 5 ""x""" & vbCrLf
        newLine = "2026-10-07 12:00:00  v3.0.545  Test  err 5 ""x""  doc """ _
                & ChrW(&H3B1) & ChrW(&H3B2) & ChrW(&H3B3) & ".docx"""
        Sh_Log_Newest_First path, newLine
        after = TestErrorLog_ReadFile(path)
        .IsOk InStr(after, "2026-10-06 20:04:45") > 0, "the earlier entry was lost"
        lines = Split(after & vbCrLf, vbCrLf)     ' never empty, so lines(0) is safe on an emptied log
        ' What the Greek becomes is Windows' own ANSI conversion, which can write a look-alike
        ' rather than "?" - so the test asks only that the entry is there, whole and first.
        .IsOk Left$(lines(0), 53) = "2026-10-07 12:00:00  v3.0.545  Test  err 5 ""x""  doc """ _
              And Right$(lines(0), 6) = ".docx""", "the new entry is not the first line, got: " & lines(0)
        .IsEqual Len(lines(0)), Len(newLine), "the Greek did not come out one character for one, got: " & lines(0)
    End With

    With Suite.Test("an old ANSI log is read correctly before it is rewritten")
        TestErrorLog_WriteFile path, "2026-10-06 20:04:45  v3.0.539  Older  err 5 ""caf" & Chr$(233) & """" & vbCrLf
        Sh_Log_Newest_First path, "2026-10-07 12:00:00  v3.0.545  Test  err 5 ""y"""
        after = TestErrorLog_ReadFile(path)
        .IsOk InStr(after, "caf" & ChrW(233)) > 0, "the earlier entry's accented letter was mangled"
    End With

    ' 3.0.546 wrote the log as Unicode, and Jerry's Notepad showed it with the letters spaced out.
    ' A Unicode log must be read correctly and come back as plain ANSI.
    With Suite.Test("a Unicode log is read and written back as plain text")
        f = FreeFile
        If Len(Dir$(path)) > 0 Then Kill path
        Open path For Binary Access Write As #f
        bom(0) = &HFF
        bom(1) = &HFE
        Put #f, , bom
        uni = "2026-10-07 16:24:02  v3.0.546  Unicode  err 5 ""a""" & vbCrLf
        Put #f, , uni
        Close #f
        .IsOk Sh_Log_Newest_First(path, "2026-10-07 16:30:00  v3.0.547  Test  err 5 ""z"""), _
              "the log was not written"
        f = FreeFile
        Open path For Binary Access Read As #f
        Get #f, , bom
        Close #f
        .IsOk Not (bom(0) = &HFF And bom(1) = &HFE), "the log is still Unicode"
        after = TestErrorLog_ReadFile(path)
        .IsEqual InStr(after, vbNullChar), 0, "the log is not plain text"
        .IsOk InStr(after, "2026-10-07 16:24:02  v3.0.546  Unicode") > 0, "the Unicode entry was lost or mangled"
    End With

    ' Dialog 240 says "written to a log file" only when it was. Before #25 the caller tested
    ' Err.Number after the call, which an error handled inside the sub never reaches.
    With Suite.Test("the log writer says whether the line was written")
        .IsEqual Sh_Log_Newest_First(path, "2026-10-07 12:00:02  ok"), True
        .IsEqual Sh_Log_Newest_First(Environ$("TEMP") & "\no-such-folder-25\x.log", "a"), False, _
                 "a log that cannot be written reported as written"
        .IsOk Len(Dir$(path & ".new")) = 0, "the scratch file was left behind"
    End With

    On Error Resume Next
    Kill path
    On Error GoTo 0

    Set TestErrorLog_Suite = Suite
End Function

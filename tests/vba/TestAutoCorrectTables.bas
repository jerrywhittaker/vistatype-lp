Attribute VB_Name = "TestAutoCorrectTables"
'@uses Sh_AC_Pick_Table
'@uses Sh_AC_Differences
'@uses Sh_AutoCorrect_Switch
'@uses Sh_AC_Snapshot
'@uses Sh_AC_Apply
'@uses Sh_AC_Write
'@uses Sh_AC_Read
'@uses Sh_AC_Table_File
'@uses Sh_Setting_Read
'@uses Sh_Setting_Write
'@uses Sh_Settings_File
'
' The three "Replace text as you type" lists - ordinary documents, large print, braille - and
' issue #21: the record of which list is up said "DEF" while Word held the large print list,
' because Word stores its list only on a clean quit. The switch trusted the record, left the book's
' list standing in an ordinary document, and the next book wrote it over AutoCorrect-DEF.txt.
' Every ordinary-only entry was gone for good. Measured on the build box, 3.0.517, 9/29/2026.
'
' The first half is pure: which table a list is closest to. The second half runs the real switch
' against Word's real AutoCorrect list and the real list files, so it BACKS UP all three files
' and the record first and puts every one of them back - and puts Word's list back too, because
' the runner quits Word cleanly and Word then stores whatever list the test left in it.
'
' Version: 1.0  Date: 9/29/2026
'
Option Explicit

Private Const T_DEF_NAME As String = "vtqqdefonly"
Private Const T_LP_NAME As String = "vtqqlponly"

' "a=1|b=2" as a name-to-replacement list, compared the way Sh_AC_Snapshot compares.
Private Function T_List(ByVal spec As String) As Object
    Dim d As Object
    Dim parts As Variant
    Dim i As Long
    Dim eq As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    If Len(spec) > 0 Then
        parts = Split(spec, "|")
        For i = LBound(parts) To UBound(parts)
            eq = InStr(parts(i), "=")
            d.Item(Left$(parts(i), eq - 1)) = Mid$(parts(i), eq + 1)
        Next i
    End If
    Set T_List = d
End Function

Private Function T_Tables(ByVal defSpec As Variant, ByVal lpSpec As Variant, ByVal brlSpec As Variant) As Object
    ' Null leaves that table out, as if it had no file.
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If Not IsNull(defSpec) Then d.Add "DEF", T_List(defSpec)
    If Not IsNull(lpSpec) Then d.Add "LP", T_List(lpSpec)
    If Not IsNull(brlSpec) Then d.Add "BRL", T_List(brlSpec)
    Set T_Tables = d
End Function

Private Function T_Copy(ByVal src As Object) As Object
    Dim d As Object
    Dim k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For Each k In src.Keys
        d.Item(k) = src.Item(k)
    Next k
    Set T_Copy = d
End Function

Private Function T_Word_Has(ByVal nm As String) As Boolean
    Dim p As Object
    Dim r As Object
    If Sh_AC_Snapshot(p, r) Then T_Word_Has = p.Exists(nm)
End Function

Private Function T_File_Has(ByVal cfgType As String, ByVal nm As String) As Boolean
    Dim d As Object
    If Sh_AC_Read(cfgType, d) Then T_File_Has = d.Exists(nm)
End Function

Public Function TestAutoCorrectTables_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "The three AutoCorrect lists (issue #21)"

    ' Shapes used below: DEF has two ordinary-only entries (the fractions, say); LP and BRL are
    ' the same list without them, as both book configurations leave them.
    Const DEF_LIST As String = "teh=the|adn=and|1/2=HALF|1/4=QUARTER"
    Const BOOK_LIST As String = "teh=the|adn=and"

    With Suite.Test("entries two lists disagree on are counted")
        .IsEqual Sh_AC_Differences(T_List(BOOK_LIST), T_List(BOOK_LIST)), 0
        .IsEqual Sh_AC_Differences(T_List(DEF_LIST), T_List(BOOK_LIST)), 2
        .IsEqual Sh_AC_Differences(T_List(BOOK_LIST), T_List(DEF_LIST)), 2
        .IsEqual Sh_AC_Differences(T_List("teh=the|adn=and"), T_List("teh=THE|adn=and")), 1
        .IsEqual Sh_AC_Differences(T_List("teh=the"), T_List("TEH=the")), 0
    End With

    With Suite.Test("the fault: record says ordinary, Word holds the large print list")
        .IsEqual Sh_AC_Pick_Table("DEF", T_List(BOOK_LIST), T_Tables(DEF_LIST, BOOK_LIST, Null)), "LP"
        ' ...and when braille's list is the same, it cannot be told, so no file may be written.
        .IsEqual Sh_AC_Pick_Table("DEF", T_List(BOOK_LIST), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), ""
    End With

    With Suite.Test("the mirror: record says large print, Word holds the ordinary list")
        ' The leak through the wall, had the record been trusted.
        .IsEqual Sh_AC_Pick_Table("LP", T_List(DEF_LIST), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), "DEF"
        .IsEqual Sh_AC_Pick_Table("LP", T_List(DEF_LIST & "|3/4=MINE"), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), "DEF"
        .IsEqual Sh_AC_Pick_Table("BRL", T_List(DEF_LIST), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), "DEF"
    End With

    With Suite.Test("entries added last session still belong to the recorded list")
        .IsEqual Sh_AC_Pick_Table("DEF", T_List(DEF_LIST & "|3/4=MINE"), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), "DEF"
        .IsEqual Sh_AC_Pick_Table("DEF", T_List(DEF_LIST), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), "DEF"
    End With

    With Suite.Test("a tie between the two books goes to the record")
        .IsEqual Sh_AC_Pick_Table("LP", T_List(BOOK_LIST & "|bk=book"), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), "LP"
        .IsEqual Sh_AC_Pick_Table("BRL", T_List(BOOK_LIST), T_Tables(DEF_LIST, BOOK_LIST, BOOK_LIST)), "BRL"
    End With

    With Suite.Test("the first switch on a machine, with no list files yet")
        .IsEqual Sh_AC_Pick_Table("", T_List(DEF_LIST), T_Tables(Null, Null, Null)), "DEF"
        .IsEqual Sh_AC_Pick_Table("LP", T_List(DEF_LIST), T_Tables(Null, Null, Null)), "LP"
        .IsEqual Sh_AC_Pick_Table("", T_List(BOOK_LIST), T_Tables(DEF_LIST, BOOK_LIST, Null)), "LP"
    End With

    RunTheRealSwitch Suite

    Set TestAutoCorrectTables_Suite = Suite
End Function

' The measured fault, end to end, against Word's own list and the real files.
Private Sub RunTheRealSwitch(ByVal Suite As TestSuite)
    Dim fso As Object
    Dim cfg As Variant
    Dim path As String
    Dim hadFile(0 To 2) As Boolean
    Dim i As Long
    Dim origRecord As String
    Dim origPlain As Object
    Dim origRich As Object
    Dim defList As Object
    Dim lpList As Object
    Dim nowPlain As Object
    Dim nowRich As Object
    Dim runError As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    With Suite.Test("the list folder can be reached")
        .IsOk Len(Sh_AC_Table_File("DEF")) > 0, "Sh_AC_Table_File returned nothing"
    End With
    If Len(Sh_AC_Table_File("DEF")) = 0 Then Exit Sub
    If Not Sh_AC_Snapshot(origPlain, origRich) Then
        Suite.Test("Word's list can be read").Fail "Sh_AC_Snapshot failed"
        Exit Sub
    End If
    If origPlain.Exists(T_DEF_NAME) Or origPlain.Exists(T_LP_NAME) Then
        Suite.Test("Word's list is clean before the test").Fail "a test entry was left from an earlier run"
        Exit Sub
    End If

    ' A backup left by a run that was killed part way is the ONLY copy of that list. Copying over
    ' it would lose the real list for good, so stop and leave it for a person to put back.
    For Each cfg In Array("DEF", "LP", "BRL")
        If fso.FileExists(Sh_AC_Table_File(CStr(cfg)) & ".vttest") Then
            Suite.Test("no backup is left from an earlier run").Fail _
                Sh_AC_Table_File(CStr(cfg)) & ".vttest exists - put it back by hand first"
            Exit Sub
        End If
    Next cfg

    ' --- Back up the three files and the record.
    i = 0
    For Each cfg In Array("DEF", "LP", "BRL")
        path = Sh_AC_Table_File(CStr(cfg))
        hadFile(i) = fso.FileExists(path)
        If hadFile(i) Then fso.CopyFile path, path & ".vttest", True
        i = i + 1
    Next cfg
    origRecord = Sh_Setting_Read("AutoCorrectTables", "Loaded", "")

    On Error GoTo Failed

    ' --- The state issue #21 measured: a clean quit in a large print book stored the LP list,
    '     the next session recorded "DEF", and Word then ended without a proper quit.
    Set defList = T_Copy(origPlain)
    defList.Item(T_DEF_NAME) = "ORDINARY ONLY"
    Set lpList = T_Copy(origPlain)
    lpList.Item(T_LP_NAME) = "LARGE PRINT ONLY"
    Sh_AC_Write "DEF", defList
    Sh_AC_Write "LP", lpList
    If hadFile(2) Then fso.DeleteFile Sh_AC_Table_File("BRL"), True
    Sh_AC_Apply lpList, origPlain, origRich                 ' Word holds the LP list
    Sh_Setting_Write "AutoCorrectTables", "Loaded", "DEF"   ' and the record says ordinary

    ' An ordinary document comes to the front: the first switch of this Word session.
    Sh_AutoCorrect_Switch "DEF"

    With Suite.Test("an ordinary document gets the ordinary list back though the record was wrong")
        .IsOk T_Word_Has(T_DEF_NAME), "the ordinary-only entry is not in Word"
        .NotOk T_Word_Has(T_LP_NAME), "the large print entry is still in Word"
        .IsEqual Sh_Setting_Read("AutoCorrectTables", "Loaded", ""), "DEF"
    End With

    With Suite.Test("AutoCorrect-DEF.txt is not overwritten with the large print list")
        .IsOk T_File_Has("DEF", T_DEF_NAME), "AutoCorrect-DEF.txt lost the ordinary-only entry"
        .NotOk T_File_Has("DEF", T_LP_NAME), "AutoCorrect-DEF.txt took the large print entry"
        .IsOk T_File_Has("LP", T_LP_NAME), "AutoCorrect-LP.txt lost its own entry"
        .NotOk T_File_Has("LP", T_DEF_NAME), "AutoCorrect-LP.txt took the ordinary entry"
    End With

    ' Then a large print book: the old code wrote the LP list over AutoCorrect-DEF.txt HERE.
    Sh_AutoCorrect_Switch "LP"

    With Suite.Test("the next large print book keeps both files apart")
        .IsOk T_Word_Has(T_LP_NAME), "the large print entry is not in Word"
        .NotOk T_Word_Has(T_DEF_NAME), "the ordinary-only entry leaked into the large print list"
        .IsOk T_File_Has("DEF", T_DEF_NAME), "AutoCorrect-DEF.txt lost the ordinary-only entry"
        .NotOk T_File_Has("DEF", T_LP_NAME), "AutoCorrect-DEF.txt took the large print entry"
    End With

Restore:
    ' --- Everything back as it was, whatever happened above. Reached through Resume on an error,
    '     so the handler is no longer active and On Error Resume Next really does hold here.
    On Error Resume Next
    If Sh_AC_Snapshot(nowPlain, nowRich) Then Sh_AC_Apply origPlain, nowPlain, nowRich
    i = 0
    For Each cfg In Array("DEF", "LP", "BRL")
        path = Sh_AC_Table_File(CStr(cfg))
        If fso.FileExists(path) Then fso.DeleteFile path, True
        If fso.FileExists(path & ".new") Then fso.DeleteFile path & ".new", True
        If hadFile(i) Then fso.MoveFile path & ".vttest", path
        i = i + 1
    Next cfg
    If Len(origRecord) > 0 Then
        Sh_Setting_Write "AutoCorrectTables", "Loaded", origRecord
    Else
        ' Sh_Setting_Write refuses "" on purpose, because "" deletes the key - which is exactly
        ' what is wanted here: there was no record before the test.
        System.PrivateProfileString(Sh_Settings_File(), "AutoCorrectTables", "Loaded") = ""
    End If
    Err.Clear
    On Error GoTo 0

    If Len(runError) > 0 Then Suite.Test("the real switch ran without a VBA error").Fail runError

    With Suite.Test("Word's list and the list files are put back after the test")
        .NotOk T_Word_Has(T_DEF_NAME), "a test entry was left in Word"
        .NotOk T_Word_Has(T_LP_NAME), "a test entry was left in Word"
        .IsEqual Sh_Setting_Read("AutoCorrectTables", "Loaded", ""), origRecord
        i = 0
        For Each cfg In Array("DEF", "LP", "BRL")
            .IsEqual fso.FileExists(Sh_AC_Table_File(CStr(cfg))), hadFile(i), CStr(cfg) & " file presence"
            i = i + 1
        Next cfg
    End With
    Exit Sub

Failed:
    ' Nothing that could raise belongs in here: an error inside an active handler would leave
    ' this Sub without reaching the restore. It is reported after the restore instead.
    runError = "VBA error " & Err.Number & " - " & Err.Description
    Resume Restore
End Sub

Attribute VB_Name = "TestPgValCheck"
'@uses Sh_PgVal_Parse
'@uses Sh_PgVal_Check
'@uses Sh_PgVal_Mark
'@uses Sh_PgVal_Summary_Text
'@uses Sh_PgVal_Is_Note_Line
'
' The checks on the $pg tag list - Validate Reference Page Tags (367).
'
' Until 10/16/2026 the list was read entirely by eye. Now each tag is read and a short red line,
' starting with a pointer mark, goes in ahead of any entry that looks wrong. These are the rules
' agreed with Jerry, 10/16/2026: roman to arabic is not a problem, any other step backwards is; a
' letter comes only after a number; Duxbury codes are read past, and the number after a
' continuation code must be the end of the range. And, Jerry 10/17/2026: a range of roman
' numerals is a problem - v-vi, XII-XIV - while a single roman page is not.
'
' Found() below turns the answer into one line - "1=8 missing" means the warning "8 missing" goes
' ahead of the entry numbered 1, counting from 0 - so a test reads like the list it describes.
'
' Version: 1.1  Date: 10/17/2026 - roman numeral ranges
' Version: 1.0  Date: 10/16/2026
'
Option Explicit

' Every warning, in order, as "index=text"; the mark is taken off, and two warnings ahead of the
' same entry are joined with " + ".
Private Function Found(ParamArray tags() As Variant) As String
    Dim v As Variant
    Dim w As Variant
    Dim i As Long
    Dim s As String
    Dim t As String

    v = tags
    w = Sh_PgVal_Check(v)
    For i = LBound(w) To UBound(w)
        If Len(w(i)) > 0 Then
            t = Replace(Replace(w(i), Sh_PgVal_Mark() & " ", ""), vbCr, " + ")
            If Len(s) > 0 Then s = s & "; "
            s = s & i & "=" & t
        End If
    Next i
    Found = s
End Function

Public Function TestPgValCheck_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Checking the $pg tag list"

    Dim w As Variant
    Dim i As Long
    Dim lines As Variant
    Dim j As Long
    Dim bad As String

    With Suite.Test("a tag is read into kind, first, last, letters, prefix, problem, case")
        .IsEqual Sh_PgVal_Parse("$pg12"), "A|12|12|||||"
        .IsEqual Sh_PgVal_Parse("$pg 12 "), "A|12|12|||||", "spaces around"
        .IsEqual Sh_PgVal_Parse("$pg" & ChrW(160) & "12"), "A|12|12|||||", "non-breaking space"
        .IsEqual Sh_PgVal_Parse("$pg12B"), "A|12|12|B|B|||"
        .IsEqual Sh_PgVal_Parse("$pg12-14"), "A|12|14|||||"
        .IsEqual Sh_PgVal_Parse("$pg12" & ChrW(8211) & "14"), "A|12|14|||||", "en dash"
        .IsEqual Sh_PgVal_Parse("$pgxiv"), "R|14|14|||||l"
        .IsEqual Sh_PgVal_Parse("$pgXIV"), "R|14|14|||||"
        .IsEqual Sh_PgVal_Parse("$pgA12"), "P|12|12|||A||"
        .IsEqual Sh_PgVal_Parse("12" & vbCr), "A|12|12|||||", "no $pg, paragraph mark"
    End With

    With Suite.Test("Duxbury codes are read past")
        .IsEqual Sh_PgVal_Parse("$pg12-14[[*lec*]][[*i*]]14"), "A|12|14|||||", "continuation"
        .IsEqual Sh_PgVal_Parse("$pg12-14[[*lec*]][[*i*]]13"), "A|12|14||||mismatch|", "mismatch"
        .IsEqual Sh_PgVal_Parse("$pg[[*ii*]]v"), "R|5|5|||||l", "EBAE lower-case roman"
        .IsEqual Sh_PgVal_Parse("$pgv-vi[[*lec*]][[*i*]]vi"), "R|5|6||||romanrange|l", "roman continuation"
    End With

    With Suite.Test("a roman numeral range is read as one")
        .IsEqual Sh_PgVal_Parse("$pgv-vi"), "R|5|6||||romanrange|l", "plain"
        .IsEqual Sh_PgVal_Parse("$pgxii-xiv"), "R|12|14||||romanrange|l", "plain, longer"
        .IsEqual Sh_PgVal_Parse("$pgXII-XIV"), "R|12|14||||romanrange|", "upper case"
        .IsEqual Sh_PgVal_Parse("$pgv-v"), "R|5|5||||romanrange|l", "same page both sides"
        .IsEqual Sh_PgVal_Parse("$pg[[*ii*]]v-vi"), "R|5|6||||romanrange|l", "EBAE code in front"
        .IsEqual Sh_PgVal_Parse("$pg[[*ii*]]v-vi[[*lec*]][[*i*]]vi"), "R|5|6||||romanrange|l", _
                 "EBAE code in front and continuation behind"
        .IsEqual Sh_PgVal_Parse("$pgv-vi[[*lec*]][[*i*]]v"), "R|5|6||||mismatch romanrange|l", _
                 "and a mismatch"
        .IsEqual Sh_PgVal_Parse("$pgv"), "R|5|5|||||l", "a single roman page is not a range"
        .IsEqual Sh_PgVal_Parse("$pg[[*ii*]]v"), "R|5|5|||||l", "nor in EBAE"
        .IsEqual Sh_PgVal_Parse("$pgword-word"), "X|0|0||||notnum|", "hyphenated word"
    End With

    With Suite.Test("what is not a page number")
        .IsEqual Sh_PgVal_Parse("$pg"), "X|0|0||||notnum|", "empty"
        .IsEqual Sh_PgVal_Parse("$pgIntroduction"), "X|0|0||||notnum|", "a word"
        .IsEqual Sh_PgVal_Parse("$pgAB"), "X|0|0||||notnum|", "letters only"
        .IsEqual Sh_PgVal_Parse("$pg 40 Chapter Two"), "X|0|0||||notnum|", "text after"
        .IsEqual Sh_PgVal_Parse("$pg12BC"), "X|0|0||||notnum|", "two letters after"
        .IsEqual Sh_PgVal_Parse("$pg-13"), "X|0|0||||notnum|", "stray hyphen"
        .IsEqual Sh_PgVal_Parse("$pg12-xiv"), "X|0|0||||notnum|", "mixed range"
        .IsEqual Sh_PgVal_Parse("$pg12[[*lec*"), "X|0|0||||notnum|", "code never closed"
    End With

    With Suite.Test("a clean list says nothing")
        .IsEqual Found("1", "2", "3", "4"), ""
        .IsEqual Found("5"), "", "one tag"
        .IsEqual Found("3", "4"), "", "a book whose first tag is not 1"
    End With

    With Suite.Test("gap")
        .IsEqual Found("7", "9"), "1=8 missing"
        .IsEqual Found("216", "217", "220"), "2=218-219 missing"
        .IsEqual Found("10", "31"), "1=11-30 missing", "20 missing is still a gap"
    End With

    With Suite.Test("big jump")
        .IsEqual Found("12", "120"), "1=12 to 120: typo?"
        .IsEqual Found("10", "32"), "1=10 to 32: typo?", "21 missing is a jump"
    End With

    With Suite.Test("duplicate")
        .IsEqual Found("216", "217", "217", "218"), "2=217 twice"
        .IsEqual Found("11", "12-14", "13", "15"), "2=13 twice", "inside a range already seen"
    End With

    With Suite.Test("backwards")
        .IsEqual Found("20", "22", "21", "23"), "2=out of order", "21 is only late, not missing"
        .IsEqual Found("20", "21", "15", "22"), "2=out of order"
    End With

    With Suite.Test("roman to arabic is fine, arabic to roman is not")
        .IsEqual Found("i", "ii", "iii", "1", "2"), ""
        .IsEqual Found("1", "2", "iii"), "2=out of order"
        .IsEqual Found("i", "ii", "iv", "1"), "2=iii missing", "roman gap, in roman"
    End With

    With Suite.Test("a letter after the number")
        .IsEqual Found("12", "12B", "13"), ""
        .IsEqual Found("12", "12A", "12B", "13"), ""
        .IsEqual Found("12", "12B", "14"), "2=13 missing"
        .IsEqual Found("11", "12B", "13"), "1=12B out of place", "no 12 ahead of it"
        .IsEqual Found("12", "12A", "12C"), "2=12C out of place", "12B skipped"
    End With

    With Suite.Test("ranges")
        .IsEqual Found("11", "12-14", "15"), ""
        .IsEqual Found("24", "25-27", "28"), ""
        .IsEqual Found("11", "12-99", "100"), "1=range 12-99?", "wide"
        .IsEqual Found("11", "12-17", "18"), "1=range 12-17?", "six pages"
        .IsEqual Found("11", "12-16", "17"), "", "five pages is fine"
        .IsEqual Found("11", "14-12", "15"), "1=range 14-12? + 12-13 missing", "backwards"
    End With

    With Suite.Test("braille continuation")
        .IsEqual Found("11", "12-14[[*lec*]][[*i*]]14", "15"), ""
        .IsEqual Found("11", "12-14[[*lec*]][[*i*]]13", "15"), "1=range end mismatch"
    End With

    With Suite.Test("EBAE lower-case roman")
        .IsEqual Found("[[*ii*]]iv", "[[*ii*]]v", "[[*ii*]]vi", "1"), ""
        .IsEqual Found("[[*ii*]]iv", "[[*ii*]]vi"), "1=v missing"
    End With

    With Suite.Test("roman numeral range")
        .IsEqual Found("iv", "v-vi", "vii", "1"), "1=roman numeral range"
        .IsEqual Found("xi", "xii-xiv", "xv"), "1=roman numeral range"
        .IsEqual Found("XI", "XII-XIV", "XV"), "1=roman numeral range", "upper case"
        .IsEqual Found("iv", "v-vi[[*lec*]][[*i*]]vi", "vii"), "1=roman numeral range", "braille"
        .IsEqual Found("[[*ii*]]iv", "[[*ii*]]v-vi[[*lec*]][[*i*]]vi", "[[*ii*]]vii"), _
                 "1=roman numeral range", "EBAE"
        .IsEqual Found("i", "ii-x", "xi"), "1=roman numeral range", "said once, not as a wide range too"
        .IsEqual Found("iv", "v-vi[[*lec*]][[*i*]]v", "vii"), _
                 "1=range end mismatch + roman numeral range", "and a mismatch"
        .IsEqual Found("iv", "v", "vi", "1"), "", "single roman pages are fine"
        .IsEqual Found("1", "Self-Help", "2"), "1=not a page number", "a hyphenated word"
    End With

    With Suite.Test("not a page number")
        .IsEqual Found("1", "Introduction", "2"), "1=not a page number", "a word"
        .IsEqual Found("1", "", "2"), "1=not a page number", "empty"
        .IsEqual Found("1", "AB", "2"), "1=not a page number", "letters only"
        .IsEqual Found("39", "40 Chapter Two", "41"), "1=not a page number; 2=40 missing", "text after"
    End With

    With Suite.Test("a letter prefix is a stretch of its own")
        .IsEqual Found("10", "A1", "A2", "11"), ""
        .IsEqual Found("A1", "A2", "A4"), "2=A3 missing"
    End With

    With Suite.Test("the test book on the build box")
        .IsEqual Found("$pgi", "$pgii", "$pgiii", "$pgiv", "$pgv", _
                       "$pg1", "$pg2", "$pg3", "$pg4", "$pg5", "$pg6", "$pg7", "$pg9", "$pg10", "$pg11", _
                       "$pg12", "$pg12B", "$pg13", "$pg14", "$pg15", "$pg16", "$pg17", "$pg17", "$pg18", _
                       "$pg19", "$pg20", "$pg22", "$pg21", "$pg23", "$pg24", "$pg25-27", "$pg28", "$pg29", _
                       "$pg30", "$pg130", "$pgIntroduction", "$pg", "$pg 40 Chapter Two", _
                       "$pg140-142[[*lec*]][[*i*]]141"), _
                 "12=8 missing; 22=17 twice; 27=out of order; 34=30 to 130: typo?; " _
                 & "35=not a page number; 36=not a page number; 37=not a page number; " _
                 & "38=range end mismatch + 131-139 missing"
    End With

    With Suite.Test("every warning starts with the mark and fits a large print column")
        w = Sh_PgVal_Check(Array("1", "2", "2", "iii", "5", "100", "100B", "101-199", _
                                 "word", "1000-1002[[*lec*]][[*i*]]1001", "A1", "A9", "A10", _
                                 "123456", "123460", "123480", "123470", "xxxviii-xxxix"))
        bad = ""
        For i = LBound(w) To UBound(w)
            If Len(w(i)) > 0 Then
                lines = Split(w(i), vbCr)
                For j = LBound(lines) To UBound(lines)
                    If Left$(lines(j), 2) <> Sh_PgVal_Mark() & " " Or Len(lines(j)) > 23 Then
                        bad = bad & " [" & lines(j) & "]"
                    End If
                Next j
            End If
        Next i
        .IsEqual bad, "", "these warnings break the rule: {0}"
    End With

    With Suite.Test("the count at the top")
        .IsEqual Sh_PgVal_Summary_Text(0), "No problems found. Please still look over the list."
        .IsEqual Sh_PgVal_Summary_Text(1), "1 possible problem: look for " & Sh_PgVal_Mark()
        .IsEqual Sh_PgVal_Summary_Text(3), "3 possible problems: look for " & Sh_PgVal_Mark()
        .IsEqual AscW(Sh_PgVal_Mark()), 9658, "the mark is U+25BA"
    End With

    With Suite.Test("Locate and Return know a warning line and the count from an entry")
        .IsOk Sh_PgVal_Is_Note_Line(Sh_PgVal_Mark() & " 217 twice" & vbCr), "a warning"
        .IsOk Sh_PgVal_Is_Note_Line(Sh_PgVal_Summary_Text(0) & vbCr), "no problems"
        .IsOk Sh_PgVal_Is_Note_Line(Sh_PgVal_Summary_Text(4) & vbCr), "a count"
        .NotOk Sh_PgVal_Is_Note_Line("$pg12" & vbCr), "an entry"
        .NotOk Sh_PgVal_Is_Note_Line("$pg12-14[[*lec*]][[*i*]]14" & vbCr), "a braille entry"
        .NotOk Sh_PgVal_Is_Note_Line("$pg12 Chapter..." & vbCr), "a cut entry"
        .NotOk Sh_PgVal_Is_Note_Line(vbCr), "a blank line"
    End With

    Set TestPgValCheck_Suite = Suite
End Function

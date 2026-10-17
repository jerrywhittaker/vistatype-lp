Attribute VB_Name = "TestPageNumbers"
'@uses Lp_Pg_Tag_Number
'@uses Lp_Merged_Pg_Number
'@uses Dx_Pg_Text_Is_Range
'@uses Dx_Closed_Up_Pg_Range
'@uses Dx_Pg_Range_Continuation
'@uses Dx_Pg_Range_Side_Shape
'@uses Dx_Pg_Continuation_Code_Start
'
' Reference page numbers - the print book's page numbers, cited in the large print or braille
' edition so a reader can follow a class or a citation. These two decide what number a tag
' carries and what happens when two tags are merged.
'
' Version: 1.0  Date: 9/20/2026
'
Option Explicit

Public Function TestPageNumbers_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Reference page numbers"

    With Suite.Test("the number is read out of the tag text")
        .IsEqual Lp_Pg_Tag_Number("  $pg 12 "), "12"
        .IsEqual Lp_Pg_Tag_Number("$pg12"), "12"
        .IsEqual Lp_Pg_Tag_Number("some text $pg 7"), "7"
        .IsEqual Lp_Pg_Tag_Number("$pg xiv"), "xiv"
    End With

    With Suite.Test("the tag marker is matched whatever its case")
        .IsEqual Lp_Pg_Tag_Number("$PG 12"), "12"
        .IsEqual Lp_Pg_Tag_Number("$Pg 12"), "12"
    End With

    With Suite.Test("a non-breaking space around the number is not part of it")
        .IsEqual Lp_Pg_Tag_Number("$pg" & ChrW(160) & "12"), "12"
        .IsEqual Lp_Pg_Tag_Number("$pg 12" & ChrW(160)), "12"
    End With

    With Suite.Test("a tag with no number gives nothing")
        .IsEqual Lp_Pg_Tag_Number("$pg"), ""
        .IsEqual Lp_Pg_Tag_Number("$pg   "), ""
    End With

    With Suite.Test("text that is not a tag gives nothing")
        .IsEqual Lp_Pg_Tag_Number("no tag here"), ""
        .IsEqual Lp_Pg_Tag_Number(""), ""
    End With

    With Suite.Test("merging two plain numbers gives a range")
        .IsEqual Lp_Merged_Pg_Number("12", "14"), "12-14"
    End With

    With Suite.Test("an already hyphenated tag does not grow a second hyphen")
        .IsEqual Lp_Merged_Pg_Number("12-13", "14"), "12-14"
        .IsEqual Lp_Merged_Pg_Number("12", "13-14"), "12-14"
        .IsEqual Lp_Merged_Pg_Number("12-13", "14-15"), "12-15"
    End With

    With Suite.Test("two identical tags give one number, not a range of one")
        .IsEqual Lp_Merged_Pg_Number("12", "12"), "12"
        .IsEqual Lp_Merged_Pg_Number("xiv", "XIV"), "xiv"
    End With

    With Suite.Test("a hyphen at either extreme is a stray, not a range boundary")
        .IsEqual Lp_Merged_Pg_Number("-13", "14"), "-13-14"
        .IsEqual Lp_Merged_Pg_Number("12", "14-"), "12-14-"
    End With

    With Suite.Test("an en dash or an em dash counts as a hyphen")
        ' A converted book can arrive carrying any of them.
        .IsEqual Lp_Merged_Pg_Number("12" & ChrW(8211) & "13", "14"), "12-14"
        .IsEqual Lp_Merged_Pg_Number("12" & ChrW(8212) & "13", "14"), "12-14"
    End With

    With Suite.Test("an empty number gives nothing rather than a broken range")
        .IsEqual Lp_Merged_Pg_Number("", "14"), ""
        .IsEqual Lp_Merged_Pg_Number("12", ""), ""
    End With

    ' Braille AutoTag, 10/7/2026 (Jerry: no difference between large print and braille). Measured
    ' on 3.0.540: four ranges on consecutive lines were merged into 16-A29B, and a range with a
    ' space round its hyphen was never tagged.
    With Suite.Test("braille: a line holding a range is a range, and is left out of the merge")
        .IsOk Dx_Pg_Text_Is_Range("16-17")
        .IsOk Dx_Pg_Text_Is_Range("22B-23B")
        .IsOk Dx_Pg_Text_Is_Range("A27B-A29B")
        .IsOk Dx_Pg_Text_Is_Range("AA30BB-AA33BB")
        .IsOk Dx_Pg_Text_Is_Range("12" & ChrW(8211) & "13")
    End With

    With Suite.Test("braille: a single page number, or a stray hyphen at either end, is not a range")
        .NotOk Dx_Pg_Text_Is_Range("16")
        .NotOk Dx_Pg_Text_Is_Range("B13")
        .NotOk Dx_Pg_Text_Is_Range("-13")
        .NotOk Dx_Pg_Text_Is_Range("13-")
        .NotOk Dx_Pg_Text_Is_Range("")
    End With

    With Suite.Test("braille: spaces round the hyphen of a page range are closed up")
        .IsEqual Dx_Closed_Up_Pg_Range("16 - 17"), "16-17"
        .IsEqual Dx_Closed_Up_Pg_Range("22B -23B"), "22B-23B"
        .IsEqual Dx_Closed_Up_Pg_Range("A27B- A29B"), "A27B-A29B"
        .IsEqual Dx_Closed_Up_Pg_Range("AA30BB - AA33BB"), "AA30BB-AA33BB"
        .IsEqual Dx_Closed_Up_Pg_Range("v - vi"), "v-vi"
        .IsEqual Dx_Closed_Up_Pg_Range("16" & ChrW(160) & "-" & ChrW(160) & "17"), "16-17"
    End With

    With Suite.Test("braille: two or more hyphens in a page range become one")
        .IsEqual Dx_Closed_Up_Pg_Range("16--17"), "16-17"
        .IsEqual Dx_Closed_Up_Pg_Range("16 -- 17"), "16-17"
    End With

    With Suite.Test("braille: anything that is not a spaced page range comes back unchanged")
        .IsEqual Dx_Closed_Up_Pg_Range("16-17"), "16-17"
        .IsEqual Dx_Closed_Up_Pg_Range("16"), "16"
        .IsEqual Dx_Closed_Up_Pg_Range("well - known"), "well - known"
        .IsEqual Dx_Closed_Up_Pg_Range("Chapter 3 - The End"), "Chapter 3 - The End"
        .IsEqual Dx_Closed_Up_Pg_Range("- 13"), "- 13"
        .IsEqual Dx_Closed_Up_Pg_Range("16 -"), "16 -"
        .IsEqual Dx_Closed_Up_Pg_Range("16 " & ChrW(8211) & " 17"), "16 " & ChrW(8211) & " 17"
        .IsEqual Dx_Closed_Up_Pg_Range("12345678901234 - 15"), "12345678901234 - 15"
    End With

    ' Braille Manual Tag Ref Page, 10/17/2026 (Jerry): a range gets the continuation code exactly
    ' as AutoTag writes it.
    With Suite.Test("braille: a range gets the DBT continuation code AutoTag gives it")
        .IsEqual Dx_Pg_Range_Continuation("16-18"), "[[*lec*]][[*i*]]18"
        .IsEqual Dx_Pg_Range_Continuation("A27-A29"), "[[*lec*]][[*i*]]A29"
        .IsEqual Dx_Pg_Range_Continuation("22B-23B"), "[[*lec*]][[*i*]]23B"
        .IsEqual Dx_Pg_Range_Continuation("A27B-A29B"), "[[*lec*]][[*i*]]A29B"
        .IsEqual Dx_Pg_Range_Continuation("AA30BB-AA33BB"), "[[*lec*]][[*i*]]AA33BB"
        .IsEqual Dx_Pg_Range_Continuation("v-vi"), "[[*lec*]][[*i*]]vi"
    End With

    With Suite.Test("braille: what AutoTag does not code gets no continuation code")
        .IsEqual Dx_Pg_Range_Continuation("16"), ""
        .IsEqual Dx_Pg_Range_Continuation("22B"), ""
        .IsEqual Dx_Pg_Range_Continuation("22-23B"), ""
        .IsEqual Dx_Pg_Range_Continuation("16 - 18"), ""
        .IsEqual Dx_Pg_Range_Continuation("16--18"), ""
        .IsEqual Dx_Pg_Range_Continuation("16" & ChrW(8211) & "18"), ""
        .IsEqual Dx_Pg_Range_Continuation("-18"), ""
        .IsEqual Dx_Pg_Range_Continuation("16-"), ""
        .IsEqual Dx_Pg_Range_Continuation(""), ""
    End With

    With Suite.Test("braille: a line already carrying the code does not get it twice")
        .IsEqual Dx_Pg_Range_Continuation("16-18[[*lec*]][[*i*]]18"), ""
    End With

    ' Manual Tag takes an old code off before it retags, reading the line with its hidden text.
    With Suite.Test("braille: an old continuation code is found at the end of the line")
        .IsEqual Dx_Pg_Continuation_Code_Start("16-18[[*lec*]][[*i*]]18"), 6
        .IsEqual Dx_Pg_Continuation_Code_Start("$pg16-18[[*lec*]][[*i*]]18"), 9
        .IsEqual Dx_Pg_Continuation_Code_Start("AA30BB-AA33BB[[*lec*]][[*i*]]AA33BB"), 14
        .IsEqual Dx_Pg_Continuation_Code_Start("16-19[[*lec*]][[*i*]]18"), 6
    End With

    With Suite.Test("braille: no continuation code, or one that is not at the end, is left alone")
        .IsEqual Dx_Pg_Continuation_Code_Start("16-18"), 0
        .IsEqual Dx_Pg_Continuation_Code_Start("$pg[[*ii*]]v"), 0
        .IsEqual Dx_Pg_Continuation_Code_Start("16-18[[*lec*]][[*i*]]18[[*ii*]]"), 0
        .IsEqual Dx_Pg_Continuation_Code_Start(""), 0
    End With

    With Suite.Test("braille: the shape of one side of a range")
        .IsEqual Dx_Pg_Range_Side_Shape("16"), "D"
        .IsEqual Dx_Pg_Range_Side_Shape("22B"), "DL"
        .IsEqual Dx_Pg_Range_Side_Shape("A27B"), "LDL"
        .IsEqual Dx_Pg_Range_Side_Shape("AA30BB"), "LDL"
        .IsEqual Dx_Pg_Range_Side_Shape("xiv"), "L"
        .IsEqual Dx_Pg_Range_Side_Shape("1 2"), ""
        .IsEqual Dx_Pg_Range_Side_Shape(""), ""
    End With

    Set TestPageNumbers_Suite = Suite
End Function

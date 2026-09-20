Attribute VB_Name = "TestPageNumbers"
'@uses Lp_Pg_Tag_Number
'@uses Lp_Merged_Pg_Number
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

    Set TestPageNumbers_Suite = Suite
End Function

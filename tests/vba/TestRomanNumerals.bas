Attribute VB_Name = "TestRomanNumerals"
'@uses Sh_IsValidRomanNumeral
'@uses Sh_Number_To_Roman
'
' Roman page numbers.
'
' The reason this suite exists: until 8/26/2026 Sh_IsValidRomanNumeral compared against a
' hand-typed list of 36 spelled-out numerals, and the list ran I to X and then jumped to XVI.
' Measured against the first hundred it rejected 72 - including xi to xv and xxvi to xxix,
' which is exactly where roman front matter lives. Both AutoTag macros ask this function
' whether a paragraph is a roman page number, so those pages were left untagged in every book,
' braille and large print alike, with nothing to say so.
'
' Version: 1.0  Date: 9/20/2026
'
Option Explicit

Public Function TestRomanNumerals_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Roman page numbers"

    Dim i As Long
    Dim bad As String
    Dim numeral As String

    With Suite.Test("every one of the first hundred numerals is accepted")
        bad = ""
        For i = 1 To 100
            numeral = Sh_Number_To_Roman(i)
            If Not Sh_IsValidRomanNumeral(numeral) Then bad = bad & " " & i
        Next i
        .IsEqual bad, "", "rejected these numbers: {0}"
    End With

    With Suite.Test("the numerals the old hand-typed list rejected")
        .IsOk Sh_IsValidRomanNumeral("xi"), "xi"
        .IsOk Sh_IsValidRomanNumeral("xii"), "xii"
        .IsOk Sh_IsValidRomanNumeral("xiii"), "xiii"
        .IsOk Sh_IsValidRomanNumeral("xiv"), "xiv"
        .IsOk Sh_IsValidRomanNumeral("xv"), "xv"
        .IsOk Sh_IsValidRomanNumeral("xxvi"), "xxvi"
        .IsOk Sh_IsValidRomanNumeral("xxvii"), "xxvii"
        .IsOk Sh_IsValidRomanNumeral("xxviii"), "xxviii"
        .IsOk Sh_IsValidRomanNumeral("xxix"), "xxix"
    End With

    With Suite.Test("case does not matter")
        .IsOk Sh_IsValidRomanNumeral("XIV")
        .IsOk Sh_IsValidRomanNumeral("xiv")
        .IsOk Sh_IsValidRomanNumeral("XiV")
    End With

    With Suite.Test("a paragraph mark or cell marker on the end does not matter")
        .IsOk Sh_IsValidRomanNumeral("xiv" & vbCr), "trailing paragraph mark"
        .IsOk Sh_IsValidRomanNumeral("xiv" & Chr$(7)), "trailing cell marker"
        .IsOk Sh_IsValidRomanNumeral(vbCr & "xiv" & vbLf), "both, either side"
        .IsOk Sh_IsValidRomanNumeral("  xiv  "), "surrounding spaces"
    End With

    With Suite.Test("a malformed numeral is refused")
        .NotOk Sh_IsValidRomanNumeral("IIII"), "IIII"
        .NotOk Sh_IsValidRomanNumeral("VX"), "VX"
        .NotOk Sh_IsValidRomanNumeral("IC"), "IC"
        .NotOk Sh_IsValidRomanNumeral("XXXX"), "XXXX"
        .NotOk Sh_IsValidRomanNumeral("VV"), "VV"
    End With

    With Suite.Test("something that is not a numeral at all is refused")
        .NotOk Sh_IsValidRomanNumeral(""), "empty"
        .NotOk Sh_IsValidRomanNumeral("   "), "spaces only"
        .NotOk Sh_IsValidRomanNumeral("12"), "an arabic number - Jerry typed this one"
        .NotOk Sh_IsValidRomanNumeral("Chapter"), "a word"
        .NotOk Sh_IsValidRomanNumeral("x i v"), "spaced out"
        .NotOk Sh_IsValidRomanNumeral("xiv2"), "a numeral with a digit stuck on"
    End With

    Set TestRomanNumerals_Suite = Suite
End Function

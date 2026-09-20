Attribute VB_Name = "TestTextHelpers"
'@uses Sh_Trailing_Para_Marks
'@uses Sh_AC_Escape
'@uses Sh_AC_Unescape
'@uses Lp_Marker_Value
'@uses Lp_Roman_Value
'
' The small shared helpers. Nothing here touches a document, which is exactly why they can be
' tested at all.
'
' Version: 1.0  Date: 9/20/2026
'
Option Explicit

Public Function TestTextHelpers_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Shared text helpers"

    ' --- Sh_Trailing_Para_Marks: half of the rule that the marks travel WITH the text and
    '     come back in the same number.
    With Suite.Test("paragraph marks on the end are counted")
        .IsEqual Sh_Trailing_Para_Marks("abc"), 0
        .IsEqual Sh_Trailing_Para_Marks("abc" & vbCr), 1
        .IsEqual Sh_Trailing_Para_Marks("abc" & vbCr & vbCr), 2
        .IsEqual Sh_Trailing_Para_Marks("abc" & vbCr & vbCr & vbCr), 3
    End With

    With Suite.Test("a paragraph mark in the middle is not on the end")
        .IsEqual Sh_Trailing_Para_Marks("a" & vbCr & "b"), 0
        .IsEqual Sh_Trailing_Para_Marks("a" & vbCr & "b" & vbCr), 1
    End With

    With Suite.Test("text that is nothing but paragraph marks")
        .IsEqual Sh_Trailing_Para_Marks(""), 0
        .IsEqual Sh_Trailing_Para_Marks(vbCr), 1
        .IsEqual Sh_Trailing_Para_Marks(vbCr & vbCr), 2
    End With

    ' --- The AutoCorrect table is written one entry to a line, tab separated.
    With Suite.Test("the three characters that need escaping are escaped")
        .IsEqual Sh_AC_Escape("a" & vbTab & "b"), "a\tb"
        .IsEqual Sh_AC_Escape("a" & vbCr & "b"), "a\rb"
        .IsEqual Sh_AC_Escape("a" & vbLf & "b"), "a\nb"
        .IsEqual Sh_AC_Escape("a\b"), "a\\b"
    End With

    With Suite.Test("ordinary text is left alone")
        .IsEqual Sh_AC_Escape("teh"), "teh"
        .IsEqual Sh_AC_Escape(""), ""
        .IsEqual Sh_AC_Unescape("the"), "the"
    End With

    With Suite.Test("escaping and unescaping is a round trip")
        Dim s As String
        s = "a" & vbTab & "b\c" & vbCr & vbLf & "d"
        .IsEqual Sh_AC_Unescape(Sh_AC_Escape(s)), s
    End With

    With Suite.Test("a backslash followed by t is not a tab")
        ' The reason Sh_AC_Unescape walks the string instead of running four Replace passes:
        ' done in the wrong order, "\\t" turns into a tab.
        .IsEqual Sh_AC_Escape("\t"), "\\t"
        .IsEqual Sh_AC_Unescape("\\t"), "\t"
        .NotEqual Sh_AC_Unescape("\\t"), vbTab
    End With

    ' --- List markers.
    With Suite.Test("a letter marker is its position in the alphabet")
        .IsEqual Lp_Marker_Value("a"), 1
        .IsEqual Lp_Marker_Value("z"), 26
        .IsEqual Lp_Marker_Value("A"), 1
        .IsEqual Lp_Marker_Value(" c "), 3
    End With

    With Suite.Test("a number marker is itself, up to 999")
        .IsEqual Lp_Marker_Value("1"), 1
        .IsEqual Lp_Marker_Value("999"), 999
        .IsEqual Lp_Marker_Value("1000"), 0
        .IsEqual Lp_Marker_Value("0"), 0
    End With

    With Suite.Test("anything else is not a marker")
        .IsEqual Lp_Marker_Value(""), 0
        .IsEqual Lp_Marker_Value("ab"), 0
        .IsEqual Lp_Marker_Value("-"), 0
    End With

    With Suite.Test("the roman reading of a marker is kept separate from the letter reading")
        ' "i" and "v" are both letters and numerals, and which one a list means only shows
        ' in the run - so the caller tries both.
        .IsEqual Lp_Marker_Value("i"), 9, "i read as a letter"
        .IsEqual Lp_Roman_Value("i"), 1, "i read as a numeral"
        .IsEqual Lp_Roman_Value("iv"), 4
        .IsEqual Lp_Roman_Value("x"), 10
    End With

    With Suite.Test("the roman marker reading stops at ten")
        .IsEqual Lp_Roman_Value("xi"), 0
        .IsEqual Lp_Roman_Value(""), 0
        .IsEqual Lp_Roman_Value("chapter"), 0
    End With

    Set TestTextHelpers_Suite = Suite
End Function

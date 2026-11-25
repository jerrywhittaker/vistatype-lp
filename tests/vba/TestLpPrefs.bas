Attribute VB_Name = "TestLpPrefs"
'@uses Lp_Pref_Line_Values
'@uses Lp_Pref_Letter_Values
'@uses Lp_Pref_Marker_Names
'@uses Lp_Pref_Marker_Canonical
'@uses Lp_Pref_Marker_Index
'@uses Lp_Pref_Marker_Style_Name
'@uses Lp_Pref_Marker_Look_Name
'@uses Lp_Pref_Marker_Caption
'@uses Lp_Pref_Number_Text
'@uses Lp_Pref_Parse
'@uses Lp_Pref_Line_Valid
'@uses Lp_Pref_Letter_Valid
'@uses Lp_Pref_Pick_Line
'@uses Lp_Pref_Pick_Letter
'@uses Lp_Pref_Pick_Marker
'@uses Lp_Pref_Line_Index
'@uses Lp_Pref_Letter_Index
'@uses Lp_Pref_Line_Label
'@uses Lp_Pref_Letter_Label
'@uses Lp_Pref_Builtin_Line_Multiple
'@uses Lp_Pref_Builtin_Letter_Spacing
'@uses Lp_Pref_Builtin_Italics_Marker
'
' Spacing and Italics Settings (Jerry, 11/24/2026): the pure helpers behind the book's line
' spacing, letter spacing and italics marker - the lists the box offers, the checks a stored
' value must pass, the order of the three sources (the book, the saved default, built in), the
' "= 27 pt at 18 pt" text and marker name to style name. Strings and numbers in and out; the
' getters that read a Document or VistaType.ini are checked on the source by
' tests/test_lp_prefs.py.
'
' Numbers are compared as Lp_Pref_Number_Text strings: a Single 0.12 is not the Double 0.12.
'
' Version: 1.1  Date: 11/24/2026 - the marker's style is the tag "Italics X", based on "Text X"
' Version: 1.0  Date: 11/24/2026
'
Option Explicit

Public Function TestLpPrefs_Suite() As TestSuite
    Dim Suite As New TestSuite
    Dim v As Variant
    Suite.Description = "Spacing and Italics Settings"

    ' --- the lists the box offers (Jerry, 11/24/2026)
    With Suite.Test("line spacing is 1.00 to 2.00 in steps of 0.05")
        v = Lp_Pref_Line_Values()
        .IsEqual LBound(v), 0
        .IsEqual UBound(v), 20
        .IsEqual Lp_Pref_Number_Text(v(0)), "1.00"
        .IsEqual Lp_Pref_Number_Text(v(1)), "1.05"
        .IsEqual Lp_Pref_Number_Text(v(5)), "1.25"
        .IsEqual Lp_Pref_Number_Text(v(20)), "2.00"
        .IsEqual v(5), CSng(1.25)
    End With

    With Suite.Test("letter spacing is 0.00 to 0.25 in steps of 0.01")
        v = Lp_Pref_Letter_Values()
        .IsEqual LBound(v), 0
        .IsEqual UBound(v), 25
        .IsEqual Lp_Pref_Number_Text(v(0)), "0.00"
        .IsEqual Lp_Pref_Number_Text(v(12)), "0.12"
        .IsEqual Lp_Pref_Number_Text(v(25)), "0.25"
    End With

    With Suite.Test("the six markers, Dash first")
        v = Lp_Pref_Marker_Names()
        .IsEqual UBound(v) - LBound(v), 5
        .IsEqual v(LBound(v)), "Dash"
        .IsEqual v(LBound(v) + 1), "TextBlue"
        .IsEqual v(LBound(v) + 5), "TextViolet"
    End With

    ' --- where a value sits in its list
    With Suite.Test("a value's place in its list")
        .IsEqual Lp_Pref_Line_Index(1.25), 5
        .IsEqual Lp_Pref_Line_Index(CSng(1.15)), 3
        .IsEqual Lp_Pref_Line_Index(2), 20
        .IsEqual Lp_Pref_Line_Index(2.05), -1
        .IsEqual Lp_Pref_Line_Index(1.27), -1
        .IsEqual Lp_Pref_Letter_Index(CSng(0.12)), 12
        .IsEqual Lp_Pref_Letter_Index(0), 0
        .IsEqual Lp_Pref_Letter_Index(0.3), -1
        .IsEqual Lp_Pref_Marker_Index("TextRed"), 4
        .IsEqual Lp_Pref_Marker_Index("Dash"), 0
        .IsEqual Lp_Pref_Marker_Index("Purple"), -1
    End With

    ' --- stored values: always a period, and anything odd refused
    With Suite.Test("numbers are written with a period and two places")
        .IsEqual Lp_Pref_Number_Text(1.25), "1.25"
        .IsEqual Lp_Pref_Number_Text(1), "1.00"
        .IsEqual Lp_Pref_Number_Text(CSng(0.07)), "0.07"
        .IsEqual Lp_Pref_Number_Text(-1), "0.00"
    End With

    With Suite.Test("only digits and one period are read")
        .IsEqual Lp_Pref_Parse("1.25"), 1.25
        .IsEqual Lp_Pref_Parse(" 2 "), 2
        .IsEqual Lp_Pref_Parse(""), -1
        .IsEqual Lp_Pref_Parse("1,25"), -1
        .IsEqual Lp_Pref_Parse("abc"), -1
        .IsEqual Lp_Pref_Parse("1.2.5"), -1
        .IsEqual Lp_Pref_Parse("-0.1"), -1
        .IsEqual Lp_Pref_Parse("."), -1
    End With

    With Suite.Test("a stored line spacing must be 1.00 to 2.00 on a step of 0.05")
        .IsOk Lp_Pref_Line_Valid("1.25")
        .IsOk Lp_Pref_Line_Valid("1")
        .IsOk Lp_Pref_Line_Valid("2.00")
        .NotOk Lp_Pref_Line_Valid("0.95")
        .NotOk Lp_Pref_Line_Valid("2.05")
        .NotOk Lp_Pref_Line_Valid("1.27")
        .NotOk Lp_Pref_Line_Valid("1,25")
        .NotOk Lp_Pref_Line_Valid("")
    End With

    With Suite.Test("a stored letter spacing must be 0.00 to 0.25 on a step of 0.01")
        .IsOk Lp_Pref_Letter_Valid("0.12")
        .IsOk Lp_Pref_Letter_Valid("0")
        .IsOk Lp_Pref_Letter_Valid("0.25")
        .NotOk Lp_Pref_Letter_Valid("0.26")
        .NotOk Lp_Pref_Letter_Valid("0.125")
        .NotOk Lp_Pref_Letter_Valid("")
        .NotOk Lp_Pref_Letter_Valid("zero")
    End With

    ' --- the order of the sources: the book, then the saved default, then built in
    With Suite.Test("the book's own line spacing wins")
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Line("1.50", "1.75")), "1.50"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Line("", "1.75")), "1.75"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Line("9", "1.75")), "1.75"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Line("", "")), "1.25"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Line("bad", "worse")), "1.25"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Builtin_Line_Multiple()), "1.25"
    End With

    With Suite.Test("the book's own letter spacing wins, and 0 is a real choice")
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Letter("0.20", "0.05")), "0.20"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Letter("0.00", "0.05")), "0.00"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Letter("", "0.05")), "0.05"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Pick_Letter("", "")), "0.12"
        .IsEqual Lp_Pref_Number_Text(Lp_Pref_Builtin_Letter_Spacing()), "0.12"
    End With

    With Suite.Test("the book's own marker wins, and a missing one is Dash")
        .IsEqual Lp_Pref_Pick_Marker("TextRed", "TextBlue"), "TextRed"
        .IsEqual Lp_Pref_Pick_Marker("", "TextBlue"), "TextBlue"
        .IsEqual Lp_Pref_Pick_Marker("Purple", "textgreen"), "TextGreen"
        .IsEqual Lp_Pref_Pick_Marker("", ""), "Dash"
        .IsEqual Lp_Pref_Builtin_Italics_Marker(), "Dash"
    End With

    ' --- marker to style
    With Suite.Test("a marker's style and caption")
        ' the tag style the marker puts on, and the LP style it is based on (Jerry, 11/24/2026)
        .IsEqual Lp_Pref_Marker_Style_Name("TextBlue"), "Italics Blue"
        .IsEqual Lp_Pref_Marker_Style_Name("TextViolet"), "Italics Violet"
        .IsEqual Lp_Pref_Marker_Style_Name("Text Orange"), "Italics Orange"
        .IsEqual Lp_Pref_Marker_Look_Name("TextBlue"), "Text Blue"
        .IsEqual Lp_Pref_Marker_Look_Name("TextRed"), "Text Red"
        .IsEqual Lp_Pref_Marker_Look_Name("Dash"), ""
        .IsEqual Lp_Pref_Marker_Style_Name("Dash"), ""
        .IsEqual Lp_Pref_Marker_Style_Name("Purple"), ""
        .IsEqual Lp_Pref_Marker_Canonical(" textred "), "TextRed"
        .IsEqual Lp_Pref_Marker_Caption("Dash"), "Dashed underline"
        .IsEqual Lp_Pref_Marker_Caption("TextGreen"), "Green"
    End With

    ' --- the "= nn pt at nn pt" text beside the lists
    With Suite.Test("line spacing in points: multiple x 1.207 x size, whole points")
        ' measured 27.15 pt at 18 pt, 36.20 at 24 (vistabuild, 11/24/2026)
        .IsEqual Lp_Pref_Line_Label(1.25, 18), "= 27 pt at 18 pt"
        .IsEqual Lp_Pref_Line_Label(1.25, 24), "= 36 pt at 24 pt"
        .IsEqual Lp_Pref_Line_Label(1, 18), "= 22 pt at 18 pt"
        .IsEqual Lp_Pref_Line_Label(2, 18), "= 43 pt at 18 pt"
    End With

    With Suite.Test("letter spacing in points, to the twentieth Word keeps")
        .IsEqual Lp_Pref_Letter_Label(CSng(0.12), 18), "= 2.15 pt at 18 pt"
        .IsEqual Lp_Pref_Letter_Label(CSng(0.1), 20), "= 2 pt at 20 pt"
        .IsEqual Lp_Pref_Letter_Label(0, 18), "= 0 pt at 18 pt"
        .IsEqual Lp_Pref_Letter_Label(CSng(0.25), 18), "= 4.5 pt at 18 pt"
    End With

    Set TestLpPrefs_Suite = Suite
End Function

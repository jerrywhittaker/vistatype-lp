Attribute VB_Name = "TestEmptyPagenums"
'@uses Sh_Tag_Empty_Pagenums
'
' The DAISY/NIMAS converter's second $pg pass: page numbers written as EMPTY <pagenum> elements,
' the number only in the id, as a DAISY book made from an EPUB writes them (Jerry, 10/25/2026,
' Bookshare's "Basic Math & Pre-Algebra For Dummies"). The shapes below are the ones measured in
' that book. Nothing here touches a document.
'
' Version: 1.0  Date: 10/25/2026
'
Option Explicit

Private Function Tagged(ByVal openTag As String, ByVal pg As String) As String
    ' What the converter's DAISY pattern leaves for a page number: the opening tag without its
    ' ">", then the $pg tag in paragraphs of its own.
    Tagged = openTag & "<p>$pg" & pg & "<p></p></pagenum><p></p>"
End Function

Public Function TestEmptyPagenums_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "Empty page number elements (DAISY from EPUB)"

    Dim t As String

    With Suite.Test("the number comes from the end of the id, its zeros dropped")
        t = "<pagenum epub:type=""pagebreak"" page=""normal"" id=""pagenum_0000163"" " & _
            "smilref=""Book00011.smil#pagenum_0000163"""
        .IsEqual Sh_Tag_Empty_Pagenums(t & " />"), Tagged(t, "163")
    End With

    With Suite.Test("the id may come before page=""normal""")
        t = "<pagenum epub:type=""pagebreak"" id=""pagenum_0000012"" page=""normal"" " & _
            "smilref=""Book00002.smil#pagenum_0000012"""
        .IsEqual Sh_Tag_Empty_Pagenums(t & " />"), Tagged(t, "12")
    End With

    With Suite.Test("5, 6 and 7 digits of padding, and zeros inside the number kept")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum page=""normal"" id=""pagenum_00007""/>"), _
                 Tagged("<pagenum page=""normal"" id=""pagenum_00007""", "7")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum page=""normal"" id=""pagenum_000070""/>"), _
                 Tagged("<pagenum page=""normal"" id=""pagenum_000070""", "70")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum page=""normal"" id=""pagenum_0000700""/>"), _
                 Tagged("<pagenum page=""normal"" id=""pagenum_0000700""", "700")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum id=""12"" />"), Tagged("<pagenum id=""12""", "12")
    End With

    With Suite.Test("only the LAST run of digits in the id counts")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum id=""ch3_page_0045"" />"), _
                 Tagged("<pagenum id=""ch3_page_0045""", "45")
    End With

    With Suite.Test("an id with no number at the end is $pg?")
        t = "<pagenum epub:type=""pagebreak"" id=""p"" page=""normal"" smilref=""Book00001.smil#p"""
        .IsEqual Sh_Tag_Empty_Pagenums(t & " />"), Tagged(t, "?")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum id=""page12a"" />"), Tagged("<pagenum id=""page12a""", "?")
    End With

    With Suite.Test("an id that is all zeros is $pg?")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum page=""normal"" id=""pagenum_00000"" />"), _
                 Tagged("<pagenum page=""normal"" id=""pagenum_00000""", "?")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum page=""normal"" id=""pagenum_0000000"" />"), _
                 Tagged("<pagenum page=""normal"" id=""pagenum_0000000""", "?")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum id=""page_10_000"" />"), _
                 Tagged("<pagenum id=""page_10_000""", "?")
    End With

    With Suite.Test("no id at all is $pg?")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum page=""normal"" />"), Tagged("<pagenum page=""normal""", "?")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum/>"), Tagged("<pagenum", "?")
        ' the number in smilref is not an id
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum smilref=""a.smil#pagenum_0000009""/>"), _
                 Tagged("<pagenum smilref=""a.smil#pagenum_0000009""", "?")
    End With

    With Suite.Test("an opening and closing tag with nothing but white space between")
        t = "<pagenum id=""pagenum_0000042"" page=""normal"""
        .IsEqual Sh_Tag_Empty_Pagenums(t & "></pagenum>"), Tagged(t, "42")
        .IsEqual Sh_Tag_Empty_Pagenums(t & ">  " & vbCrLf & vbTab & "</pagenum>"), Tagged(t, "42")
        .IsEqual Sh_Tag_Empty_Pagenums("<pagenum id=""p"">" & vbLf & "</pagenum>"), Tagged("<pagenum id=""p""", "?")
    End With

    With Suite.Test("in the middle of a sentence, inside bold")
        t = "<pagenum epub:type=""pagebreak"" page=""normal"" id=""pagenum_0000099"""
        .IsEqual Sh_Tag_Empty_Pagenums("<p>Some <strong>bold</span>" & t & " /> text</strong> here.</p>"), _
                 "<p>Some <strong>bold</span>" & Tagged(t, "99") & " text</strong> here.</p>"
    End With

    With Suite.Test("every one in the book, and nothing between them changed")
        .IsEqual Sh_Tag_Empty_Pagenums("<level2><pagenum id=""pagenum_01"" /><h2>A</h2>" & _
                                       "<li>x</li><pagenum id=""p"" /><li>y</li></level2>"), _
                 "<level2>" & Tagged("<pagenum id=""pagenum_01""", "1") & "<h2>A</h2><li>x</li>" & _
                 Tagged("<pagenum id=""p""", "?") & "<li>y</li></level2>"
    End With

    With Suite.Test("a page number with text in it is left alone")
        ' as the book has it, before the converter's own patterns
        t = "<p>x</p><pagenum id=""p12"" page=""normal"">12</pagenum><pagenum id=""p13"">13-14</pagenum>"
        .IsEqual Sh_Tag_Empty_Pagenums(t), t
        ' and as the converter's DAISY and NIMAS patterns leave it
        t = "<pagenum id=""p12"" page=""normal""<p>$pg12<p></p></pagenum><p></p>"
        .IsEqual Sh_Tag_Empty_Pagenums(t), t
        t = "<pagenum id=""p12"" page=""normal""<p>$pg12-13</pagenum><p></p>"
        .IsEqual Sh_Tag_Empty_Pagenums(t), t
    End With

    With Suite.Test("other tags and text are left alone")
        t = "<p>Cost: $5 and $10.</p><pagenumber id=""p1"" /><prodnote id=""n2"" />"
        .IsEqual Sh_Tag_Empty_Pagenums(t), t
        .IsEqual Sh_Tag_Empty_Pagenums(""), ""
    End With

    Set TestEmptyPagenums_Suite = Suite
End Function

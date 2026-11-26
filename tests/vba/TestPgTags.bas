Attribute VB_Name = "TestPgTags"
'@uses Sh_Pg_Tags_No_Char_Style
'@uses Sh_Pg_Red_Replace
'
' $pg tags are never italic or underlined (Jerry, 11/26/2026). In a hidden scratch book of its
' own, "$pg 12" is made italic, dash-underlined, bold, 20 pt Arial and "Italics Red" - the shape
' an italic run around a tag would give it once the italics step had been - and the two steps
' of Sh_Color_Dollar_PG_Red's targetDoc arm are run on it. The tag must come out red, not italic,
' not underlined, not bold, with no character style and still 20 pt Arial; "12" must keep all of
' what it had. A second tag with no character style must come out red and plain too.
'
' Like TestReviewProdnotes, the scratch book is made Visible:=False, never touches the Selection,
' makes its own style so no template is attached, and is closed unsaved whatever happens.
'
' Version: 1.0  Date: 11/26/2026
'
Option Explicit

Public Function TestPgTags_Suite() As TestSuite
    Dim Suite As New TestSuite
    Suite.Description = "$pg tags are never italic or underlined"

    TestPgTags_In_A_Book Suite.Test("in a scratch book: an italic, underlined, Italics Red $pg 12")

    Set TestPgTags_Suite = Suite
End Function

Private Sub TestPgTags_In_A_Book(ByVal T As TestCase)
    Dim d As Document
    Dim st As Style
    Dim r As Range
    Dim tag As Range, num As Range, tag2 As Range
    Dim dpf As String
    Dim stage As String

    On Error GoTo Failed
    stage = "making the scratch book"
    Set d = Documents.Add(Visible:=False)
    Set st = d.Styles.Add("Italics Red", wdStyleTypeCharacter)
    st.Font.Color = wdColorRed
    st.Font.Italic = True
    dpf = d.Styles(wdStyleDefaultParagraphFont).NameLocal
    ' Positions: "Before " is 0-6, "$pg" 7-9, " " 10, "12" 11-12; second paragraph "$pg 34" from 20
    d.Content.Text = "Before $pg 12 after" & vbCr & "$pg 34"
    Set r = d.Range(7, 13)
    r.Style = st
    r.Font.Italic = True
    r.Font.Underline = wdUnderlineDash
    r.Font.Bold = True
    r.Font.Size = 20
    r.Font.Name = "Arial"
    Set r = d.Range(20, 23)
    r.Font.Italic = True
    r.Font.Underline = wdUnderlineSingle
    T.IsEqual d.Range(7, 10).Text, "$pg", "the first tag is where the test expects"
    T.IsEqual d.Range(20, 23).Text, "$pg", "the second tag is where the test expects"
    T.IsEqual TestPgTags_Char_Style(d.Range(7, 10)), "Italics Red", "the tag starts in Italics Red"

    stage = "running Sh_Pg_Tags_No_Char_Style and Sh_Pg_Red_Replace"
    Sh_Pg_Tags_No_Char_Style d
    Sh_Pg_Red_Replace d

    stage = "reading the first tag"
    Set tag = d.Range(7, 10)
    T.IsEqual tag.Text, "$pg", "the tag's text is kept"
    T.IsEqual TestPgTags_Char_Style(tag), dpf, "the tag has no character style"
    T.IsEqual tag.Font.Italic, False, "the tag is not italic"
    T.IsEqual tag.Font.Underline, wdUnderlineNone, "the tag is not underlined"
    T.IsEqual tag.Font.Bold, False, "the tag is not bold"
    T.IsEqual tag.Font.Color, wdColorRed, "the tag is red"
    T.IsEqual tag.Font.Size, 20, "the tag keeps its size"
    T.IsEqual tag.Font.Name, "Arial", "the tag keeps its font"

    stage = "reading the page number"
    Set num = d.Range(11, 13)
    T.IsEqual num.Text, "12", "the page number's text is kept"
    T.IsEqual TestPgTags_Char_Style(num), "Italics Red", "the page number keeps its style"
    T.IsEqual num.Font.Italic, True, "the page number keeps its italic"
    T.IsEqual num.Font.Underline, wdUnderlineDash, "the page number keeps its underline"
    T.IsEqual num.Font.Bold, True, "the page number keeps its bold"

    stage = "reading the second tag"
    Set tag2 = d.Range(20, 23)
    T.IsEqual tag2.Text, "$pg", "the second tag's text is kept"
    T.IsEqual tag2.Font.Italic, False, "the second tag is not italic"
    T.IsEqual tag2.Font.Underline, wdUnderlineNone, "the second tag is not underlined"
    T.IsEqual tag2.Font.Color, wdColorRed, "the second tag is red"

    d.Close SaveChanges:=False
    Exit Sub

Failed:
    T.Fail "VBA error " & Err.Number & " while " & stage & " - " & Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.Close SaveChanges:=False
    Err.Clear
End Sub

' The name of a range's character style. IsEqual compares a Style object by identity, not name.
Private Function TestPgTags_Char_Style(ByVal r As Range) As String
    Dim nm As String
    nm = r.CharacterStyle
    TestPgTags_Char_Style = nm
End Function

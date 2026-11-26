"""$pg tags are never italic or underlined (Jerry, 11/26/2026).

Sh_Color_Dollar_PG_Red runs last in File Cleanup and Normalize Styles, after the italics step. An
italic run around a tag would make the tag italic, and the italics step would dash it or give it
the "Italics <color>" style. So the tag is set not italic and not underlined, on both arms (the
Selection one the older callers use, and the targetDoc one), and a character style on it is taken
off first.

Measured on vistabuild, 11/26/2026: Replacement.Style = Normal (a paragraph style) leaves a
character style such as "Italics Red" on the tag, and Replacement.Style = Default Paragraph Font
takes it off but strips the tag's direct font and size too. Hence Sh_Pg_Tags_No_Char_Style, a
walk that puts font and size back. tests/vba/TestPgTags.bas runs both helpers on a scratch book.

What this checks, on the source:
  1. Both arms set Replacement.Font.Italic = False and Replacement.Font.Underline = wdUnderlineNone
     alongside the red and the not-bold, and still replace "$pg" with "$pg" only.
  2. Both arms take the character style off first.
  3. The walk sets Default Paragraph Font on three characters only and restores size and font.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures


def body_of(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return "\n".join(procs[name])


def test_selection_arm_not_italic_not_underlined(repo_root):
    body = body_of(repo_root, "Sh_Color_Dollar_PG_Red")
    sel = body.split('Application.Run MacroName:="Sh_Is_Doc_Open"', 1)[1]
    for line in (r"Selection\.Find\.Replacement\.Font\.Color = wdColorRed",
                 r"Selection\.Find\.Replacement\.Font\.Bold = False",
                 r"Selection\.Find\.Replacement\.Font\.Italic = False",
                 r"Selection\.Find\.Replacement\.Font\.Underline = wdUnderlineNone"):
        assert re.search(line, sel), line
    # the tag is the whole of what is found and replaced - not the page number after it
    assert re.search(r'\.Text = "\$pg"\s*\n\s*\.Replacement\.Text = "\$pg"', sel)
    assert ".MatchWildcards = False" in sel
    # the character style comes off before the replace
    assert sel.index("Sh_Pg_Tags_No_Char_Style ActiveDocument") < sel.index("Execute Replace:=wdReplaceAll")


def test_target_doc_arm_uses_the_helpers(repo_root):
    body = body_of(repo_root, "Sh_Color_Dollar_PG_Red")
    arm = body.split("If Not targetDoc Is Nothing Then", 1)[1].split("Exit Sub", 1)[0]
    assert arm.index("Sh_Pg_Tags_No_Char_Style targetDoc") < arm.index("Sh_Pg_Red_Replace targetDoc")


def test_red_replace_not_italic_not_underlined(repo_root):
    body = body_of(repo_root, "Sh_Pg_Red_Replace")
    for line in (r"\.Replacement\.Style = doc\.Styles\(wdStyleNormal\)",
                 r"\.Replacement\.Font\.Color = wdColorRed",
                 r"\.Replacement\.Font\.Bold = False",
                 r"\.Replacement\.Font\.Italic = False",
                 r"\.Replacement\.Font\.Underline = wdUnderlineNone",
                 r'\.Text = "\$pg"\s*\n\s*\.Replacement\.Text = "\$pg"',
                 r"\.MatchWildcards = False",
                 r"\.Execute Replace:=wdReplaceAll"):
        assert re.search(line, body), line


def test_char_style_walk_keeps_size_and_font(repo_root):
    body = body_of(repo_root, "Sh_Pg_Tags_No_Char_Style")
    assert re.search(r"Set rr = doc\.Range\(starts\(i\), starts\(i\) \+ 3\)", body)
    assert re.search(r"rr\.Style = doc\.Styles\(wdStyleDefaultParagraphFont\)", body)
    assert re.search(r"rr\.Font\.Size = sz", body)
    assert re.search(r"rr\.Font\.Name = nm", body)
    # search first, edit after: nothing is restyled inside the Find loop
    loop = body.split("Do While .Execute", 1)[1].split("Loop", 1)[0]
    assert ".Style =" not in loop
    assert re.search(r"If guard > \d+ Then Exit Do", loop)

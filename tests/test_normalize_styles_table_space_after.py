"""Normalize Styles gives text in a table 0 points after (Jerry, 11/24/2026).

"For text in a table, the space after is set to the base font size (18 points for a base font
size of 18)... that is too much it should be set to the 0 point setting." The line spacing (At
least, size + 4) stays.

A check on the source: Normalize Styles reaches a Document, which the headless VBA runner
cannot do (tests/vba/README.md). The table pass must come AFTER the base-size pass and the
style loop that gives list and TOC paragraphs 4 points, or one of those puts space back, and
before Sh_Color_Dollar_PG_Red, which stays last. It must not touch line spacing.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

SUB = "Lp_Normalize_Styles"


def first(lines, pattern, after=-1):
    return next((i for i, c in enumerate(lines) if i > after and re.search(pattern, c)), None)


def test_table_text_gets_zero_after_last(repo_root):
    lines = procedures(module_text(repo_root))[SUB]
    base = first(lines, r"Content\.ParagraphFormat\.SpaceAfter\s*=\s*Lp_Base_Font_Size")
    loop = first(lines, r"^\s*\.SpaceAfter\s*=\s*spaceAfterPts")
    each = first(lines, r"For\s+Each\s+\w+\s+In\s+ActiveDocument\.Tables")
    zero = first(lines, r"\.Range\.ParagraphFormat\.SpaceAfter\s*=\s*0\s*$", each or -1)
    red = first(lines, r"Sh_Color_Dollar_PG_Red")
    assert None not in (base, loop, each, zero, red)
    assert base < loop < each < zero < red
    body = "\n".join(lines[each:zero + 3])
    assert "LineSpacing" not in body, "the table pass must leave line spacing alone"

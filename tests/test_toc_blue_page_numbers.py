"""Format the TOC treats a line with a blue page number as an entry, and takes the blue off.

Jerry, 9/21/2026, on "TOC Table 2 Formatting.docx": "There are items in blue which are
underlined and have page numbers. The page numbers have been ignored, and the blue underlines
should be automatic color with no underlines, and the page number should be recognized (and
marked with a tab before the page numbers) and styled as TOC 1."

Measured in a copy of that file: the blue and the underline are direct formatting, not links,
and 77 of the 94 blue lines are bold from end to end. The heading test in
Lp_TOC_CleanAndFormat_TOC called every wholly bold line with no dot leader and no tab a section
heading, so those 77 got no tab and no TOC 1.

The text-and-color decision is Lp_TOC_Line_Is_Link_Entry, tested in VBA by
tests/vba/TestTocLinkEntry.bas. What that cannot see is whether the macro ASKS it when building
the heading list, and whether the blue is taken off only AFTER that list is built - the other
way round, there is no blue left to read. The macro works on a Range, which hangs the headless
runner (tests/vba/README.md), so this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

ROOT_SUB = "Lp_TOC_CleanAndFormat_TOC"


def the_macro(repo_root):
    procs = procedures(module_text(repo_root))
    assert ROOT_SUB in procs, "Lp_TOC_CleanAndFormat_TOC is missing from LPandBrlMacros.bas"
    return procs


def heading_decision(lines):
    """The statement that fills isHeading(paraNo), with its continuation lines joined."""
    start = next((i for i, c in enumerate(lines) if re.search(r"isHeading\(paraNo\)\s*=", c)),
                 None)
    assert start is not None, "The heading list is no longer filled in by isHeading(paraNo) ="
    stmt = []
    for c in lines[start:]:
        stmt.append(c.rstrip())
        if not c.rstrip().endswith("_"):
            break
    return start, " ".join(s.rstrip("_ ") for s in stmt)


def test_a_blue_page_number_is_not_a_heading(repo_root):
    _, stmt = heading_decision(the_macro(repo_root)[ROOT_SUB])
    assert re.search(
        r"And\s+Not\s+Lp_TOC_Line_Is_Link_Entry\(\s*paraText\s*,\s*"
        r"Lp_TOC_Last_Char_Color\(\s*para\.Range\s*\)\s*\)", stmt), (
        "A wholly bold line whose page number is blue must not be read as a heading - the "
        "heading test must ask Lp_TOC_Line_Is_Link_Entry with the line's last color.")


def test_the_blue_comes_off_after_the_heading_list_is_built(repo_root):
    lines = the_macro(repo_root)[ROOT_SUB]
    decided, _ = heading_decision(lines)
    unblue = next((i for i, c in enumerate(lines) if re.search(r"\bLp_TOC_Unblue\s+workRng\b", c)),
                  None)
    assert unblue is not None, "Format the TOC must call Lp_TOC_Unblue on the TOC (workRng)"
    assert unblue > decided, (
        "The blue must be taken off AFTER the heading list reads it, or no line is ever a link entry")


def test_unblue_makes_blue_automatic_with_no_underline_and_stays_in_the_toc(repo_root):
    body = "\n".join(the_macro(repo_root)["Lp_TOC_Unblue"])
    for needed in (r"\.Font\.Color\s*=\s*wdColorBlue",
                   r"\.Replacement\.Font\.Color\s*=\s*wdColorAutomatic",
                   r"\.Replacement\.Font\.Underline\s*=\s*wdUnderlineNone",
                   r"\.Format\s*=\s*True",
                   r"\.Wrap\s*=\s*wdFindStop",
                   r'\.Replacement\.Text\s*=\s*""'):
        assert re.search(needed, body), f"Lp_TOC_Unblue is missing {needed}"

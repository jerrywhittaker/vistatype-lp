"""Format the TOC on a contents page where nearly every entry is bold, and the bullets are black squares.

Jerry, 9/28/2026, on "TOC with Black Boxes.docx": no leaders and no tabs, nearly every title and
page number bold, 105 lesson lines opening with U+25A0 BLACK SQUARE, and four "Chapter Practice"
entries whose page number is followed by a manual line break and a three-line legend. Measured
on a copy with the real passes:

  * every black square survived - U+25A0 was not in the bullet class;
  * 63 wholly bold entries ("Vocabulary 2", "Chapter Practice 47", "Glossary A1") were read as
    SECTION HEADINGS - the 9/7/2026 rule, "a wholly bold line with no leader or tab is a
    heading" - so they got no tab and no TOC 1. Jerry's decision the same day: "yes, bold
    doesn't mean heading there". When most lines ending in a page number are wholly bold, bold
    no longer marks a heading (Lp_TOC_Bold_Marks_Headings);
  * the four "Chapter Practice 79<line break>..." entries got no tab, because the page number
    was not at the end of the paragraph. The break after it becomes a paragraph mark first
    (Lp_TOC_Split_Line_Breaks, deciding through Lp_TOC_Break_Ends_Entry).

The two decisions are tested in VBA by tests/vba/TestTocBoldEntries.bas. What those cannot see
is where the macro asks them: the bullets must come off, and the breaks be split, BEFORE the
heading list is built - otherwise the bullet makes a bold lesson line look mixed, and the line
count shifts under the list. The macro works on a Range, which hangs the headless runner
(tests/vba/README.md), so this is a check on the source.
"""
import re

from test_toc_blue_page_numbers import ROOT_SUB, heading_decision, the_macro


def first_line(lines, pattern, what):
    i = next((n for n, c in enumerate(lines) if re.search(pattern, c)), None)
    assert i is not None, what
    return i


def bullet_pass(lines):
    """The Lp_TOC_Replace that strips the bullet class, with its continuation lines joined."""
    start = first_line(lines, r'Lp_TOC_Replace\s+workRng\s*,\s*"\["\s*&\s*ChrW\(&H2022\)',
                       "The bullet pass (Lp_TOC_Replace workRng, \"[\" & ChrW(&H2022) ...) is gone")
    stmt = []
    for c in lines[start:]:
        stmt.append(c.rstrip())
        if not c.rstrip().endswith("_"):
            break
    return start, " ".join(s.rstrip("_ ") for s in stmt)


def test_the_black_square_is_in_the_bullet_class(repo_root):
    _, stmt = bullet_pass(the_macro(repo_root)[ROOT_SUB])
    assert "ChrW(&H25A0)" in stmt, (
        "U+25A0 BLACK SQUARE must be in the bullet class - 'TOC with Black Boxes.docx' opens "
        "105 lesson lines with it")


def test_the_bullets_come_off_before_the_heading_list(repo_root):
    lines = the_macro(repo_root)[ROOT_SUB]
    decided, _ = heading_decision(lines)
    bullets, _ = bullet_pass(lines)
    assert bullets < decided, (
        "The bullets must be stripped BEFORE isHeading(paraNo) is filled in - a plain bullet in "
        "front of a bold lesson line makes it look mixed, not wholly bold")


def test_bold_marks_a_heading_only_when_the_count_says_so(repo_root):
    lines = the_macro(repo_root)[ROOT_SUB]
    decided, stmt = heading_decision(lines)
    asked = first_line(lines, r"=\s*Lp_TOC_Bold_Marks_Headings\(",
                       "Format the TOC must ask Lp_TOC_Bold_Marks_Headings whether bold marks "
                       "a heading in this contents page")
    assert asked < decided, "Lp_TOC_Bold_Marks_Headings must be asked before the heading list"
    assert re.search(r"isHeading\(paraNo\)\s*=\s*boldMarksHeadings\s+And\s+Lp_TOC_Para_Is_Bold",
                     stmt), (
        "A wholly bold line may be a heading only when boldMarksHeadings says bold means "
        "heading here")


def test_the_line_breaks_are_split_before_the_heading_list(repo_root):
    lines = the_macro(repo_root)[ROOT_SUB]
    decided, _ = heading_decision(lines)
    split = first_line(lines, r"^\s*Lp_TOC_Split_Line_Breaks\s+workRng\b",
                       "Format the TOC must call Lp_TOC_Split_Line_Breaks on the TOC (workRng)")
    assert split < decided, (
        "A line break must become a paragraph mark BEFORE the heading list - the list is "
        "counted by paragraph")


def test_the_split_asks_the_tested_decision(repo_root):
    body = "\n".join(the_macro(repo_root)["Lp_TOC_Split_Line_Breaks"])
    assert re.search(r"\bLp_TOC_Break_Ends_Entry\(", body), (
        "Lp_TOC_Split_Line_Breaks must decide through Lp_TOC_Break_Ends_Entry, which is the "
        "part tests/vba/TestTocBoldEntries.bas can reach")


def test_the_bold_count_reads_the_normalized_line(repo_root):
    """Found in review, 9/28/2026: the count runs before the tab and leader passes, so
    "Chapter 5<tab>36" and "Chapter 5.....36" were never counted as plain entries, and a page of
    them with two bold "Section 1" headings stopped treating bold as a heading."""
    lines = the_macro(repo_root)[ROOT_SUB]
    asked = first_line(lines, r"=\s*Lp_TOC_Bold_Marks_Headings\(",
                       "Format the TOC must ask Lp_TOC_Bold_Marks_Headings")
    counted = [c for c in lines[:asked] if "Lp_TOC_Line_Ends_In_Page_Number(" in c]
    assert counted, "The bold count must ask Lp_TOC_Line_Ends_In_Page_Number"
    assert all("Lp_TOC_Count_Text(" in c for c in counted), (
        "The bold count must read each line through Lp_TOC_Count_Text, so a tabbed or "
        "leadered entry is counted")


def test_the_split_needs_a_bullet_after_the_break(repo_root):
    body = "\n".join(the_macro(repo_root)["Lp_TOC_Break_Ends_Entry"])
    assert "ChrW(&H25A0)" in body and "ChrW(&H2022)" in body, (
        "Lp_TOC_Break_Ends_Entry must split only before a bullet - a wrapped heading such as "
        "'Unit 1<line break>Place Value and Numbers' must stay one paragraph")

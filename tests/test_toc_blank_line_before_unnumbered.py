"""Format the TOC puts a blank line before every line that has no page number.

Jerry, 9/21/2026: "the blank line goes before any line that has no associated page number."

Until 3.0.476 only a WHOLLY BOLD line got that blank line (Lp_TOC_Space_A_Heading, from
3.0.407). On "TOC Table 2 Formatting.docx" the leading bullet and its space are never bold, so
no bulleted line was a heading, and every line with no page number - the "Poetry:" reading
titles, the "ACTIVITY Unit" lines - went down the branch that resets a line to Normal and set
no space before it. From 3.0.410 to 3.0.474 those lines were deleted outright, which hid it.

The decision itself - does this line text want a blank line - is a pure string function,
Lp_TOC_Line_Gets_Blank_Before, and is tested in VBA by tests/vba/TestTocBlankLine.bas. What
that cannot see is whether the macro CALLS it, and whether it asks after the non-breaking
spaces have been turned into ordinary ones. The macro works on a Range, which hangs the
headless runner (tests/vba/README.md), so this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

ROOT_SUB = "Lp_TOC_CleanAndFormat_TOC"


def no_number_branch(lines):
    """The lines of the 'If Not hasNumberOrRoman Then' branch, up to the SkipPara label."""
    start = next(i for i, c in enumerate(lines)
                 if re.match(r"\s*If\s+Not\s+hasNumberOrRoman\s+Then\s*$", c, re.IGNORECASE))
    end = next(i for i, c in enumerate(lines)
               if i > start and re.match(r"\s*SkipPara:", c))
    return lines[start:end]


def first_line(lines, pattern):
    return next((i for i, c in enumerate(lines) if re.search(pattern, c)), None)


def the_macro(repo_root):
    procs = procedures(module_text(repo_root))
    assert ROOT_SUB in procs, "Lp_TOC_CleanAndFormat_TOC is missing from LPandBrlMacros.bas"
    return procs


def test_a_line_with_no_page_number_gets_the_blank_line(repo_root):
    branch = "\n".join(no_number_branch(the_macro(repo_root)[ROOT_SUB]))
    assert re.search(
        r"If\s+Lp_TOC_Line_Gets_Blank_Before\(paraRange\.Text\)\s+Then\s+"
        r"Lp_TOC_Space_A_Heading\s+para\b", branch), (
        "The no-page-number branch of Format the TOC must give the line its blank line "
        "(Lp_TOC_Space_A_Heading), guarded by Lp_TOC_Line_Gets_Blank_Before.")


def test_the_page_number_is_looked_for_after_the_non_breaking_spaces_go(repo_root):
    # "1.2<NBSP>Perception Is Everything<NBSP>6" is an ordinary entry. Asked before the
    # non-breaking space is an ordinary one, it would read as a line with no page number and
    # be given a blank line.
    lines = the_macro(repo_root)[ROOT_SUB]
    nbsp = first_line(lines, r'\.Text\s*=\s*Chr\(160\)')
    bullets = first_line(lines, r"ChrW\(&HF0B7\)")
    decided = first_line(lines, r"hasNumberOrRoman\s*=\s*Lp_TOC_Line_Ends_In_Page_Number\(")
    assert decided is not None, (
        "hasNumberOrRoman must come from Lp_TOC_Line_Ends_In_Page_Number, the tested function")
    assert nbsp is not None and nbsp < decided
    assert bullets is not None and bullets < decided


def test_the_tab_pass_uses_the_same_page_number_pattern(repo_root):
    # The line that gets no blank line must be the line that gets its tab and TOC 1, so both
    # questions are asked with one pattern.
    procs = the_macro(repo_root)
    lines = procs[ROOT_SUB]
    assert first_line(lines, r"\.pattern\s*=\s*Lp_TOC_Page_Number_Pattern\(\)") is not None
    assert first_line(procs["Lp_TOC_Line_Ends_In_Page_Number"],
                      r"\.pattern\s*=\s*Lp_TOC_Page_Number_Pattern\(\)") is not None

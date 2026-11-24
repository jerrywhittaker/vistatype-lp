"""Clean and format the TOC leaves a line already in TOC 1 to TOC 5 in its own style.

Jerry, 11/24/2026: "it changes all entries to TOC 1 style. If an entry is already in a TOC
style (1-5) then it should remain unchanged. Only non TOC styles should be changed to TOC
Styles."

Two places changed the style: the second loop put TOC 1 on every line ending in a page number,
and the first loop reset every line with no page number to Normal. Both must now ask
Lp_TOC_Para_Has_Entry_Style first. The macro works on a Range, which hangs the headless runner
(tests/vba/README.md), so this is a check on the source; which style names count is tested in
VBA by tests/vba/TestTocBox.bas (Lp_TOC_Is_Entry_Style).
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

ROOT_SUB = "Lp_TOC_CleanAndFormat_TOC"
HELPER = "Lp_TOC_Para_Has_Entry_Style"


def the_macro(repo_root):
    procs = procedures(module_text(repo_root))
    assert ROOT_SUB in procs and HELPER in procs
    return procs


def test_the_helper_asks_the_tested_style_list(repo_root):
    body = "\n".join(the_macro(repo_root)[HELPER])
    assert "Lp_TOC_Is_Entry_Style(" in body


def test_toc_1_only_goes_on_a_line_not_already_in_a_toc_style(repo_root):
    lines = the_macro(repo_root)[ROOT_SUB]
    sets_toc1 = [c for c in lines if re.search(r'\.Style\s*=\s*"TOC 1"', c)]
    assert sets_toc1, "the macro no longer sets TOC 1 anywhere"
    for c in sets_toc1:
        assert re.search(r"If\s+Not\s+" + HELPER + r"\(paraRange\)\s+Then", c), (
            f"TOC 1 must only go on a line not already in TOC 1 to TOC 5: {c.strip()}")


def test_a_toc_styled_line_is_not_reset_to_normal_but_keeps_its_blank_line(repo_root):
    lines = the_macro(repo_root)[ROOT_SUB]
    decided = next((i for i, c in enumerate(lines)
                    if re.search(r"keepTocStyle\s*=\s*" + HELPER + r"\(paraRange\)", c)), None)
    branch = next(i for i, c in enumerate(lines)
                  if re.match(r"\s*If\s+Not\s+hasNumberOrRoman\s+Then\s*$", c, re.IGNORECASE))
    guard = next((i for i, c in enumerate(lines)
                  if i > branch and re.match(r"\s*If\s+Not\s+keepTocStyle\s+Then\s*$", c)), None)
    first_normal = next(i for i, c in enumerate(lines)
                        if i > branch and "Styles(wdStyleNormal)" in c and ".Style" in c)
    blank = next(i for i, c in enumerate(lines)
                 if i > branch and "Lp_TOC_Space_A_Heading" in c)
    assert decided is not None and decided < branch
    assert guard is not None and guard < first_normal, (
        "the reset to Normal must sit inside If Not keepTocStyle")
    # Jerry's 9/21/2026 rule still holds for a TOC-styled line with no page number: the blank
    # line call must not be inside the keepTocStyle guard.
    depth = 0
    for c in lines[guard:blank]:
        if re.match(r"\s*If\b.*\bThen\s*$", c):
            depth += 1
        elif re.match(r"\s*End\s+If\b", c):
            depth -= 1
    assert depth == 0, "the blank line before must not depend on keepTocStyle"


def test_a_toc_that_ends_the_book_keeps_its_last_line_style(repo_root):
    body = "\n".join(the_macro(repo_root)[ROOT_SUB])
    assert re.search(
        r"If\s+endClipped\s+Then\s+On Error Resume Next\s+endStyle\s*=\s*vbNullString\s+"
        r"endStyle\s*=\s*doc\.Paragraphs\(doc\.Paragraphs\.count\)\.Style\.NameLocal\s+"
        r"If\s+Lp_TOC_Is_Entry_Style\(endStyle\)\s+Then", body), (
        "a TOC that ends the book must carry its last line's TOC style into the scratch document")

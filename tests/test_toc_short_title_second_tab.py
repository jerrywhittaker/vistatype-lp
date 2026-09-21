"""Format the TOC adds a second tab after a title too short to reach the dot leader.

Jerry, 9/21/2026, on "TOC Table 1 Formatting.docx", build 3.0.480: Lp_TOC_CleanAndFormat_TOC made
"Index<tab>18" in TOC 1 and the page showed "Index 18" with no leader. TOC 1 to TOC 5 in
LargePrintTemplate.dotx hang by 0.75", and Word treats a hanging indent as a tab stop, so a title
that ends before it sends its tab to the indent. Jerry's cure: "if the ellipses are not present
after applying the TOC style put in a second tab next to the first one."

The decision is Lp_TOC_Tab_Stops_At_Indent, tested in VBA by tests/vba/TestTocShortTitle.bas.
The measuring and the inserting work on a Range, which hangs the headless runner
(tests/vba/README.md), so what this checks is WHERE the macro does it - each of these, got
wrong, gives back the fault or a new one:

  * after the TOC style is applied, or there is no hanging indent yet to measure against;
  * in the hidden scratch document, before the text comes home, or every added tab is one more
    Ctrl+Z in the book;
  * from the bottom up, or a tab added above moves every position still to be read;
  * before the reference page passes, which insert characters and move them all;
  * after the pass that turns every tab into a space, which is what stops a second run of the
    macro stacking a third tab.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

ROOT_SUB = "Lp_TOC_CleanAndFormat_TOC"


def macro_lines(repo_root):
    procs = procedures(module_text(repo_root))
    assert ROOT_SUB in procs, "Lp_TOC_CleanAndFormat_TOC is missing from LPandBrlMacros.bas"
    return procs, procs[ROOT_SUB]


def first(lines, pattern, what):
    i = next((n for n, c in enumerate(lines) if re.search(pattern, c)), None)
    assert i is not None, f"Format the TOC no longer has {what}"
    return i


def test_the_second_tab_pass_sits_where_it_must(repo_root):
    _, lines = macro_lines(repo_root)
    tabs_to_spaces = first(lines, r'Lp_TOC_Replace\s+workRng\s*,\s*"\^t"\s*,\s*" "',
                           'the pass that turns every tab into a space')
    style = first(lines, r'paraRange\.Style\s*=\s*"TOC 1"', 'the line that applies TOC 1')
    second_tab = first(lines, r"If\s+Lp_TOC_Tab_Falls_Short\(\s*tabRng\s*\)\s+Then\s+"
                              r"tabRng\.InsertAfter\s+vbTab",
                       "the second tab after a short title")
    ref_pages = first(lines, r"paraRange\.InsertBefore\s+Chr\(160\)",
                      "the reference page pass")
    home = first(lines, r"homeRng\.FormattedText\s*=", "the assignment that brings the text home")

    assert tabs_to_spaces < second_tab, (
        "Every tab must become a space BEFORE the second tab is added, or running the macro "
        "twice stacks them")
    assert style < second_tab, "The second tab must be decided AFTER the TOC style is applied"
    assert second_tab < ref_pages, (
        "The second tab must go in BEFORE the reference page passes, which move every position")
    assert second_tab < home, (
        "The second tab must go in in the scratch document, before the text comes home - in the "
        "book each one is another Ctrl+Z")


def test_it_works_from_the_bottom_up_in_the_scratch_document(repo_root):
    _, lines = macro_lines(repo_root)
    second_tab = first(lines, r"Lp_TOC_Tab_Falls_Short\(\s*tabRng\s*\)",
                       "the second tab after a short title")
    loop = "\n".join(lines[max(0, second_tab - 3):second_tab + 1])
    assert re.search(r"For\s+i\s*=\s*tabCount\s+To\s+1\s+Step\s+-1", loop), (
        "The short-title pass must walk the tabs from the last to the first")
    assert re.search(r"Set\s+tabRng\s*=\s*tempDoc\.Range\(", loop), (
        "The short-title pass must measure in tempDoc, the hidden scratch document")


def test_the_tab_positions_are_recorded_where_the_tab_is_made(repo_root):
    _, lines = macro_lines(repo_root)
    made = first(lines, r"paraRange\.Characters\(matchStart\)\.Text\s*=\s*vbTab",
                 "the line that writes the page number's tab")
    recorded = first(lines, r"tabStarts\(tabCount\)\s*=\s*paraRange\.start\s*\+\s*matchStart\s*-\s*1",
                     "the record of where each tab is")
    assert made < recorded < made + 8, "Each tab's position must be recorded where it is written"


def test_the_helpers_take_everything_by_value(repo_root):
    procs, _ = macro_lines(repo_root)
    for name in ("Lp_TOC_Tab_Stops_At_Indent", "Lp_TOC_Tab_Falls_Short"):
        assert name in procs, f"{name} is missing"
        header = " ".join(l.rstrip(" _") for l in procs[name][:5])
        header = header[:header.index(")") + 1]
        params = header[header.index("(") + 1:-1].split(",")
        for p in params:
            assert p.strip().startswith("ByVal "), f"{name}: {p.strip()} must be ByVal"

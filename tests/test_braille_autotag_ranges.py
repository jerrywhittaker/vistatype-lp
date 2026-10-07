"""Braille AutoTag tags the same page ranges as large print.

Reported by Jerry 10/6/2026 on 3.0.540: braille AutoTag Ref Pages did not tag 16-17, 22B-23B,
A27B-A29B or AA30BB-AA33BB, which large print tags. Jerry: "there should be no difference between
lp and braille for the autotag". Two gaps, both measured on vistabuild:

  * Dx_Merge_Adjacent_Pg_Numbers joined ranges on consecutive lines into one - the first one's
    head and the last one's tail, 16-A29B - which no tagging pass accepts. From 10/7/2026 a line
    that already holds a range is not a bare page number (Dx_Pg_Text_Is_Range), so it is never
    merged into or from.
  * A space round the hyphen, 16 - 17, stopped every range pass. Large print closes up every
    hyphen in the book first (Lp_Fix_Hyphen_Errors); braille now closes up page-number lines only
    (Dx_Close_Up_Pg_Range_Hyphens), before the merge, so prose in a braille book is not touched.

What the two pure helpers return is tested in tests/vba/TestPageNumbers.bas. This checks the
wiring in the source: the sub works on a Document, which hangs the headless runner.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

AUTOTAG = "Dx_AutoTag_Page_Numbers"


def _code(procs, name):
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return [strip_comment(c) for c in procs[name]]


def _first_line(code, pattern):
    return next((i for i, c in enumerate(code) if re.search(pattern, c)), None)


def test_autotag_closes_up_page_ranges_before_the_merge(repo_root):
    code = _code(procedures(module_text(repo_root)), AUTOTAG)
    close_up = _first_line(code, r"\bDx_Close_Up_Pg_Range_Hyphens\s+ActiveDocument\b")
    merge = _first_line(code, r"\bDx_Merge_Adjacent_Pg_Numbers\s+ActiveDocument\b")
    assert close_up is not None, (
        f"{AUTOTAG} must close up spaced page ranges (16 - 17) - no range pass accepts them.")
    assert merge is not None, f"{AUTOTAG} no longer merges adjacent page numbers"
    assert close_up < merge, "the close-up must come before the merge, so the merge sees a range"


def test_braille_autotag_leaves_prose_hyphens_alone(repo_root):
    code = _code(procedures(module_text(repo_root)), AUTOTAG)
    assert _first_line(code, r"\bLp_Fix_Hyphen_Errors\b") is None, (
        f"{AUTOTAG} must not run Lp_Fix_Hyphen_Errors: it closes up every hyphen in the book, "
        "prose included. Page-number lines only (Dx_Close_Up_Pg_Range_Hyphens).")


def test_a_range_is_not_a_bare_page_number(repo_root):
    code = _code(procedures(module_text(repo_root)), "Dx_Is_Bare_Pg_Number_Paragraph")
    assert _first_line(code, r"If\s+Dx_Pg_Text_Is_Range\(\s*s\s*\)\s+Then\s+Exit\s+Function") is not None, (
        "Dx_Is_Bare_Pg_Number_Paragraph must turn a range away, or the merge joins 16-17 / 22B-23B "
        "/ A27B-A29B into 16-A29B, which nothing tags.")


def test_the_new_helpers_take_byval(repo_root):
    procs = procedures(module_text(repo_root))
    for name in ("Dx_Pg_Text_Is_Range", "Dx_Is_Pg_Range_Part", "Dx_Closed_Up_Pg_Range",
                 "Dx_Close_Up_Pg_Range_Hyphens"):
        header = procs[name][0] if name in procs else ""
        assert header, f"{name} is missing"
        params = re.search(r"\((.*)\)", header).group(1).split(",")
        assert all(re.match(r"^\s*ByVal\b", p, re.IGNORECASE) for p in params), (
            f"{name} must take its parameters ByVal: {header.strip()}")

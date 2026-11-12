"""Word equations, superscripts and subscripts are enlarged in LARGE PRINT ONLY (Jerry, 11/12/2026).

Lp_Size_Equations sets every equation so its characters print at the book's size, its scripts so
they do too, and an in-line equation's fractions written across. It must be called from the LP
attach alone - never from braille, the DAISY/NIMAS converter or an ordinary document - and after
every pass that lays one size over the whole book, or that pass undoes it. Tested for real in Word
by tests/vba/TestEquationSize.bas.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment


def test_called_only_from_the_lp_attach(repo_root):
    procs = procedures(module_text(repo_root))
    callers = [name for name, lines in procs.items() if name != "Lp_Size_Equations"
               and any(re.search(r"\bLp_Size_Equations\b", strip_comment(l)) for l in lines)]
    assert callers == ["Lp_Attach_The_Template"], callers


def test_runs_after_every_whole_book_size(repo_root):
    code = [strip_comment(l) for l in procedures(module_text(repo_root))["Lp_Attach_The_Template"]]

    def at(pattern):
        hits = [i for i, l in enumerate(code) if re.search(pattern, l)]
        assert hits, pattern
        return hits

    size = at(r"\bLp_Size_Equations\b")[0]
    for before in (r"\bSh_Set_Whole_Document_Font\b", r"\bLp_Normalize_Styles\b",
                   r"\bLp_Restore_Style_Sizes\b"):
        assert max(at(before)) < size, f"{before} runs after Lp_Size_Equations and would undo it"

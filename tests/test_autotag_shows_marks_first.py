"""Both AutoTags turn formatting marks on before their first tagging pass.

Reported by Jerry 10/7/2026 on 3.0.541: braille AutoTag Ref Pages on a book with the Normal
template left 16-18, 22B-26B, A27B-A29B and AA30B-AA33BB untagged. Measured on vistabuild with
copies of his "Auto Tag Page Numbers.docx": on those four lines the paragraph mark itself is
formatted hidden, and Selection.Find cannot see a hidden paragraph mark while formatting marks and
hidden text are off - so every "^013...^013" pass missed them (10 tags instead of 14). With
formatting marks on, all 14 tagged, on Normal and on BANA alike.

Braille AutoTag turned formatting marks on only at its END. The letter configuration turns them
off, and until 10/6/2026 (issue #8) the BANA attach ran first and turned them on, which hid the
fault. Large print uses the same passes; its configuration normally turns them on, but not when
Sh_Config_Skip_Display is up - so it turns them on for the passes and puts them back after.

Not unit-testable in VBA: the macros work on the document, which hangs the headless runner
(tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

SHOW_ON = re.compile(r"\.(ShowAll|ShowHiddenText)\s*=\s*True\b", re.IGNORECASE)
FIRST_PASS = re.compile(r'\.Text\s*=\s*"\^013', re.IGNORECASE)


def _code(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return [strip_comment(c) for c in procs[name]]


def _first(code, pattern):
    return next((i for i, c in enumerate(code) if pattern.search(c)), None)


def _check(repo_root, name):
    code = _code(repo_root, name)
    on = _first(code, SHOW_ON)
    first = _first(code, FIRST_PASS)
    assert first is not None, f"{name} has no ^013 pass any more - update this test"
    assert on is not None and on < first, (
        f"{name} must turn formatting marks on (ShowAll = True) before its first ^013 Find pass - "
        "a page number whose paragraph mark is hidden is invisible to the passes otherwise.")


def test_braille_autotag_shows_marks_first(repo_root):
    _check(repo_root, "Dx_AutoTag_Page_Numbers")


def test_large_print_autotag_shows_marks_first(repo_root):
    _check(repo_root, "Lp_AutoTag_Page_Numbers")


def test_large_print_autotag_puts_the_marks_back(repo_root):
    code = _code(repo_root, "Lp_AutoTag_Page_Numbers")
    assert any(re.search(r"If\s+Not\s+showAllPrev\s+Then\s+\S*ShowAll\s*=\s*False", c,
                         re.IGNORECASE) for c in code), (
        "Lp_AutoTag_Page_Numbers must put formatting marks back as they were after the passes.")

"""AutoTag Ref Pages tags "letter, hyphen, letter" only when there is a hyphen.

Issue #35, found by reading 10/16/2026: in Lp_AutoTag_Page_Numbers the pass commented
"letter ,Hyphen, letter" had lost its hyphen. Its wildcard pattern was

    ^013([A-Za-z]{1,})([A-Za-z]{1,})^013

with the replacement ^p$pg\\1\\2^p, so any paragraph made only of two or more letters -
"Introduction", "Contents", "Yes" - could be tagged $pg. The braille copy of the pass in
Dx_AutoTag_Page_Numbers always had the hyphen (^045). Both must keep it.

This checks the source: the sub works on a Document, which hangs the headless runner.
"""
from test_toc_never_deletes_a_line import module_text, procedures

WITH_HYPHEN = '"^013([A-Za-z]{1,}^045)([A-Za-z]{1,})^013"'
WITHOUT_HYPHEN = '"^013([A-Za-z]{1,})([A-Za-z]{1,})^013"'


def _code(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return "\n".join(procs[name])


def _check(repo_root, name):
    code = _code(repo_root, name)
    assert WITHOUT_HYPHEN not in code, (
        f"{name}: the letter, hyphen, letter pass has lost its hyphen (^045) - it would tag "
        "any paragraph made only of letters, such as 'Introduction', as $pg (issue #35).")
    assert WITH_HYPHEN in code, (
        f"{name}: the letter, hyphen, letter pass {WITH_HYPHEN} is missing.")


def test_large_print_letter_hyphen_letter_pass_has_its_hyphen(repo_root):
    _check(repo_root, "Lp_AutoTag_Page_Numbers")


def test_braille_letter_hyphen_letter_pass_has_its_hyphen(repo_root):
    _check(repo_root, "Dx_AutoTag_Page_Numbers")

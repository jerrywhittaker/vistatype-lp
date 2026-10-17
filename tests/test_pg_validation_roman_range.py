"""Validate $pg Tags marks a roman numeral range as a likely problem.

Jerry, 10/17/2026: a range of roman numerals such as v-vi must be shown as a problem in the
validation list. Braille Manual Tag Ref Page refuses one (message 245) and stays that way, but
braille AutoTag Ref Pages' letter-hyphen-letter pass tags it - $pgv-vi[[*lec*]][[*i*]]vi - and
large print's pass does the same. Until this change the list checker (Sh_PgVal_Check, 3.0.583)
read v-vi as an ordinary two-page range and said nothing.

What the checker answers for each shape is tested in tests/vba/TestPgValCheck.bas, which needs
make vba-test on the build box. This checks the source, so the rule is held here as well:
Sh_PgVal_Parse reports "romanrange" for a roman range, Sh_PgVal_Check turns it into the line
"roman numeral range", and that line fits a large print column.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

WARNING = "roman numeral range"


def _code(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return [strip_comment(c) for c in procs[name]]


def test_parse_reports_a_roman_range(repo_root):
    code = _code(repo_root, "Sh_PgVal_Parse")
    hits = [c for c in code if '"romanrange"' in c or "romanrange" in c]
    assert hits, "Sh_PgVal_Parse never reports a roman numeral range (romanrange)"
    assert any(re.search(r'\bisRange\b.*\ba\(0\)\s*=\s*"R"', c) for c in hits), (
        "the roman range must be judged from the tag's own hyphen AND a roman reading - "
        "a single roman page (v) is valid and must not be reported")


def test_check_writes_the_roman_range_line(repo_root):
    code = _code(repo_root, "Sh_PgVal_Check")
    assert any(f'" {WARNING}"' in c for c in code), (
        f'Sh_PgVal_Check never writes "{WARNING}" for a roman range')
    assert any(re.search(r'InStr\(.*" romanrange "', c) for c in code), (
        "Sh_PgVal_Check must look for romanrange among the problem words, not compare the "
        "whole field - a range can be a mismatch as well")
    # The mark, a space and the words must fit a large print column (23 characters).
    assert 2 + len(WARNING) <= 23


def test_mismatch_is_still_found_beside_a_roman_range(repo_root):
    code = _code(repo_root, "Sh_PgVal_Check")
    assert not any(re.search(r'f\(6\)\s*=\s*"mismatch"', c) for c in code), (
        'Sh_PgVal_Check compares the whole problem field with "mismatch" - a roman range '
        "with a wrong continuation number would lose its mismatch line")


def test_the_vba_suite_covers_roman_ranges(repo_root):
    text = (repo_root / "tests" / "vba" / "TestPgValCheck.bas").read_bytes().decode("latin-1")
    for shape in ('"$pgv-vi"', '"$pgXII-XIV"', '"$pg[[*ii*]]v-vi[[*lec*]][[*i*]]vi"',
                  '"1=roman numeral range"', '"$pgword-word"'):
        assert shape in text, f"tests/vba/TestPgValCheck.bas has no case for {shape}"

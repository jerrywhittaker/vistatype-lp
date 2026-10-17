"""Braille Manual Tag Ref Page gives a page range the DBT continuation code, as AutoTag does.

Jerry, 10/17/2026: a paragraph "16-18" tagged with Manual Tag Ref Page should come out as
braille AutoTag Ref Pages makes it, $pg16-18[[*lec*]][[*i*]]18. Until then Manual Tag gave it
only $pg16-18.

AutoTag builds the code in its own Find passes (Dx_AutoTag_Page_Numbers), and those are left
alone. Manual Tag gets it from Dx_Pg_Range_Continuation, which mirrors them. These checks hold
the two together: the same code, the same range shapes, the code typed after the $pg tag and
before Dx_Set_DBT_Codes_Color_and_Style gives it the hidden DBT Code style, and a spaced hyphen
closed up first, as AutoTag closes it up. The helper's answers on real ranges are tested in
tests/vba/TestPageNumbers.bas. The sub itself is not unit-testable in VBA: it works on the
Selection, and anything that reaches a Document hangs the headless runner.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

MANUAL = "Dx_Manual_Tag_with_Dollar_pg"
AUTOTAG = "Dx_AutoTag_Page_Numbers"
HELPER = "Dx_Pg_Range_Continuation"
SHAPE = "Dx_Pg_Range_Side_Shape"
CODE = "[[*lec*]][[*i*]]"

FIND_TEXT = re.compile(r'^\s*\.Text\s*=\s*"(.*)"\s*$')
REPL_TEXT = re.compile(r'^\s*\.Replacement\.Text\s*=\s*"(.*)"\s*$')
RANGE_PATTERN = re.compile(r"^\^013\((.*)\^045\)\((.*)\)\^013$")
TOKENS = {"[A-Za-z]{1,}": "L", "[0-9]{1,}": "D"}


def _procs(repo_root):
    procs = procedures(module_text(repo_root))
    for name in (MANUAL, AUTOTAG, HELPER, SHAPE):
        assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return procs


def _shape(pattern):
    """'[A-Za-z]{1,}[0-9]{1,}' -> 'LD', as Dx_Pg_Range_Side_Shape names it."""
    shape = ""
    rest = pattern
    while rest:
        for token, kind in TOKENS.items():
            if rest.startswith(token):
                shape += kind
                rest = rest[len(token):]
                break
        else:
            raise AssertionError(f"AutoTag range pattern not understood: {pattern!r}")
    return shape


def _autotag_range_shapes(procs):
    """The range shapes AutoTag's Find passes give the continuation code."""
    shapes = set()
    last_find = None
    for line in procs[AUTOTAG]:
        m = FIND_TEXT.match(line)
        if m:
            last_find = m.group(1)
            continue
        m = REPL_TEXT.match(line)
        if m and CODE in m.group(1) and last_find is not None:
            pm = RANGE_PATTERN.match(last_find)
            assert pm, f"AutoTag codes a pattern that is not one hyphen between two sides: {last_find!r}"
            left, right = _shape(pm.group(1)), _shape(pm.group(2))
            assert left == right, f"AutoTag codes a range whose sides differ: {last_find!r}"
            shapes.add(left)
    return shapes


def _index(lines, pattern):
    for i, line in enumerate(lines):
        if re.search(pattern, line, re.IGNORECASE):
            return i
    return -1


def test_manual_tag_adds_the_code_after_the_tag_and_before_the_style_pass(repo_root):
    code = _procs(repo_root)[MANUAL]
    tag = _index(code, r'Selection\.TypeText\s+Text:="\$pg"')
    call = _index(code, rf"=\s*{HELPER}\s*\(")
    typed = _index(code, r"Selection\.TypeText\s+Text:=contCode\b")
    style = _index(code, r"Dx_Set_DBT_Codes_Color_and_Style")
    assert call >= 0, f"{MANUAL} does not ask {HELPER} for the range's continuation code"
    assert typed >= 0, f"{MANUAL} does not type the continuation code"
    assert tag >= 0 and style >= 0
    assert tag < typed < style, (
        f"{MANUAL} must type the continuation code after the $pg tag and before "
        f"Dx_Set_DBT_Codes_Color_and_Style, or the code is not given the hidden DBT Code style")


def test_manual_tag_closes_up_a_spaced_hyphen_first(repo_root):
    code = _procs(repo_root)[MANUAL]
    closed = _index(code, r"=\s*Dx_Closed_Up_Pg_Range\s*\(\s*Possible_LC_Roman\s*\)")
    call = _index(code, rf"=\s*{HELPER}\s*\(\s*Possible_LC_Roman\s*\)")
    assert closed >= 0, f"{MANUAL} does not close up 16 - 18 as AutoTag does"
    assert 0 <= closed < call


def test_the_helper_writes_the_same_code_as_autotag(repo_root):
    procs = _procs(repo_root)
    repl = [m.group(1) for m in map(REPL_TEXT.match, procs[AUTOTAG]) if m and CODE in m.group(1)]
    assert repl, f"{AUTOTAG} no longer writes {CODE} - this test needs looking at"
    for r in repl:
        assert r == r"^p$pg\1\2" + CODE + r"\2^p", f"AutoTag's continuation tag changed: {r!r}"
    assert any(f'"{CODE}" & rightPart' in line for line in procs[HELPER]), (
        f"{HELPER} must write {CODE} followed by the side after the hyphen, as AutoTag does")


def test_the_helper_codes_the_same_shapes_as_autotag(repo_root):
    procs = _procs(repo_root)
    autotag = _autotag_range_shapes(procs)
    assert autotag == {"D", "LD", "DL", "LDL", "L"}, f"AutoTag's range passes changed: {autotag}"
    cases = [line for line in procs[HELPER] if re.match(r"\s*Case\s+\"", line)]
    assert len(cases) == 1, f"{HELPER} should have one Case line listing the shapes: {cases}"
    helper = set(re.findall(r'"([A-Z]+)"', cases[0]))
    assert helper == autotag, (
        f"{HELPER} codes {sorted(helper)}; AutoTag codes {sorted(autotag)} - they must agree")


def test_the_helpers_take_their_arguments_byval(repo_root):
    procs = _procs(repo_root)
    for name in (HELPER, SHAPE):
        head = procs[name][0]
        assert re.search(r"\(\s*ByVal\s+\w+\s+As\s+String\s*\)", head), (
            f"{name} must take its argument ByVal: {head.strip()}")


# Retagging, 10/17/2026 (review): the old code is hidden, and with hidden text not shown
# Selection.Text does not see it, so the line read "16-18" and got the code a second time.
# Sh_Remove_DollarPG_For_Retag, shared with large print, takes off only the $pg.

REMOVER = "Dx_Pg_Continuation_Code_Start"


def test_manual_tag_takes_off_an_old_code_before_it_reads_the_line(repo_root):
    procs = procedures(module_text(repo_root))
    assert REMOVER in procs, f"{REMOVER} is missing from LPandBrlMacros.bas"
    code = procs[MANUAL]
    retag = _index(code, r'"Sh_Remove_DollarPG_For_Retag"')
    hidden = _index(code, r"\.TextRetrievalMode\.IncludeHiddenText\s*=\s*True")
    found = _index(code, rf"=\s*{REMOVER}\s*\(")
    delete = _index(code, r"\bcodeRng\.Delete\b")
    read = _index(code, r"Possible_LC_Roman\s*=.*\bSelection\.Text\b")
    assert hidden >= 0, (
        f"{MANUAL} must read the line with IncludeHiddenText - the old code is hidden")
    assert found >= 0 and delete >= 0, f"{MANUAL} does not take off an old continuation code"
    assert 0 <= retag < hidden < found < delete < read, (
        f"{MANUAL} must take off the old code after the $pg comes off and before the line is read")


def test_the_remover_takes_its_argument_byval(repo_root):
    head = procedures(module_text(repo_root))[REMOVER][0]
    assert re.search(r"\(\s*ByVal\s+\w+\s+As\s+String\s*\)", head), head.strip()


def test_the_shared_retag_is_left_as_it_was(repo_root):
    # Large print uses it too; the continuation code is braille-only and is handled in Manual Tag.
    code = procedures(module_text(repo_root))["Sh_Remove_DollarPG_For_Retag"]
    assert not any("[[*lec*]]" in line for line in code), (
        "Sh_Remove_DollarPG_For_Retag is shared with large print - keep the braille code out of it")

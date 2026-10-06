"""Selected Cleanup's Italics to Dashed Underline stays inside the selection.

Issue #4, 10/6/2026. Every find-and-replace pass in Lp_Italics_To_Dashed_Underline used
.Wrap = wdFindContinue, so a pass that reached the end of the selection carried on through the
rest of the book. Measured on vistabuild 10/6/2026: with the middle of three paragraphs selected,
all three italic words became dashed underline.

The first cure, wdFindStop on the Replace All passes (built as 3.0.532), was NOT enough. Jerry
selected one paragraph of an all-italic document and it, and every paragraph below it, was dashed.
Measured on vistabuild 10/6/2026: when everything in the range already matches the Find's
formatting, Word treats the range as already found and searches on from its end, so Replace All
runs to the end of the book. A Range made from the selection does the same.

So with selectionOnly the 12 passes go through Lp_Dash_Italics_Pass_In_Range, which starts each
Find from a collapsed point and cuts every hit off at the end of the selection, and uses no Replace
All. File Cleanup (Lp_Fix_Common_File_Errors) and Lp_Normalize_Styles call the sub with a
collapsed cursor and rely on the wdFindContinue wrap of the Replace All passes to cover the book.

Selected Cleanup also checks the LP template is attached before it changes anything (Jerry,
10/6/2026).

Not unit-testable in VBA: the sub works on the Selection, and anything that reaches a Document
hangs the headless runner (tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

SUB = "Lp_Italics_To_Dashed_Underline"
HELPER = "Lp_Dash_Italics_Pass_In_Range"
PASSES = 12          # 13 when issue #4 was filed; the bold pass came out 9/20/2026
FORM = "Lp_Selected_Cleanup_Form.frm"
WHOLE_BOOK_CALLERS = ("Lp_Fix_Common_File_Errors", "Lp_Normalize_Styles")

WRAP_CONTINUE = re.compile(r"\.Wrap\s*=\s*wdFindContinue\b", re.IGNORECASE)
WRAP_ANY = re.compile(r"\.Wrap\s*=", re.IGNORECASE)
REPLACE_ALL = re.compile(r"\bwdReplaceAll\b", re.IGNORECASE)
CALL = re.compile(r"\b" + SUB + r"\b", re.IGNORECASE)
HELPER_CALL = re.compile(r"^\s*(?:Call\s+)?" + HELPER + r"\b", re.IGNORECASE)
PASSES_TRUE = re.compile(r"selectionOnly\s*:=\s*True\b|\b" + SUB + r"\b\W+True\b", re.IGNORECASE)


def _sub(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return procs[name]


def _header(code):
    """The declaration line, with any ' _' continuation lines joined on."""
    parts = []
    for c in code:
        parts.append(c.rstrip())
        if not parts[-1].endswith("_"):
            break
        parts[-1] = parts[-1][:-1]
    return " ".join(parts)


def _selection_branch(code):
    """The code lines of the If selectionOnly Then ... End If block."""
    start = next((i for i, c in enumerate(code)
                  if re.match(r"^\s*If\s+selectionOnly\s+Then\s*$", c, re.IGNORECASE)), None)
    assert start is not None, f"{SUB} has no 'If selectionOnly Then' block"
    end = next(i for i in range(start, len(code)) if re.match(r"^\s*End\s+If\b", code[i], re.IGNORECASE))
    return code[start:end + 1], end


def test_selection_path_walks_the_helper_and_never_replaces_all(repo_root):
    code = _sub(repo_root, SUB)
    branch, _ = _selection_branch(code)
    calls = [c for c in branch if HELPER_CALL.search(c)]
    assert len(calls) == PASSES, (
        f"{SUB}'s selectionOnly block calls {HELPER} {len(calls)} times; expected {PASSES}, one "
        "for each Replace All pass. If a pass was added or taken out on purpose, change PASSES here "
        "and add or take out the matching call on both paths.")
    assert not [c for c in branch if REPLACE_ALL.search(c)], (
        f"{SUB}'s selectionOnly block uses Replace All - on an all-italic selection Word searches "
        "on from the end of the selection and changes the rest of the book (issue #4).")
    assert any(re.match(r"^\s*Exit\s+Sub\b", c, re.IGNORECASE) for c in branch), (
        f"{SUB}'s selectionOnly block must Exit Sub, or the whole-book Replace All passes run too.")


def test_whole_book_passes_wrap_and_come_after_the_selection_path(repo_root):
    code = _sub(repo_root, SUB)
    _, branch_end = _selection_branch(code)
    before = [c.strip() for c in code[:branch_end] if REPLACE_ALL.search(c)]
    assert not before, f"{SUB} runs a Replace All before the selectionOnly block: {before}"
    wraps = [c.strip() for c in code if WRAP_ANY.search(c)]
    continuing = [c for c in wraps if WRAP_CONTINUE.search(c)]
    assert len(continuing) == PASSES and len(wraps) == PASSES, (
        f"{SUB}: the {PASSES} whole-book passes must each use .Wrap = wdFindContinue, so File "
        f"Cleanup and Normalize Styles cover the book from a collapsed cursor. Found: {wraps}")


def test_the_sub_and_helper_take_byval(repo_root):
    assert re.search(r"\(\s*Optional\s+ByVal\s+selectionOnly\s+As\s+Boolean\s*=\s*False\s*\)",
                     _header(_sub(repo_root, SUB)), re.IGNORECASE), (
        f"{SUB} must take Optional ByVal selectionOnly As Boolean = False - ByVal, as every "
        "helper here does, and False so File Cleanup and Normalize Styles keep the whole book.")
    header = _header(_sub(repo_root, HELPER))
    params = re.search(r"\((.*)\)", header).group(1).split(",")
    assert len(params) == 6 and all(re.match(r"^\s*ByVal\b", p, re.IGNORECASE) for p in params), (
        f"{HELPER} must take its 6 parameters ByVal: {header}")


def test_helper_stops_at_the_end_of_the_range(repo_root):
    body = " ".join(c.strip() for c in _sub(repo_root, HELPER))
    assert not REPLACE_ALL.search(body), f"{HELPER} must not use Replace All (issue #4)."
    assert re.search(r"\.Wrap\s*=\s*wdFindStop\b", body, re.IGNORECASE), (
        f"{HELPER} must use wdFindStop, or a Find past the last hit wraps to the top of the book.")
    assert re.search(r"If\s+hit\.Start\s*>=\s*rangeEnd\s+Then\s+Exit\s+Do", body, re.IGNORECASE), (
        f"{HELPER} must stop at a hit that starts at or past the end of the range.")
    assert re.search(r"If\s+hit\.End\s*>\s*rangeEnd\s+Then\s+hit\.End\s*=\s*rangeEnd", body,
                     re.IGNORECASE), (
        f"{HELPER} must cut a hit off at the end of the range - a Find for italic in an italic "
        "paragraph returns the whole run, past the selection.")
    assert re.search(r"Set\s+hit\s*=\s*within\.Duplicate\s+hit\.SetRange\s+nextStart\s*,\s*nextStart\b",
                     body, re.IGNORECASE), (
        f"{HELPER} must start each Find from a collapsed point - a Find on a range that already "
        "matches searches on from its end.")
    # vba-review, 10/6/2026: ActiveDocument.Range is always the main text, so a selection in a
    # footnote, text box, header or footer would have dashed the same positions in the main text.
    assert not re.search(r"ActiveDocument\.Range\b", body, re.IGNORECASE), (
        f"{HELPER} must build its ranges from the range it is given (within.Duplicate), not "
        "ActiveDocument.Range - that is always the main text, whatever story the selection is in.")


def test_selected_cleanup_passes_true(repo_root):
    text = (repo_root / "src" / "forms" / FORM).read_bytes().decode("latin-1")
    calls = [strip_comment(line) for line in text.splitlines() if CALL.search(strip_comment(line))]
    assert calls, f"{FORM} no longer calls {SUB}"
    for c in calls:
        assert PASSES_TRUE.search(c), (
            f"{FORM} calls {SUB} without selectionOnly True - it would change italics outside "
            f"the selection (issue #4): {c.strip()}")


def test_whole_book_callers_do_not_pass_true(repo_root):
    for caller in WHOLE_BOOK_CALLERS:
        calls = [c for c in _sub(repo_root, caller) if CALL.search(c)]
        assert calls, f"{caller} no longer calls {SUB}"
        for c in calls:
            assert not PASSES_TRUE.search(c), (
                f"{caller} passes selectionOnly True to {SUB} - it runs from a collapsed "
                f"cursor and needs the wrap to cover the whole book: {c.strip()}")


def test_selected_cleanup_checks_the_lp_template_first(repo_root):
    # Jerry, 10/6/2026: Italics to Dashed Underline from Selected Cleanup checks the LP template
    # is attached before it changes anything.
    text = (repo_root / "src" / "forms" / FORM).read_bytes().decode("latin-1")
    lines = [strip_comment(line) for line in text.splitlines()]
    branch = next(i for i, c in enumerate(lines) if re.search(r"ElseIf\s+ItalicsToDashedUnderline\b", c))
    call = next(i for i in range(branch, len(lines)) if CALL.search(lines[i]))
    guard = [i for i in range(branch, call) if re.search(r"\bLp_Is_Lp_Template_Attached\b", lines[i])]
    assert guard, (
        f"{FORM}: the Italics to Dashed Underline option must call Lp_Is_Lp_Template_Attached "
        f"before {SUB}.")

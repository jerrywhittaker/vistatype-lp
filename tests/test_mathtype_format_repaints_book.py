"""MathType's Format Equations must not leave the converted book showing black.

Jerry, 10/25/2026 (issue #39): Convert DAISY/NIMAS to Word with "MathType equations" chosen,
and a blank Document1 open at the start. When the bar reached "MathType is formatting the
equations", Word showed Document2 - the converted book - with its document area black.

Sh_MathType_Format_Equations has to bring the book forward with targetDoc.Activate, because
MathType's Format Equations works only on ActiveDocument. But the converter switched the screen
off before Documents.Add made the book, so the book had never been drawn, and Word does not
repaint a window brought forward with ScreenUpdating off.

The first cure, Application.ScreenRefresh straight after the Activate with the screen still
off, did NOT paint it: on 3.0.608 the book was still black and MathType drew its equation
outlines on the black (measured by Jerry, 10/25/2026). The cure now is to switch the screen ON
and refresh it after the Activate, immediately before the Format Equations call - which
switches the screen off itself as it starts - and to switch it off again after Format
Equations, so the font pass and the repaginate still run with the screen off.

The painting itself cannot be tested (it needs Word on a screen), so these are checks on the
source of Sh_MathType_Format_Equations.
"""
from test_toc_never_deletes_a_line import module_text, procedures

FUNC = "Sh_MathType_Format_Equations"
CALL = 'application.run "mtformatequations.doformatequations"'


def _statements(repo_root):
    """The function's code lines, comments and blank lines taken out, trimmed."""
    lines = procedures(module_text(repo_root))[FUNC]
    return [l.strip() for l in lines if l.strip() and not l.strip().startswith("'")]


def _key(line):
    return " ".join(line.lower().split())


def _activate_and_call(code):
    acts = [i for i, l in enumerate(code) if _key(l) == "targetdoc.activate"]
    assert len(acts) == 1, f"expected one targetDoc.Activate in {FUNC}, found {len(acts)}"
    calls = [i for i, l in enumerate(code) if _key(l).startswith(CALL)]
    assert len(calls) == 1, f"expected one Format Equations call in {FUNC}, found {len(calls)}"
    assert acts[0] < calls[0], f"{FUNC}: the Activate must come before Format Equations"
    return acts[0], calls[0]


def test_screen_on_and_refreshed_just_before_format_equations(repo_root):
    code = _statements(repo_root)
    _, c = _activate_and_call(code)
    assert c >= 2 and _key(code[c - 2]) == "application.screenupdating = true" \
        and _key(code[c - 1]) == "application.screenrefresh", (
        f"{FUNC}: the two statements before Format Equations must be "
        f"Application.ScreenUpdating = True then Application.ScreenRefresh, or the book stays "
        f"black behind the bar (issue #39; a refresh with the screen off did not paint it, "
        f"measured by Jerry on 3.0.608); found {code[max(c - 2, 0):c]!r}")


def test_no_refresh_with_the_screen_off(repo_root):
    code = _statements(repo_root)
    a, c = _activate_and_call(code)
    between = [_key(l) for l in code[a + 1:c - 2]]
    bad = [l for l in between if "screenrefresh" in l or "screenupdating" in l]
    assert not bad, (
        f"{FUNC}: between the Activate and the screen going on, nothing may touch the screen - "
        f"the early exits there must leave it off for the font pass and the repaginate: {bad}")


def test_the_screen_is_not_touched_before_the_activate(repo_root):
    code = _statements(repo_root)
    a, _ = _activate_and_call(code)
    before = [l for l in code[:a] if "screenupdating" in l.lower()]
    assert not before, (
        f"{FUNC} must not change ScreenUpdating before the Activate - the font pass and the "
        f"repaginate after it rely on the screen staying off: {before}")


def test_screen_off_again_after_format_equations(repo_root):
    code = _statements(repo_root)
    _, c = _activate_and_call(code)
    after = [_key(l) for l in code[c + 1:]]
    end = after.index("exit function") if "exit function" in after else len(after)
    assert "application.screenupdating = false" in after[:end], (
        f"{FUNC}: ScreenUpdating must be set False again after Format Equations (which ends "
        f"by switching the screen on), so the font pass and the repaginate run with it off")

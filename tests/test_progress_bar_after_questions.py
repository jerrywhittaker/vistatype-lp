"""Braille Selected Cleanup (333): the progress bar waits until the questions are answered.

Jerry, 10/12/2026: "The progress bar should not appear until the selection is made." Two of the
thirteen passes ask a question - Remove Bullets shows "Choose paragraph style" (329) and Replace
Manual Line Breaks shows "Choose Replacement Character" (362) - and both were asked from inside
the loop, with the bar already on screen behind them. The form now asks both before
Sh_Progress_Open and hands the answers to the passes, which then ask nothing.

Also: Sh_Replace_Manual_Line_Break ended a Cancel with End, which closes any box that stays open.
It leaves with Exit Sub now.

Not unit-testable in VBA: a UserForm blocks the headless runner. So this is a check on the source.
"""
import re

from test_selection_cleanup_reselects import _okay
from test_toc_never_deletes_a_line import procedures
from test_no_end_while_box_open import module_text


def test_both_questions_come_before_the_bar(repo_root):
    lines = _okay(repo_root)
    bar = next(i for i, l in enumerate(lines) if "Sh_Progress_Open" in l)
    for show in ("Dx_Bullet_Removal_Form.Show", "Sh_Space_Or_Para_Form.Show"):
        at = [i for i, l in enumerate(lines) if show in l]
        assert at, f"Selected Cleanup no longer asks {show} itself"
        assert at[0] < bar, f"Selected Cleanup opens the progress bar before {show}"


def test_the_passes_get_the_answers(repo_root):
    body = "\n".join(_okay(repo_root))
    assert re.search(r'Application\.Run\s+"Dx_Remove_Bullets",\s*bulletStyle,\s*bulletHyper', body)
    assert re.search(r'Application\.Run\s+"Sh_Replace_Manual_Line_Break",\s*breakAnswer', body)


def test_remove_bullets_does_not_ask_again_when_told(repo_root):
    body = procedures(module_text(repo_root))["Dx_Remove_Bullets"]
    assert re.match(r"\s*Sub Dx_Remove_Bullets\(Optional ByVal style As String, "
                    r"Optional ByVal hyper As String\)", body[0])
    show = next(n for n, c in enumerate(body) if "Dx_Bullet_Removal_Form.Show" in c)
    near = "\n".join(body[show - 5:show])
    assert re.search(r'If style <> "" Then', near) and re.search(r"^\s*Else\s*$", near, re.M), (
        "Dx_Remove_Bullets shows the box even when the caller already asked")


def test_line_break_cancel_has_no_end(repo_root):
    body = procedures(module_text(repo_root))["Sh_Replace_Manual_Line_Break"]
    code = [c.split("'")[0] for c in body]
    assert not any(re.match(r"^\s*End\s*$", c) for c in code), (
        "Sh_Replace_Manual_Line_Break has a bare End - it would close every box that stays open")


def test_selected_cleanup_with_nothing_selected_has_no_end(repo_root):
    """Issue #33, 10/12/2026: Okay with nothing selected said 317 and ran End, closing every box
    that stays open. It stops with Exit Sub now."""
    from test_toc_box_stays_open import form_code
    code = [c.split("'")[0] for c in form_code(repo_root, "Dx_Selected_Cleanup_Form")]
    assert not any(re.match(r"^\s*End\s*$", c) for c in code), (
        "Dx_Selected_Cleanup_Form has a bare End again - it would close every box that stays open")


def test_unticking_remove_hyperlinks_clears_the_answer(repo_root):
    """Issue #34, 10/12/2026: every click set "Remove_Hyper", so ticking then unticking still
    turned the links into plain text. The answer follows the box now."""
    from test_red_x_acts_like_cancel import form_procedures
    body = "\n".join(form_procedures(repo_root, "Dx_Bullet_Removal_Form")["Remove_Hyperlinks_Click"])
    assert re.search(r"If\s+Remove_Hyperlinks\s+Then", body), (
        "Remove_Hyperlinks_Click no longer looks at the box before setting the answer")
    assert re.search(r'Dx_GP_String_1\s*=\s*""', body), (
        "Unticking Remove Hyperlinks does not clear the answer")

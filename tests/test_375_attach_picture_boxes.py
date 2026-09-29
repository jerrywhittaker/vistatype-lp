"""Resize Pictures in a Selected Range (375): the two "Attach the picture" tick boxes.

Jerry, 9/29/2026: the old box, "Move the picture to the paragraph which follows it.", is
reworded "Attach the picture to the paragraph which follows it.", and a second box, "Attach the
picture to the paragraph which precedes it.", replaces the paragraph mark in front of the
picture with a space. Both may be ticked.

The decision about WHICH marks may go is tested in tests/vba/TestSamePicture.bas. What is held
here is the wiring the VBA runner cannot reach, because it lives in a UserForm or touches a
Range: the wording and letters on the box, that the new box is read and handed on, and that
the mark after a picture is taken before the mark in front of it.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures

FORM = "Lp_Same_Pic_Range_Form"


def initialize(repo_root):
    code = "\n".join(form_code(repo_root, FORM))
    m = re.search(r"Sub UserForm_Initialize\(\)(.*?)End Sub", code, re.S)
    assert m, "no UserForm_Initialize on the 375 box"
    return m.group(1)


def test_the_two_boxes_say_what_jerry_asked(repo_root):
    init = initialize(repo_root)
    assert ('JoinNextParaButton.Caption = '
            '"Attach the picture to the paragraph which follows it."') in init
    assert ('JoinPrevParaButton.Caption = '
            '"Attach the picture to the paragraph which precedes it."') in init
    assert "Move the picture" not in init


def test_no_two_controls_share_a_letter(repo_root):
    """Two controls on one letter make Alt cycle between them instead of pressing either. The
    letters set in code win over the .frx, so those are the ones read; the rest are the .frx's,
    measured on the build box 9/29/2026."""
    init = initialize(repo_root)
    in_code = dict(re.findall(r'(\w+)\.Accelerator = "(\w)"', init))
    assert in_code["JoinNextParaButton"].upper() == "F"
    assert in_code["JoinPrevParaButton"].upper() == "P"
    letters = {"LeftAlignButton": "L", "CenterButton": "C", "ApplyButton": "A",
               "DoneButton": "D", "LastSizeButton": "U"}
    letters.update(in_code)
    upper = [v.upper() for v in letters.values()]
    assert len(upper) == len(set(upper)), f"two controls share a letter: {letters}"
    # O and C are Okay and Cancel everywhere else; C is already Center's here.
    for name in ("JoinNextParaButton", "JoinPrevParaButton"):
        assert letters[name].upper() not in ("O", "C")


def test_the_new_box_is_read_and_handed_on(repo_root):
    procs = procedures(module_text(repo_root))
    reader = "\n".join(procs["Lp_Rst_Join_Prev_Wanted"])
    assert "Lp_Same_Pic_Range_Form.JoinPrevParaButton.Value" in reader
    run = "\n".join(procs["Lp_Rst_Run"])
    assert re.search(r"joinPrev = Lp_Rst_Join_Prev_Wanted\(\)", run)
    assert re.search(r"Lp_Same_Pic_Apply matches, ref, refW, refH, alignMode, joinNext, joinPrev,",
                     run)
    # the count on the box's own line is shown when EITHER box is ticked
    assert "If joinNext Or joinPrev Then" in run


def test_the_mark_after_goes_before_the_mark_in_front(repo_root):
    """For the model picture and for every copy, and each only when its own box is ticked."""
    apply_ = procedures(module_text(repo_root))["Lp_Same_Pic_Apply"]
    nexts = [n for n, c in enumerate(apply_) if "Lp_Same_Pic_Join_Next_Para(" in c]
    prevs = [n for n, c in enumerate(apply_) if "Lp_Same_Pic_Join_Prev_Para(" in c]
    assert len(nexts) == 2 and len(prevs) == 2, (nexts, prevs)
    for nx, pv in zip(nexts, prevs):
        assert nx < pv
        assert apply_[nx - 1].strip() == "If joinNext Then"
        assert apply_[pv - 1].strip() == "If joinPrev Then"


def test_the_join_in_front_asks_the_tested_decision(repo_root):
    """The document-facing half only gathers facts; the decision is the one the VBA suite tests,
    and nothing is replaced unless it says yes."""
    body = procedures(module_text(repo_root))["Lp_Same_Pic_Join_Prev_Para"]
    text = "\n".join(body)
    ask = next(n for n, c in enumerate(body) if "Lp_Same_Pic_Prev_Mark_May_Go(" in c)
    write = next(n for n, c in enumerate(body) if 'before.Text = " "' in c)
    assert ask < write
    assert "Then Exit Function" in "".join(body[ask:ask + 2])
    # no range is reached back past the start of the story
    assert re.search(r"If picStart > 0 Then", text)

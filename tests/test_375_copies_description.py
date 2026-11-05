"""Resize Pictures in a Selected Range (375): the clicked picture's description goes to every
copy it resizes (Jerry, 11/5/2026).

He sets the new description on the clicked picture BEFORE pressing Apply, so the copies still
carry the old description, or none, and the search by description finds nothing. Then the copies
are found by the image (Lp_Same_Pic_Find_By_Image), and the description is copied inside the same
undo record as the resizing.

The rule for WHEN a description is copied is in tests/vba/TestSamePicture.bas. What is held here
is the wiring the VBA runner cannot reach, because it touches pictures in a document.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures


def _procs(repo_root):
    return procedures(module_text(repo_root))


def test_the_image_search_runs_only_when_the_description_finds_nothing(repo_root):
    run = "\n".join(_procs(repo_root)["Lp_Rst_Run"])
    find = run.index("Set matches = Lp_Same_Pic_Find(")
    guard = re.search(r"If byAlt And matches\.count = 0 Then", run)
    assert guard and guard.start() > find
    by_image = run.index("Lp_Same_Pic_Find_By_Image(")
    assert by_image > guard.start()
    # its hidden document is closed before the undo record opens
    assert run.index("reader.Close", by_image) < run.index("StartCustomRecord")


def test_the_description_is_copied_inside_the_one_undo_record(repo_root):
    run = "\n".join(_procs(repo_root)["Lp_Rst_Run"])
    start = run.index("StartCustomRecord")
    copy = run.index("Lp_Same_Pic_Copy_Alt(matches, giveAlt)")
    end = run.index("EndCustomRecord", start)
    assert start < copy < end


def test_copying_goes_through_the_rule(repo_root):
    copy = "\n".join(_procs(repo_root)["Lp_Same_Pic_Copy_Alt"])
    assert "Lp_Same_Pic_Alt_Needs_Copy(alt, refAlt)" in copy
    assert copy.count(".AlternativeText = ") == 1


def test_the_image_search_compares_shape_with_a_copy_not_the_clicked_picture(repo_root):
    """The clicked picture may just have been stretched by hand (issue #49)."""
    body = "\n".join(_procs(repo_root)["Lp_Same_Pic_Find_By_Image"])
    assert "Lp_Same_Pic_Aspect_Close(s, shapeOf(i))" in body
    assert "Lp_Same_Pic_Aspect_Close(s, ref)" not in body
    # no undo record of its own: reading image data inside one costs the single Ctrl+Z
    assert "UndoRecord" not in body


# --- Edit Description (Jerry, 11/5/2026): the description is typed in Picture Description (414)
# before Apply, so the clicked picture keeps its old one and its copies are still found by it.

def test_apply_finds_copies_by_the_pictures_own_description(repo_root):
    run = "\n".join(_procs(repo_root)["Lp_Rst_Run"])
    assert "refAlt = Trim$(ref.AlternativeText)" in run
    assert re.search(r"Lp_Same_Pic_Find\([^)]*refAlt", run)
    # what is given is the typed text, else the picture's own
    assert "giveAlt = Lp_Rst_Take_New_Alt(ref)" in run
    assert "If Not typed Then giveAlt = refAlt" in run
    assert "Lp_Same_Pic_Copy_Alt(matches, giveAlt)" in run


def test_the_clicked_picture_gets_the_typed_text_inside_the_undo_record(repo_root):
    run = "\n".join(_procs(repo_root)["Lp_Rst_Run"])
    start = run.index("StartCustomRecord")
    end = run.index("EndCustomRecord", start)
    give = run.index("ref.AlternativeText = giveAlt")
    assert start < give < end


def test_typed_text_belongs_to_one_picture_and_one_apply(repo_root):
    procs = _procs(repo_root)
    take = "\n".join(procs["Lp_Rst_Take_New_Alt"])
    assert "Lp_Rst_New_Alt_Is_For(ref)" in take
    # used up whichever picture Apply is on
    assert 'Lp_Rst_NewAlt = ""' in take and "Set Lp_Rst_NewAltAt = Nothing" in take
    is_for = "\n".join(procs["Lp_Rst_New_Alt_Is_For"])
    assert "StoryType" in is_for and ".Start = ref.Range.Start" in is_for
    finish = "\n".join(procs["Lp_Rst_Finish"])
    assert 'Lp_Rst_NewAlt = ""' in finish and "Set Lp_Rst_NewAltAt = Nothing" in finish


def test_edit_description_changes_nothing_in_the_book(repo_root):
    body = "\n".join(_procs(repo_root)["Lp_Rst_Edit_Description"])
    assert ".AlternativeText =" not in body
    assert "UndoRecord" not in body
    assert "Lp_Pic_Description_Form.Show" in body
    # the answer is read only after Okay, and an empty one is never kept
    assert "If Lp_Rst_Desc_Ok Then" in body
    assert "Set Lp_Rst_NewAltAt = ref.Range.Duplicate" in body


def test_the_button_and_alt_e_run_it(repo_root):
    from test_toc_box_stays_open import form_code
    code = "\n".join(form_code(repo_root, "Lp_Same_Pic_Range_Form"))
    assert re.search(r"Sub EditAltButton_Click\(\)\s+Lp_Rst_Edit_Description", code)
    alt = "\n".join(_procs(repo_root)["Lp_Rst_AltKey"])
    assert re.search(r'Case "E"\s+Lp_Rst_Edit_Description', alt)


def test_the_description_box_says_okay_and_cancel(repo_root):
    from test_toc_box_stays_open import form_code
    code = "\n".join(form_code(repo_root, "Lp_Pic_Description_Form"))
    assert "Lp_Rst_Desc_Ok = True" in code
    # Cancel and the red X leave the answer as it was - no Okay - and never hold the box open
    q = re.search(r"Sub UserForm_QueryClose\(.*?End Sub", code, re.S)
    assert q and not re.search(r"(?m)^\s*Cancel = True", q.group(0))
    m = re.search(r"Sub CancelButton_Click\(\)(.*?)End Sub", code, re.S)
    assert m and "Lp_Rst_Desc_Ok = True" not in m.group(1)


def test_a_deleted_picture_does_not_hand_its_text_to_the_next(repo_root):
    """From review, 11/5/2026: a deleted picture leaves the held Range collapsed at its start,
    where the next picture slides in. Both ends must match."""
    is_for = "\n".join(_procs(repo_root)["Lp_Rst_New_Alt_Is_For"])
    assert ".End = ref.Range.End" in is_for


def test_375_is_busy_while_414_is_up(repo_root):
    body = "\n".join(_procs(repo_root)["Lp_Rst_Edit_Description"])
    on = body.index("Lp_Rst_Busy = True")
    show = body.index("Lp_Pic_Description_Form.Show")
    off = body.index("Lp_Rst_Busy = False", show)
    assert on < show < off


def test_the_typed_text_is_taken_only_once_nothing_can_refuse(repo_root):
    run = "\n".join(_procs(repo_root)["Lp_Rst_Run"])
    take = run.index("Lp_Rst_Take_New_Alt(ref)")
    assert run.index('"VistaType LP (383)"') < take < run.index("StartCustomRecord")


def test_alt_text_is_copied_when_no_size_changes(repo_root):
    """Jerry, 11/5/2026: "If only the alt-text has been changed (and the size is still the same)
    the replacement should be done." Nothing between the search and the copy may leave Apply
    because no copy needed resizing."""
    run = "\n".join(_procs(repo_root)["Lp_Rst_Run"])
    find = run.index("Set matches = Lp_Same_Pic_Find(")
    copy = run.index("Lp_Same_Pic_Copy_Alt(matches, giveAlt)")
    between = run[find:copy]
    assert "Lp_Same_Pic_Resized" not in between
    assert "Lp_Same_Pic_Already" not in between
    assert "Exit Sub" not in between


def test_the_button_says_alt_text(repo_root):
    from test_toc_box_stays_open import form_code
    code = "\n".join(form_code(repo_root, "Lp_Same_Pic_Range_Form"))
    assert 'EditAltButton.Caption = "Edit picture\'s Alt Text"' in code

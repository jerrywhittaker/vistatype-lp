"""The red X on a form acts like Cancel.

Jerry's rule, 10/11/2026: "All red X's on forms should act like a cancel." A survey of all 48
forms that day found four that did not:

  * 329 Dx_Bullet_Removal_Form, 336 Dx_UEB_EBAE_Fill_In_YN_Form and 362 Sh_Space_Or_Para_Form
    had no UserForm_QueryClose. The X only unloaded the box, skipping what Cancel does (End, or
    clearing the answer), so the macro behind carried on with an answer left from an earlier
    run. On 10/11 each sent the X through its own Cancel handler with Cancel = True - and on
    10/12 Jerry found the X on 329 did nothing: Word ignores an Unload inside QueryClose. The
    X now only clears the answer and lets the box close (the DAISY/NIMAS boxes 325 and 327 had
    the same shape and got the same change).
  * 346 Lp_Columns_Wanted_Form's userform_terminate wrote the column count back on every close
    but Cancel, so the X converted the table as if Okay had been pressed. Only Okay writes it
    now, and Initialize clears it.

None of this can be driven headlessly (a UserForm blocks the runner), so this holds the source.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import procedures

# form -> the answer its red X must clear (None: the form needs no QueryClose at all, because a
# plain close already leaves no answer)
X_CLEARS_THE_ANSWER = {
    "Dx_Bullet_Removal_Form": "Dx_GP_String_2",
    "Sh_Space_Or_Para_Form": "Sh_GP_String_1",
    "DN_XML_Type_Form": "Sh_GP_String_1",
    "DN_Keep_Or_Omit_Images_Form": "Sh_GP_String_2",
    "Dx_UEB_EBAE_Fill_In_YN_Form": None,
}


def form_procedures(repo_root, name):
    return procedures("\n".join(form_code(repo_root, name)))


def test_the_x_clears_the_answer_and_lets_the_box_close(repo_root):
    """Jerry, 10/12/2026: the red X on 329 did nothing. QueryClose set Cancel = True and then ran
    Cancel, whose Unload Me Word ignores inside QueryClose - so the box stayed open. The X must
    clear the answer and leave Cancel alone, so the box closes by itself."""
    for form, answer in X_CLEARS_THE_ANSWER.items():
        procs = form_procedures(repo_root, form)
        if answer is None:
            assert "UserForm_QueryClose" not in procs, (
                f"{form} has a UserForm_QueryClose again - a plain close is already Cancel")
            continue
        assert "UserForm_QueryClose" in procs, (
            f"{form} has no UserForm_QueryClose - its red X no longer clears the answer")
        body = "\n".join(l for l in procs["UserForm_QueryClose"]
                         if not l.strip().startswith("'"))
        assert re.search(rf'CloseMode\s*=\s*vbFormControlMenu\s+Then\s+{answer}\s*=\s*""',
                         body), f"{form}'s red X does not clear {answer}"
        assert not re.search(r"\bCancel\s*=\s*True\b", body), (
            f"{form}'s red X sets Cancel = True - the box would not close")
        assert not re.search(r"\bUnload\b|Cancel_Click", body), (
            f"{form}'s red X unloads from inside QueryClose - Word ignores it")


def test_columns_wanted_x_writes_no_column_count(repo_root):
    procs = form_procedures(repo_root, "Lp_Columns_Wanted_Form")
    for name, body in procs.items():
        if name.lower() not in ("userform_terminate", "userform_queryclose"):
            continue
        for line in body:
            assert not re.search(r"Lp_GP_String_1\s*=(?!\s*\"\")", line), (
                f"Lp_Columns_Wanted_Form's {name} writes Lp_GP_String_1 - the red X would act "
                f"like Okay: {line.strip()}")


def test_columns_wanted_starts_empty_and_only_okay_answers(repo_root):
    procs = form_procedures(repo_root, "Lp_Columns_Wanted_Form")
    init = "\n".join(procs["UserForm_Initialize"])
    assert re.search(r'Lp_GP_String_1\s*=\s*""', init), (
        "Initialize no longer clears the column count - the red X would leave an old one")
    assert re.search(r'Lp_GP_String_3\s*=\s*""', init), (
        "Initialize no longer clears the border settings - an old run's could be used again")
    assert re.search(r"Lp_GP_String_1\s*=\s*TextBox1", "\n".join(procs["OkayButton_Click"]))
    writers = [name for name, body in procs.items() if name != "OkayButton_Click"
               and re.search(r"Lp_GP_String_1\s*=\s*TextBox1", "\n".join(body))]
    assert not writers, f"Only Okay may write the column count; also written in {writers}"

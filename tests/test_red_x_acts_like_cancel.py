"""The red X on a form acts like Cancel.

Jerry's rule, 10/11/2026: "All red X's on forms should act like a cancel." A survey of all 48
forms that day found four that did not:

  * 329 Dx_Bullet_Removal_Form, 336 Dx_UEB_EBAE_Fill_In_YN_Form and 362 Sh_Space_Or_Para_Form
    had no UserForm_QueryClose. The X only unloaded the box, skipping what Cancel does (End, or
    clearing the answer), so the macro behind carried on with an answer left from an earlier
    run. Each now sends the X through its own Cancel handler.
  * 346 Lp_Columns_Wanted_Form's userform_terminate wrote the column count back on every close
    but Cancel, so the X converted the table as if Okay had been pressed. Only Okay writes it
    now, and Initialize clears it.

None of this can be driven headlessly (a UserForm blocks the runner), so this holds the source.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import procedures

# form -> the Cancel handler its red X must go through
X_GOES_THROUGH_CANCEL = {
    "Dx_Bullet_Removal_Form": "CmdCancel_Click",
    "Dx_UEB_EBAE_Fill_In_YN_Form": "CmdCancel_Click",
    "Sh_Space_Or_Para_Form": "Cmd_Cancel_Click",
}


def form_procedures(repo_root, name):
    return procedures("\n".join(form_code(repo_root, name)))


def test_the_x_goes_through_cancel(repo_root):
    for form, cancel in X_GOES_THROUGH_CANCEL.items():
        procs = form_procedures(repo_root, form)
        assert "UserForm_QueryClose" in procs, (
            f"{form} has no UserForm_QueryClose - its red X no longer acts like Cancel")
        body = "\n".join(procs["UserForm_QueryClose"])
        assert re.search(r"CloseMode\s*=\s*vbFormControlMenu", body), form
        assert re.search(r"\bCancel\s*=\s*True\b", body), form
        assert re.search(rf"^\s*(Call\s+)?{cancel}\s*$", body, re.MULTILINE), (
            f"{form}'s red X does not run {cancel}")
        assert cancel in procs, f"{form} has no {cancel} for the red X to run"


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

"""Table and TOC Tools and Selected Cleanup never leave through End (issue #18).

A bare `End` resets the whole VBA project: every UserForm is unloaded, including the boxes that
stay open while the transcriber works - Resize Pictures (375), the TOC box (354), the fill-in
line box (357) and the $pg menus - and nothing is said. Found by reading, 9/23/2026.

Jerry approved, 10/9/2026: Table and TOC Tools (Lp_Table_Tools), its table box (356,
Lp_Table_Tools_Menu_Form) and the convert options box under it, and Selected Cleanup's Cancel,
leave with Exit Sub / Unload instead. The guard subs (Lp_Is_Text_Selected and the like) keep
their End on purpose - it is what stops the macro that called them.

None of this can be driven headlessly (a UserForm blocks the runner), so this holds the source.
"""
import re

from test_toc_box_stays_open import form_code, has_end_statement
from test_toc_never_deletes_a_line import module_text, procedures

EXIT_SUB = re.compile(r"^\s*Exit Sub\s*$", re.IGNORECASE)


def form_procedures(repo_root, name):
    return procedures("\n".join(form_code(repo_root, name)))


def test_table_tools_has_no_end(repo_root):
    tools = procedures(module_text(repo_root))["Lp_Table_Tools"]
    assert not has_end_statement(tools), (
        "Lp_Table_Tools has an End again - it would close Resize Pictures (375) and every other "
        "box that stays open")


def test_table_tools_stops_right_after_the_table_box(repo_root):
    """Once the table box (356) closes, nothing else runs - above all not the TOC check, which a
    whole table selected (more than one paragraph) would otherwise reach."""
    tools = procedures(module_text(repo_root))["Lp_Table_Tools"]
    show = next(n for n, c in enumerate(tools) if re.search(r"Lp_Table_Tools_Menu_Form\.Show\b", c))
    rest = [c for c in tools[show + 1:] if c.strip()]
    # No Unload here: every button and the X have unloaded the box, and naming it would only
    # load it again, running its Initialize.
    assert EXIT_SUB.match(rest[0]), "Lp_Table_Tools must leave with Exit Sub straight after the table box"
    assert not any(re.search(r"Unload Lp_Table_Tools_Menu_Form", c) for c in tools), (
        "Lp_Table_Tools unloads the table box again - it is already unloaded, so this reloads it")


def test_neither_table_form_ends_everything(repo_root):
    for name in ("Lp_Table_Tools_Menu_Form", "Lp_Table_Convert_Options_Form"):
        assert not has_end_statement(form_code(repo_root, name)), (
            f"{name} has an End statement again - it would close every box that stays open")


def test_every_table_box_button_closes_the_box(repo_root):
    """With End gone, a button that does not unload the table box leaves it up. Every button
    that used to End must say Unload, itself or through the convert box it shows."""
    procs = form_procedures(repo_root, "Lp_Table_Tools_Menu_Form")
    for name in ("Cmd_Cancel_Click", "Default_Color_Selected_Click", "PseudoButton_Click",
                 "RealButton_Click", "Default_Color_All_Click", "Yellow_Table_All_Click"):
        assert any(re.search(r"^\s*Unload Me\s*$", c) for c in procs[name]), (
            f"{name} no longer closes the table box")
    # List and Rotate leave the closing to the convert box: it has unloaded this one by the time
    # Show returns, so an Unload Me after it would load the box again only to unload it.
    for name in ("List_Button_Click", "Rotate_Button_Click"):
        body = procs[name]
        show = next(n for n, c in enumerate(body) if "Lp_Table_Convert_Options_Form.Show" in c)
        assert not any(re.search(r"^\s*Unload Me\s*$", c) for c in body[show + 1:]), (
            f"{name} unloads the table box after the convert box, which has already unloaded it")
    convert = form_procedures(repo_root, "Lp_Table_Convert_Options_Form")
    for name in ("CmdCancel_Click", "userform_terminate"):
        assert any(re.search(r"Unload Lp_Table_Tools_Menu_Form", c) for c in convert[name]), (
            f"Lp_Table_Convert_Options_Form.{name} no longer closes the table box")


def test_a_refusal_in_the_convert_box_stops_the_conversion(repo_root):
    """Each refusal in Okay used to End; it must still stop before anything is converted."""
    okay = form_procedures(repo_root, "Lp_Table_Convert_Options_Form")["CmdOkay_Click"]
    found = set()
    for n, c in enumerate(okay):
        m = re.search(r'Sh_Say ".*VistaType LP \((288|319|287|178)\)"', c)
        if m:
            found.add(m.group(1))
            follow = [x for x in okay[n + 1:n + 5] if x.strip()]
            assert EXIT_SUB.match(follow[2]), f"no Exit Sub after the message on line {n} of CmdOkay_Click"
    for number in ("288", "319", "287", "178"):
        assert number in found, f"message {number} is no longer in CmdOkay_Click - this test checks nothing for it"


def test_pseudo_columns_does_not_load_the_convert_box(repo_root):
    """Naming Lp_Table_Convert_Options_Form loads it; with no End to clear it away it stayed
    loaded, and the next List or Rotate showed it without its Initialize."""
    body = procedures(module_text(repo_root))["Lp_Convert_Table_To_Pseudo_Columns"]
    assert not any("Lp_Table_Convert_Options_Form" in c for c in body)
    assert any(re.search(r"Lp_Table_Tools_Menu_Form\.Hide", c) for c in body)


def test_selected_cleanup_cancel_does_not_end_everything(repo_root):
    cancel = form_procedures(repo_root, "Lp_Selected_Cleanup_Form")["CancelButton_Click"]
    assert not has_end_statement(cancel), (
        "Selected Cleanup's Cancel has an End again - it would close every box that stays open")
    # After Cancel, the caller only unloads the form: nothing in it may run a cleanup.
    caller = procedures(module_text(repo_root))["Lp_Selected_File_CleanUp"]
    show = next(n for n, c in enumerate(caller) if "Lp_Selected_Cleanup_Form.Show" in c)
    after = [c.strip() for c in caller[show + 1:] if c.strip()]
    assert after[:2] == ["Unload Lp_Selected_Cleanup_Form", "End Sub"]

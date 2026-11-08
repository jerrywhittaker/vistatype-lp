"""Prodnote Review (425), the large print prodnotes form - Jerry, 11/7/2026.

One button in the VistaType LP File Cleanup group opens it, in the place where Delete Prodnotes
was, with the same icon as the braille Prodnotes and Alt-Text (417) button. Delete Prodnotes
keeps its id on the hidden retired tab, so a toolbar that names it still draws a working button.

Jerry made the form himself in the VBA editor. Its Review/Edit/Delete Prodnotes button opens the
large print Review Prodnotes (426), 11/7/2026 (tests/test_lp_review_prodnotes.py), and its Convert
All Prodnotes to LP Style button asks Box Red Style or Box Blue Style on Convert All Prodnotes (427),
11/8/2026 (tests/test_lp_convert_all_prodnotes.py).
"""
import re

FORM = "Lp_Prodnotes_Menu_Form"
STARTER = "Lp_Prodnotes_Menu_Starter"
NEW_ID = "btn_Lp_Prodnotes_Menu_Starter"
OLD_ID = "btn_Sh_Delete_Prodnote_Paragraphs"


def ribbon(repo_root):
    return (repo_root / "src" / "ribbon" / "customUI14.xml").read_text(encoding="utf-8")


def tab(xml, tab_id):
    return re.search(rf'<tab id="{tab_id}".*?</tab>', xml, re.S).group(0)


def button(xml, bid):
    return re.search(rf'<button id="{bid}"[^>]*/>', xml).group(0)


def test_the_form_is_in_src_with_its_number(repo_root):
    code = (repo_root / "src" / "forms" / f"{FORM}.frm").read_bytes().decode("latin-1")
    assert 'Caption         =   "Prodnote Review (425)"' in code
    assert (repo_root / "src" / "forms" / f"{FORM}.frx").exists()


def test_review_edit_delete_prodnotes_opens_the_large_print_box(repo_root):
    """Jerry, 11/7/2026: CommandButton4, Review/Edit/Delete Prodnotes, opens Review Prodnotes (426)
    through the bridge, and the bridge hides THIS form - it hid 417, copied from that form."""
    code = (repo_root / "src" / "forms" / f"{FORM}.frm").read_bytes().decode("latin-1")
    handler = re.search(r"Private Sub CommandButton4_Click\(\)\r\n(.*?)End Sub", code, re.S)
    assert handler and 'Run_After_Menu "Lp_Review_Prodnotes"' in handler.group(1)
    bridge = re.search(r"Private Sub Run_After_Menu\(ByVal macroName As String\)\r\n(.*?)End Sub",
                       code, re.S).group(1)
    assert "Me.Hide" in bridge and "Dx_Prodnotes_Alt_Text_Menu_Form" not in bridge
    assert bridge.index("Me.Hide") < bridge.index('Application.OnTime Now, "Sh_StartSpinnerBridge"')


def test_the_starter_opens_the_form_modal_after_the_open_book_check(repo_root):
    macros = (repo_root / "src" / "vba" / "LPandBrlMacros.bas").read_bytes().decode("latin-1")
    m = re.search(rf"^Sub {STARTER}\(\)\r\n(.*?)^End Sub", macros, re.S | re.M)
    assert m, f"{STARTER} is missing"
    body = m.group(1)
    assert 'Application.Run MacroName:="Sh_Is_Doc_Open"' in body
    assert f"    {FORM}.Show\r\n" in body
    assert "vbModeless" not in body
    # Large print: no BANA template check.
    assert "Dx_Is_BANA_Template_Attached" not in body


def test_prodnote_review_replaces_delete_prodnotes_on_the_lp_file_cleanup_group(repo_root):
    lp = tab(ribbon(repo_root), "tab_VistaType_LP")
    group = re.search(r'<group id="grp_mso_c1_1C1CE543".*?</group>', lp, re.S).group(0)
    assert re.findall(r'<button id="([^"]+)"', group) == [
        "btn_Lp_File_Fix_Sequence", "btn_Lp_Selected_File_CleanUp", NEW_ID]
    b = button(group, NEW_ID)
    assert 'label="Prodnote Review"' in b
    assert f'tag="{STARTER}"' in b
    assert 'onAction="RibbonAction"' in b


def test_it_has_the_same_icon_as_the_braille_prodnotes_and_alt_text_button(repo_root):
    xml = ribbon(repo_root)
    icon = re.search(r'imageMso="([^"]+)"', button(xml, "btn_Dx_Prodnotes_Alt_Text_Menu_Starter")).group(1)
    assert f'imageMso="{icon}"' in button(xml, NEW_ID)


def test_delete_prodnotes_is_retired_not_removed(repo_root):
    xml = ribbon(repo_root)
    retired = tab(xml, "tab_VT_Retired_Buttons")
    assert f'id="{OLD_ID}"' in retired
    assert 'tag="Sh_Delete_Prodnote_Paragraphs"' in button(retired, OLD_ID)
    assert f'id="{OLD_ID}"' not in tab(xml, "tab_VistaType_LP")
    assert xml.count(f'id="{OLD_ID}"') == 1


def test_the_starter_goes_through_the_dispatch_and_the_installer_tab_file(repo_root):
    dispatch = (repo_root / "src" / "vba" / "RibbonDispatch.bas").read_text(encoding="latin-1")
    assert f'Case "{STARTER}": {STARTER}' in dispatch
    tabs = (repo_root / "installer" / "ribbon-tabs.officeUI").read_text(encoding="utf-8")
    assert f'idQ="x1:{NEW_ID}"' in tabs
    assert f'idQ="x1:{OLD_ID}"' not in tabs
    ids = (repo_root / "installer" / "ribbon-button-ids.txt").read_text(encoding="utf-8").split()
    assert NEW_ID in ids and OLD_ID in ids

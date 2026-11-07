"""Prodnotes and Alt-Text (417), the braille menu form - Jerry, 11/6/2026.

One button in the braille File Cleanup group opens it. Its four buttons run Delete Prodnotes,
Prodnote to TN, and large print's Resize Throughout Pictures in a Selected Range and Locate the
Picture in the Clipboard - all four through the bridge, because the picture tools open boxes
that stay open and cannot paint while this modal menu's handler is running. The two prodnote
buttons that used to be on the ribbon keep their ids on the hidden retired tab, so a toolbar
that names one still draws a working button.
"""
import re

FORM = "Dx_Prodnotes_Alt_Text_Menu_Form"
STARTER = "Dx_Prodnotes_Alt_Text_Menu_Starter"


def form_code(repo_root):
    return (repo_root / "src" / "forms" / f"{FORM}.frm").read_bytes().decode("latin-1")


def ribbon(repo_root):
    return (repo_root / "src" / "ribbon" / "customUI14.xml").read_text(encoding="utf-8")


def tab(xml, tab_id):
    return re.search(rf'<tab id="{tab_id}".*?</tab>', xml, re.S).group(0)


def test_each_button_runs_its_macro_through_the_bridge(repo_root):
    code = form_code(repo_root)
    for button, macro in [("ReviewProdnotes", "Dx_Review_Prodnotes"),
                          ("DeleteProdnotesButton", "Sh_Delete_Prodnote_Paragraphs"),
                          ("ProdnoteToTNButton", "Dx_Change_Prodnotes_To_Transcriber_Notes"),
                          ("SamePictureRangeButton", "Lp_Resize_Pictures_In_Range"),
                          ("LocatePictureButton", "Lp_Locate_Picture")]:
        handler = re.search(rf"Private Sub {button}_Click\(\)\r\n(.*?)End Sub", code, re.S)
        assert handler, f"{button} has no Click handler"
        assert f'Run_After_Menu "{macro}"' in handler.group(1)
    bridge = re.search(r"Private Sub Run_After_Menu\(ByVal macroName As String\)\r\n(.*?)End Sub",
                       code, re.S).group(1)
    assert "Sh_BridgeTargetMacro = macroName" in bridge
    assert 'Application.OnTime Now, "Sh_StartSpinnerBridge"' in bridge
    assert bridge.index(".Hide") < bridge.index("Application.OnTime")


def test_the_form_has_its_number_and_cancel(repo_root):
    code = form_code(repo_root)
    assert 'Caption         =   "Prodnotes and Alt-Text (417)"' in code
    assert "Private Sub Cmd_Cancel_Click()\r\n    Unload Me\r\nEnd Sub" in code


def test_the_starter_is_a_button_on_the_braille_file_cleanup_group(repo_root):
    braille = tab(ribbon(repo_root), "tab_Braille_Macros")
    group = re.search(r'<group id="grp_mso_c1_18AD7D7A".*?</group>', braille, re.S).group(0)
    assert f'tag="{STARTER}"' in group
    assert "<menu" not in braille
    macros = (repo_root / "src" / "vba" / "LPandBrlMacros.bas").read_bytes().decode("latin-1")
    assert re.search(rf"^Sub {STARTER}\(\)\r\n", macros, re.M)


def test_the_old_prodnote_buttons_are_retired_not_removed(repo_root):
    xml = ribbon(repo_root)
    retired = tab(xml, "tab_VT_Retired_Buttons")
    braille = tab(xml, "tab_Braille_Macros")
    for bid in ("btn_Sh_Delete_Prodnote_Paragraphs_x", "btn_Dx_Change_Prodnotes_To_Transcriber_Notes"):
        assert f'id="{bid}"' in retired
        assert f'id="{bid}"' not in braille

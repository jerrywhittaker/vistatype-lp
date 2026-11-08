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


# --- Jerry, 11/8/2026: Delete All Prodnotes and Convert All Prodnotes swap places, Convert All
# --- gets new words, and Delete All deletes after a Yes/No. The layout lives only in the .frx,
# --- so the positions are read from its site records, the same walk as
# --- test_form_handlers_match_controls.py, carried on into each site's position and tab index.

NEW_CAPTION = "Convert All Prodnotes to red or blue box styles"


def _frm(repo_root):
    return (repo_root / "src" / "forms" / f"{FORM}.frm").read_bytes().decode("latin-1")


def _sites(repo_root):
    """{control name: (left, top, tab index)} from the form's `f` stream. Positions are in
    HIMETRIC, as stored; a tab index the designer did not save (an Image) is None."""
    import struct
    import pytest
    olefile = pytest.importorskip("olefile")
    from test_form_handlers_match_controls import OLE_MAGIC, Reader, control_names_in_f_stream

    data = (repo_root / "src" / "forms" / f"{FORM}.frx").read_bytes()
    ole = olefile.OleFileIO(data[data.find(OLE_MAGIC):])
    try:
        f = ole.openstream("f").read()
    finally:
        ole.close()
    names = control_names_in_f_stream(f)            # proves the walk agrees with that test
    out = {}
    # Each site record starts with version 0x0000 then cbSite; find them by name instead of
    # re-walking the header, then read that one site's DataBlock and ExtraDataBlock.
    for name in names:
        tail = name.encode("latin-1")
        hit = None
        for i in range(len(f)):
            if f[i:i + 2] != b"\x00\x00":
                continue
            r = Reader(f[i:])
            try:
                r.take(2)
                cb = r.u16()
                mask = r.u32()
                base = r.pos
                vals = {}
                for bit, size in ((0, 4), (1, 4), (2, 4), (3, 4), (4, 4), (5, 4),
                                  (6, 2), (7, 2), (9, 2), (11, 4), (12, 4), (13, 4), (14, 4)):
                    if mask & (1 << bit):
                        r.align(base, size)
                        vals[bit] = r.u32() if size == 4 else r.u16()
                r.align(base, 4)
                if vals.get(0, 0) & 0x7FFFFFFF != len(tail) or r.take(len(tail)) != tail:
                    continue
                r.align(base, 4)
                if 1 in vals:                       # Tag
                    r.take(vals[1] & 0x7FFFFFFF)
                    r.align(base, 4)
                if not mask & (1 << 8):
                    continue
                left, top = struct.unpack("<ii", r.take(8))
                if r.pos > 4 + cb:
                    continue
                hit = (left, top, vals.get(6))
                break
            except ValueError:
                continue
        assert hit, f"no site record found for {name}"
        out[name] = hit
    return out


def test_convert_all_is_above_delete_all_with_their_icons(repo_root):
    """Jerry, 11/8/2026: switch the positions of Delete All Prodnotes and Convert All Prodnotes.
    Convert All (CommandButton2) is now second, Delete All (CommandButton1) third, and each
    icon moved with its button: Image12 is Convert All's picture, Image13 Delete All's."""
    s = _sites(repo_root)
    review, convert, delete = s["CommandButton4"], s["CommandButton2"], s["CommandButton1"]
    assert review[1] < convert[1] < delete[1]
    assert review[0] == convert[0] == delete[0]
    assert s["Image12"][:2] == convert[:2]
    assert s["Image13"][:2] == delete[:2]
    # Tab order follows the screen between the two.
    assert convert[2] < delete[2]


def test_the_old_buttons_from_417_are_still_covered(repo_root):
    """The buttons copied from 417 sit exactly under the new ones. The swap trades the new
    buttons between two spots that are both covered, so neither old one shows."""
    s = _sites(repo_root)
    new_spots = {s[n][:2] for n in ("CommandButton4", "CommandButton1", "CommandButton2")}
    for old in ("ReviewProdnotes", "DeleteProdnotesButton", "ProdnoteToTNButton"):
        assert s[old][:2] in new_spots, old


def test_convert_all_has_jerrys_new_words(repo_root):
    """On two lines, broken after "to", so the words clear the icon over the button's left end:
    on one line, centered in Tahoma 11.25, they start under it (measured 233 points of 270)."""
    frx = (repo_root / "src" / "forms" / f"{FORM}.frx").read_bytes()
    assert b"Convert All Prodnotes to\r\nred or blue box styles" in frx
    assert (NEW_CAPTION + " Button").encode() in frx               # the hover text
    assert b"Convert All Prodnotes to LP Style" not in frx


def test_delete_all_prodnotes_deletes_after_a_yes_no(repo_root):
    """Jerry, 11/8/2026: Delete All Prodnotes (CommandButton1) deletes every prodnote after a
    confirmation. Sh_Delete_Prodnote_Paragraphs asks Yes/No (229) before it deletes anything."""
    code = _frm(repo_root)
    handler = re.search(r"Private Sub CommandButton1_Click\(\)\r\n(.*?)End Sub", code, re.S)
    assert handler and handler.group(1).strip() == 'Run_After_Menu "Sh_Delete_Prodnote_Paragraphs"'
    macros = (repo_root / "src" / "vba" / "LPandBrlMacros.bas").read_bytes().decode("latin-1")
    body = re.search(r"^Sub Sh_Delete_Prodnote_Paragraphs\(\)\r\n(.*?)^End Sub", macros,
                     re.S | re.M).group(1)
    ask = body.index('vbYesNo + vbQuestion, IIf(isLP, "VistaType LP (229)"')
    assert ask < body.index("marked(i).Delete")
    assert "<> vbYes Then" in body[ask:ask + 120]

"""The Locate button's tip says "in the Clipboard", like its caption.

Jerry, 11/2/2026: on LP user form Background Picture Menu (344), Lp_Bakgrnd_Picture_Menu_Form,
the button's caption read "Locate the Picture in the Clipboard" while its ControlTipText read
"Locate the Picture on the Clipboard Button". Both live in the .frx as plain one-byte text.
"""

FRX = "src/forms/Lp_Bakgrnd_Picture_Menu_Form.frx"


def test_locate_tip_says_in_the_clipboard(repo_root):
    text = (repo_root / FRX).read_bytes().decode("latin-1")
    assert "Locate the Picture in the Clipboard Button" in text
    assert "on the Clipboard" not in text

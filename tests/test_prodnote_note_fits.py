"""Prodnotes in This Document (361): the whole note shows without scrolling.

Jerry, 10/28/2026, on 3.0.619: "on message 361 the text is being cut off at the bottom by the
button". Measured on the build box: the text box (Info_Text) is 254.4 points high in the layout
file, about 20 lines of Tahoma 10.2, and the note needs about 21 to 22. UserForm_Initialize makes
the box 42 points taller and moves the Close button and the form's bottom down by the same, so
the button stays below the box and inside the form.
"""
from pathlib import Path


def test_the_box_grows_and_the_button_and_form_move_with_it(repo_root):
    src = (Path(repo_root) / "src" / "forms" / "Sh_Prodnote_Info_Form.frm").read_bytes().decode("latin-1")
    init = src[src.index("Private Sub UserForm_Initialize()"):]
    init = init[:init.index("End Sub")]
    grow = "Info_Text.Height = Info_Text.Height + 42"
    move = "Close_Me.Top = Close_Me.Top + 42"
    form = "Me.Height = Me.Height + 42"
    for line in (grow, move, form):
        assert line in init, line
    # all before the form is centered, which uses its height
    assert init.index(form) < init.index("Me.Top = Application.Top")


def test_the_button_is_close_with_alt_c(repo_root):
    """Jerry, 10/28/2026: "make it Close with Alt+C"."""
    src = (Path(repo_root) / "src" / "forms" / "Sh_Prodnote_Info_Form.frm").read_bytes().decode("latin-1")
    assert 'Close_Me.Caption = "Close"' in src
    assert 'Close_Me.Accelerator = "C"' in src

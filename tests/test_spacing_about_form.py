"""About Line and Letter Spacing (429): Jerry's explanation behind 424's About button.

Jerry, 11/26/2026: the About line/letter spacing button (Alt+A) on Spacing and Italics Settings
(424) showed a message saying the explanation was still to come. His text runs to about 3,200
characters, past the roughly 1,024 a MsgBox shows before VBA cuts it off without a word, so it
is a read-only form, Lp_Spacing_About_Form, built by New-UserForm.ps1 -InfoDialog in the same
shape as Prodnotes in This Document (361): a locked text box that scrolls, and an Okay button.

The text is built in the form's code, not stored in the .frx, so it can be read and checked here.
A UserForm cannot be shown headlessly, so this holds the source.
"""
import re
from pathlib import Path

FORM = "Lp_Spacing_About_Form"


def form_text(repo_root):
    return (Path(repo_root) / "src" / "forms" / f"{FORM}.frm").read_bytes().decode("latin-1")


def form_code(repo_root):
    """The form's code with the comment lines left out."""
    return "\r\n".join(l for l in form_text(repo_root).split("\r\n")
                        if not l.strip().startswith("'"))


def about_text(repo_root):
    """The note as the form builds it: the quoted pieces of AboutText, joined."""
    frm = form_text(repo_root)
    body = frm[frm.index("Private Function AboutText() As String"):]
    body = body[:body.index("End Function")]
    out = ""
    for line in body.split("\r\n"):
        line = line.strip()
        if line.startswith("m = m & vbCrLf & vbCrLf"):
            out += "\n\n"
        for piece in re.findall(r'"((?:[^"]|"")*)"', line):
            out += piece.replace('""', '"')
    return out


def test_the_form_exists_with_its_title_and_number(repo_root):
    frm = form_text(repo_root)
    assert (Path(repo_root) / "src" / "forms" / f"{FORM}.frx").exists()
    assert re.search(r'^\s+Caption\s+=\s+"About Line and Letter Spacing \(429\)"\r$', frm, re.M)


def test_424s_about_button_opens_it(repo_root):
    frm = (Path(repo_root) / "src/forms/Lp_Spacing_Italics_Form.frm").read_bytes().decode("latin-1")
    body = re.search(r"Private Sub AboutSpacing_Click\(\)\r\n(.*?)End Sub", frm, re.S).group(1)
    assert re.fullmatch(rf"\s*{FORM}\.Show\s*", body)


def test_okay_only_closes_it(repo_root):
    """Okay (Close_Me, captioned Okay with Alt+O by New-UserForm.ps1, and left so) unloads the
    form and does nothing else; nothing renames it to Close as 361 does."""
    frm = form_code(repo_root)
    body = re.search(r"Private Sub Close_Me_Click\(\)\r\n(.*?)End Sub", frm, re.S).group(1)
    assert re.fullmatch(r"\s*Unload Me\s*", body)
    assert "Close_Me.Caption" not in frm and "Close_Me.Accelerator" not in frm
    assert "MsgBox" not in frm


def test_the_red_x_is_a_plain_close(repo_root):
    """Nothing to keep, so the X closing the form is already Cancel (Jerry, 10/11/2026)."""
    frm = form_code(repo_root)
    assert "UserForm_QueryClose" not in frm
    assert "terminate" not in frm.lower()
    assert not re.search(r"\bLp_Pref_Chosen", frm)


def test_the_note_is_jerrys_text(repo_root):
    text = about_text(repo_root)
    assert text.startswith('Using the "Spacing and Italics Settings"\n\n')
    assert len(text) > 1024          # why it is a form and not a message
    for phrase in (
        "The two most important readability factors for large print are font size and white space.",
        '"VTLP V2 (APH)" - the starting settings.',
        "World Wide Web Consortium's Web Content Accessibility Guidelines",
        "double-A (AA) level",
        "If you have never saved a default, it gets the VTLP V2 (APH) settings.",
        "Nothing is saved until you press Okay on Attach Template Form (338).",
        "blue, green, orange, red or violet",
        "the reader could not tell them apart.",
    ):
        assert phrase in text, phrase
    # the headings stand on their own lines, a blank line before and after
    for heading in ("Choice Buttons Controlling White Space Size", "Save as My Default",
                    "This Document Only", "Italics"):
        assert f"\n\n{heading}\n\n" in text, heading


def test_the_note_is_plain_ascii(repo_root):
    """A curly apostrophe or quotation mark would need ChrW in VBA; the note uses straight ones."""
    assert all(ord(c) < 128 for c in about_text(repo_root))


def test_the_note_opens_at_its_first_line(repo_root):
    frm = form_text(repo_root)
    init = frm[frm.index("Private Sub UserForm_Initialize()"):]
    init = init[:init.index("End Sub")]
    assert "Info_Text.Text = AboutText()" in init
    assert init.index("Info_Text.Text = AboutText()") < init.index("Info_Text.selStart = 0")

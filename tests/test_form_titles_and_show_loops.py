"""Two faults Jerry reported on 3.0.519 in the braille translation box (331), and the checks
that keep either from coming back on any form.

1. THE TITLE WAS CUT OFF. "What braille translation will be used for this document? (331)" is
   wider than the title bar the form gave it, so Windows ended it in "..." and the dialog's
   number -- the one thing a transcriber quotes when reporting what they saw -- was the part
   that went. The form is now 279 points wide inside, not 267.75, and Jerry has seen the whole
   title on screen at that width.

2. CANCEL LOOPED FOR EVER. Dx_Attach_BANA_Template_Run showed the form inside
   `Do While Dx_UEB_EBAE_String = ""`, which only the four table buttons could end. Cancel
   (added in 3.0.472) and the title bar's X never set it, so the box came straight back.

3. CANCEL NOW STOPS THE WHOLE ATTACH (Jerry's decision, 9/29/2026). The attach asks the
   translation before anything changes the book, leaves at once on Cancel, and
   Dx_Is_BANA_Template_Attached returns False so every braille macro that calls it stops too.
   Jerry, later the same day: "there should be a message after the cancel". The attach says
   "The BANA braille template has not been attached." (Braille Macros (393)) as it leaves --
   once, whichever button started it; the braille macro that started it adds nothing.

See docs/Reported-Errors.md, 9/29/2026.

HOW THE TITLE IS MEASURED. Windows 10 and 11 draw a title bar in Segoe UI 9 point. The glyph
widths below were measured on the build box on 9/29/2026 with .NET's
Graphics.MeasureString (GenericTypographic, 96 DPI), one character at a time; their sum for the
331 title is 321.795 px, the same as measuring the whole string at once, so adding them up is
exact enough here. How much of the bar the close button and the padding either side take is
set from what Jerry saw: see TITLE_BAR_ALLOWANCE_PX. (It was first ESTIMATED at 62 px, which
wrongly called his 279 pt layout of 331 too narrow.)

The width comes from the .frm's ClientWidth. That is only safe because the .frx agrees with it
(the form's size lives in BOTH, and on import Word draws the binary's -- the 9/22/2026 trap in
the creating-a-UserForm notes), so the first test checks the two agree, on every form.
"""
import pathlib
import re
import struct

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[1]
FORMS = ROOT / "src" / "forms"
FORM_FILES = sorted(FORMS.glob("*.frm"))

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# Segoe UI 9 point at 96 DPI, px, for character codes 32 to 126. Measured on vistabuild 9/29/2026.
SEGOE_UI_9 = dict(zip(range(32, 127), (
    3.29, 3.41, 4.71, 7.09, 6.47, 9.82, 9.60, 2.76, 3.62, 3.62, 5.00, 8.21, 2.60, 4.80, 2.60, 4.68,
    6.47, 6.47, 6.47, 6.47, 6.47, 6.47, 6.47, 6.47, 6.47, 6.47, 2.60, 2.60, 8.21, 8.21, 8.21, 5.38,
    11.46, 7.74, 6.88, 7.43, 8.41, 6.07, 5.86, 8.23, 8.52, 3.19, 4.28, 6.96, 5.65, 10.78, 8.98, 9.05,
    6.72, 9.05, 7.18, 6.38, 6.29, 8.24, 7.45, 11.21, 7.08, 6.63, 6.84, 3.62, 4.55, 3.62, 8.21, 4.98,
    3.22, 6.11, 7.05, 5.54, 7.07, 6.28, 3.76, 7.07, 6.79, 2.91, 2.91, 5.96, 2.91, 10.34, 6.79, 7.03,
    7.05, 7.07, 4.17, 5.09, 4.07, 6.79, 5.75, 8.67, 5.51, 5.81, 5.43, 3.62, 2.87, 3.62, 8.21)))

# Close button plus padding, px. Set from Jerry's on-screen check of form 331 on 9/29/2026: its
# 322 px title (321.8) showed in full, "(331)" included, in a 372 px wide form (ClientWidth 5580,
# 279 pt). So the real allowance is at most 50.2 px, and 50 is the largest whole figure that
# evidence allows. The LOWER bound is unknown -- nothing yet has been seen cut off at a known
# small margin -- except that the old 357 px (267.75 pt) 331 WAS cut off on screen, so it is
# above 35.2 px; at 50 that old width fails by 15 px, as it must.
TITLE_BAR_ALLOWANCE_PX = 50

# Forms whose title does not fit by the rule above and are not yet widened. Take a form OFF this
# list when it is widened -- the last test fails until it is, so the list cannot go stale.
# Empty since 9/29/2026: the three this test first found (issue #22) -- 334
# Dx_Spelling_List_Options_Form, 354 Lp_TOC_Format_And_Color_Form, 327 DN_XML_Type_Form, short
# by 10.4, 8.8 and 6.8 px at 50 px -- were widened to 288, 228 and 228 pt inside, about 20 px
# more than the rule asks. Set by measurement; none of the three has been seen on screen yet.
KNOWN_TOO_NARROW = set()


def _header(frm):
    text = frm.read_bytes().decode("latin-1")
    head = text.split("\nEnd", 1)[0]
    caption = re.search(r'^\s*Caption\s*=\s*"(.*)"\s*$', head, re.M)
    width = re.search(r"ClientWidth\s*=\s*(\d+)", head)
    return (caption.group(1) if caption else None), int(width.group(1))


def _frx_inside_width_himetric(frx):
    """The form's DisplayedSize width from its `f` stream ([MS-OFORMS] 2.2.10.1), HIMETRIC.

    Every form here stores DisplayedSize then LogicalSize and nothing after them, so the width
    is the u32 sixteen bytes before the end of the DataBlock and ExtraDataBlock."""
    olefile = pytest.importorskip(
        "olefile", reason="needs olefile to read the .frx files: sudo apt install python3-olefile")
    data = frx.read_bytes()
    ole = olefile.OleFileIO(data[data.find(OLE_MAGIC):])
    try:
        f = ole.openstream("f").read()
    finally:
        ole.close()
    _, cb_form, mask = struct.unpack_from("<HHI", f, 0)
    assert mask & (3 << 26) == (3 << 26) and mask >> 28 == 0, (
        f"{frx.name}: form stream mask {mask:#x} is not the shape this test reads")
    return struct.unpack_from("<I", f, 4 + cb_form - 16)[0]


def _title_px(caption):
    return sum(SEGOE_UI_9.get(ord(c), 11.46) for c in caption)   # unknown: widest, '@'


def _client_px(twips):
    return twips / 20 * 96 / 72


def _too_narrow(frm):
    caption, twips = _header(frm)
    return bool(caption) and _title_px(caption) > _client_px(twips) - TITLE_BAR_ALLOWANCE_PX


@pytest.mark.parametrize("frm", FORM_FILES, ids=lambda p: p.stem)
def test_frm_width_agrees_with_frx(frm):
    _, twips = _header(frm)
    himetric = _frx_inside_width_himetric(frm.with_suffix(".frx"))
    # Twips are rounded to whole pixels; 30 HIMETRIC is under a point.
    assert abs(himetric - twips / 20 / 72 * 2540) <= 30, (
        f"{frm.name}: ClientWidth {twips} twips but the .frx is {himetric} HIMETRIC "
        f"({himetric / 2540 * 72:.2f} pt). Word draws the .frx's -- resize it in Word and export.")


def test_the_331_title_now_fits():
    frm = FORMS / "Dx_Choose_Translation_Form.frm"
    caption, twips = _header(frm)
    assert caption.endswith("(331)")
    assert abs(_title_px(caption) - 321.795) < 0.5          # the table reproduces the measurement
    # The width Jerry saw the whole title at, or wider.
    assert twips >= 5580, f"331 is {twips} twips wide; Jerry saw the title whole at 5580 (279 pt)"
    assert not _too_narrow(frm), (
        f"331's title is {_title_px(caption):.0f} px in a {_client_px(twips):.0f} px form")


def test_the_old_331_width_still_fails():
    """267.75 pt was cut off on screen in 3.0.519. The rule must still say so, or it is too lax."""
    caption, _ = _header(FORMS / "Dx_Choose_Translation_Form.frm")
    assert _title_px(caption) > _client_px(5355) - TITLE_BAR_ALLOWANCE_PX


@pytest.mark.parametrize("frm", [f for f in FORM_FILES if f.stem not in KNOWN_TOO_NARROW],
                         ids=lambda p: p.stem)
def test_title_fits_the_title_bar(frm):
    caption, twips = _header(frm)
    assert not _too_narrow(frm), (
        f"{frm.name}: the title {caption!r} is {_title_px(caption):.0f} px, and a "
        f"{_client_px(twips):.0f} px form leaves about "
        f"{_client_px(twips) - TITLE_BAR_ALLOWANCE_PX:.0f} px for it. Windows cuts the end off "
        f"with '...', and the end is the dialog's number. Widen the form in Word and export it.")


def test_known_too_narrow_list_is_current():
    fixed = sorted(n for n in KNOWN_TOO_NARROW if not _too_narrow(FORMS / f"{n}.frm"))
    assert not fixed, f"these now fit -- take them off KNOWN_TOO_NARROW: {fixed}"


# ---- 2. no UserForm is shown inside a Do loop ------------------------------------------------

SOURCES = sorted(list((ROOT / "src" / "vba").glob("*.bas")) + list((ROOT / "src" / "vba").glob("*.cls"))
                 + FORM_FILES)
FORM_NAMES = {f.stem.lower() for f in FORM_FILES}


def _forms_shown_in_do_loops(text):
    """(line, form) for each `<UserForm>.Show` that sits inside a Do ... Loop.

    A loop round a modal form can only end on something the form sets, and Cancel and the X
    set nothing -- which is how 331 came back for ever. The one Do-wrapped .Show in the source
    on 9/29/2026 is a FileDialog's, which is not a UserForm and is not matched."""
    found, depth = [], 0
    for n, line in enumerate(text.split("\n"), 1):
        code = line.strip()
        if code.startswith("'"):
            continue
        code = code.split(" '")[0].strip()
        if re.match(r"(?i)^(public\s+|private\s+|friend\s+)?(static\s+)?(sub|function)\s", code):
            depth = 0
        if re.match(r"(?i)^do\b", code):
            depth += 1
        elif re.match(r"(?i)^loop\b", code):
            depth = max(0, depth - 1)
        if depth:
            for m in re.finditer(r"(?i)\b([A-Za-z]\w*)\.Show\b", code):
                if m.group(1).lower() in FORM_NAMES:
                    found.append((n, m.group(1)))
    return found


def test_the_scanner_finds_the_old_331_loop():
    old = ('Sub X()\n    Dx_UEB_EBAE_String = ""\n    Do While Dx_UEB_EBAE_String = ""\n'
           '        Dx_Choose_Translation_Form.Show\n    Loop\nEnd Sub\n')
    assert _forms_shown_in_do_loops(old) == [(4, "Dx_Choose_Translation_Form")]


@pytest.mark.parametrize("src", SOURCES, ids=lambda p: p.name)
def test_no_userform_is_shown_inside_a_do_loop(src):
    hits = _forms_shown_in_do_loops(src.read_bytes().decode("latin-1"))
    assert not hits, (
        f"{src.name}: a UserForm is shown inside a Do loop -- Cancel and the X set nothing, so "
        f"the loop cannot end on them:\n" + "\n".join(f"  line {n}: {f}.Show" for n, f in hits))


def test_the_attach_asks_the_translation_once():
    text = (ROOT / "src" / "vba" / "LPandBrlMacros.bas").read_bytes().decode("latin-1")
    body = text.split("Public Sub Dx_Attach_BANA_Template_Run(", 1)[1].split(
        "end of Dx_Attach_BANA_Template_Run macro", 1)[0]
    shows = [l for l in body.split("\n")
             if "Dx_Choose_Translation_Form.Show" in l and not l.strip().startswith("'")]
    assert len(shows) == 1


# ---- 3. Cancel on 331 stops the whole attach (Jerry, 9/29/2026) -------------------------------
#
# His rule: Cancel "should stop the entire attachment process" -- the book left as it was before
# Attach was pressed, and a braille macro that started the attach stopped as well. The attach
# says so, once, with message 393; the braille macro adds no second message. The
# box cannot be shown headlessly, so these check the SHAPE that makes it true: the question comes
# before anything touches the book, Cancel leaves straight away, and the check every braille
# macro runs reports "not attached" so each of them stops.

MACROS = ROOT / "src" / "vba" / "LPandBrlMacros.bas"


def _code_lines(body):
    """The body's lines with comments and blank lines dropped (a trailing ' comment is cut)."""
    out = []
    for line in body.replace("\r", "").split("\n"):
        code = line.strip()
        if not code or code.startswith("'"):
            continue
        out.append(code.split(" '")[0].strip())
    return out


def _body(start, end):
    text = MACROS.read_bytes().decode("latin-1")
    return text.split(start, 1)[1].split(end, 1)[0]


def _attach():
    return _code_lines(_body("Public Sub Dx_Attach_BANA_Template_Run(",
                             "end of Dx_Attach_BANA_Template_Run macro"))


def _first(lines, needle):
    hits = [n for n, l in enumerate(lines) if needle in l]
    assert hits, f"{needle!r} is not in the attach any more"
    return hits[0]


# Each of these changes the book (or Word's configuration for it). All must come after the
# translation question, or a Cancel there would leave something half done.
CHANGES_THE_BOOK = (
    'Application.Run MacroName:="Dx_Fix_Foreign_Languages"',
    ".AttachedTemplate = TemplatePathandName",
    'Application.Run MacroName:="MS_Set_Word_Config_For_Braille"',
    "ActiveDocument.UndoClear",
    "Sh_Config_Skip_Display = True",
    'Application.Run MacroName:="Dx_Add_Color_To_Foreign_Language_Words"',
    "Dx_Fix_Common_File_Errors_Run True",
)


def test_the_attach_asks_the_translation_before_it_changes_the_book():
    lines = _attach()
    ask = _first(lines, "Dx_Choose_Translation_Form.Show")
    assert _first(lines, "Dx_Choose_BANA_Template_Form.Show") < ask
    late = [c for c in CHANGES_THE_BOOK if _first(lines, c) < ask]
    assert not late, f"these run BEFORE the translation question, so Cancel would leave them done: {late}"


CANCEL_MESSAGE = 'Sh_Say "The BANA braille template has not been attached.", "Braille Macros (393)"'


def test_cancel_in_the_attach_says_so_then_leaves():
    """The flag is lowered just before the box, and the very next thing is the test of it. On
    Cancel the message is shown through Sh_Say (never a MsgBox), and then the sub leaves -- all of
    it before anything in CHANGES_THE_BOOK, which the test above checks comes after the box."""
    lines = _attach()
    ask = _first(lines, "Dx_Choose_Translation_Form.Show")
    assert lines[ask - 1] == "Dx_Translation_Answered = False"
    assert lines[ask + 1:ask + 6] == ["If Not Dx_Translation_Answered Then",
                                      CANCEL_MESSAGE,
                                      "Application.ScreenUpdating = su_Prev",
                                      "Exit Sub", "End If"]
    first_change = min(_first(lines, c) for c in CHANGES_THE_BOOK)
    assert ask + 5 < first_change


def test_the_cancel_message_is_worded_and_numbered_once():
    """Jerry's wording, with his typo put right ("has not be" -> "has not been"), and 393 used by
    this one message only, across every module and form."""
    assert "The BANA braille template has not been attached." in CANCEL_MESSAGE
    hits = []
    for src in SOURCES + FORM_FILES:   # the code and the form captions
        for line in src.read_bytes().decode("latin-1").replace("\r", "").split("\n"):
            code = line.strip()
            if not code.startswith("'") and "(393)" in code:
                hits.append((src.name, code))
    assert hits == [("LPandBrlMacros.bas", CANCEL_MESSAGE)], hits


def test_only_a_table_button_raises_the_answered_flag():
    """Dx_Set_BrailleType raises it; every table button calls that; Cancel calls nothing."""
    setter = _code_lines(_body("Public Sub Dx_Set_BrailleType(", "end of Dx_Set_BrailleType"))
    guard = setter.index('If brlType = "" Then Exit Sub')
    assert setter[guard + 1] == "Dx_Translation_Answered = True"   # after the guard, before any write
    frm = (FORMS / "Dx_Choose_Translation_Form.frm").read_bytes().decode("latin-1")
    subs = dict(re.findall(r"(?s)Private Sub (\w+)_Click\(\)(.*?)End Sub", frm))
    assert set(subs) == {"Cmd_Cancel", "Cmd_BANA_EBAE_Button", "Cmd_BANA_EBAE_Nemeth_Button",
                         "Cmd_BANA_UEB_Button", "Cmd_BANA_UEB_Nemeth_Button"}
    for name, body in subs.items():
        sets = "Dx_Set_BrailleType" in "\n".join(_code_lines(body))
        assert sets == (name != "Cmd_Cancel"), name


def test_the_box_does_not_configure_word_itself():
    """Shown before the attach, a braille configuration here would outlive a Cancel."""
    frm = (FORMS / "Dx_Choose_Translation_Form.frm").read_bytes().decode("latin-1")
    init = frm.split("Private Sub UserForm_Initialize()", 1)[1].split("End Sub", 1)[0]
    assert not any("MS_Set_Word_Config" in l for l in _code_lines(init))


def test_the_check_reports_not_attached_on_cancel():
    body = _body("\nPublic Function Dx_Is_BANA_Template_Attached() As Boolean",
                 "end of Dx_Is_BANA_Template_Attached macro")
    lines = _code_lines(body)
    assert not any("justAttached" in l for l in lines)
    attach = _first(lines, "Dx_Attach_BANA_Template_Run False")
    # Straight after the attach: still no template means Cancel, and False goes back -- with no
    # message of its own in between, because the attach has just shown 393. A second one here
    # would put the same news on screen twice.
    assert lines[attach + 1] == ('If InStr(ActiveDocument.AttachedTemplate, "BANA Braille") = 0 '
                                 'Then Exit Function')
    assert not any("(393)" in l for l in lines)
    # Its own question, when the book has a template but no translation: Cancel is False too.
    ask = _first(lines, "Dx_Choose_Translation_Form.Show")
    assert lines[ask - 1] == 'If Dx_Ensure_BrailleType() = "" Then'
    assert lines[ask + 1] == 'If Dx_Ensure_BrailleType() = "" Then Exit Function'
    # ...and only after an answer is Word set up for braille (the box itself no longer does it).
    assert lines[ask + 2] == 'Application.Run MacroName:="MS_Set_Word_Config_For_Braille"'
    # True is set once, as the last statement: every early way out is False.
    assert lines[-1] == "End Function"
    trues = [n for n, l in enumerate(lines) if l == "Dx_Is_BANA_Template_Attached = True"]
    assert trues == [len(lines) - 2]


def test_every_braille_macro_stops_when_the_check_says_no():
    calls = []
    for src in SOURCES:
        text = src.read_bytes().decode("latin-1")
        for n, line in enumerate(text.split("\n"), 1):
            code = line.strip()
            if code.startswith("'"):
                continue
            code = code.split(" '")[0].strip()
            if "Dx_Is_BANA_Template_Attached" not in code:
                continue
            if re.match(r"(?i)^(public\s+)?function\s+Dx_Is_BANA_Template_Attached\b", code) or \
                    code.startswith("Dx_Is_BANA_Template_Attached = "):
                continue
            calls.append((src.name, n, code))
    wrong = [c for c in calls if c[2] != "If Not Dx_Is_BANA_Template_Attached() Then Exit Sub"]
    assert not wrong, ("a call that does not stop when the user cancels -- Application.Run "
                       "cannot hear the answer:\n" + "\n".join(map(str, wrong)))
    assert len(calls) == 13, f"expected the thirteen braille macros, found {len(calls)}: {calls}"

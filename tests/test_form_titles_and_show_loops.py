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

# Forms whose title does not fit by the rule above, found by this test on 9/29/2026 and not yet
# widened: reported to Jerry (issue #22), not seen on screen. Short by, at 50 px: 334 10.4 px,
# 354 8.8 px, 327 6.8 px. Take a form OFF this list when it is widened --
# the last test fails until it is, so the list cannot go stale.
KNOWN_TOO_NARROW = {
    "Dx_Spelling_List_Options_Form",   # Format Spelling Word List (Contracted/Uncontracted) (334)
    "Lp_TOC_Format_And_Color_Form",    # Format TOC - Add/Remove Color Bars (354)
    "DN_XML_Type_Form",                # Convert DAISY or NIMAS .xml to Word (327)
}


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


def test_no_second_ask_straight_after_the_attach():
    """Dx_Is_BANA_Template_Attached runs the attach, which asks; after a Cancel there it must
    not put the same box up again at once -- that reads as the loop all over again."""
    text = (ROOT / "src" / "vba" / "LPandBrlMacros.bas").read_bytes().decode("latin-1")
    body = text.split("\nSub Dx_Is_BANA_Template_Attached()", 1)[1].split(
        "end of Dx_Is_BANA_Template_Attached macro", 1)[0]
    assert re.search(r"Dx_Attach_BANA_Template_Run False[^\n]*\r?\n\s*justAttached = True", body)
    assert re.search(r"If Not justAttached Then\s*\r?\n\s*If Dx_Ensure_BrailleType\(\) = \"\" "
                     r"Then Dx_Choose_Translation_Form\.Show", body)

"""About Line and Letter Spacing (429): the whole note can be reached, with a scroll bar.

Jerry, 11/26/2026, on 3.0.720: "The text has been truncated and needs a scrollbar." Read from the
.frx (this file does it the same way): the layout was right - Info_Text is a TextBox, MultiLine,
WordWrap, locked, enabled, not AutoSize, with a vertical scroll bar, 564 x 390 points at 12, 12;
Okay sits below it; both are inside the form's 588 x 450.75 point client area, and the .frx's size
matches the .frm header. What was missing is that the box never has the focus (TabStop off), and
MSForms does not draw a text box's scroll bar for text put in before the form is shown until the
box has had the focus. UserForm_Activate gives it the focus once, puts the caret on the first line,
and hands the focus to Okay.

The .frx is binary: a 24-byte header, then an OLE compound file whose `f` stream lists the sites
(name, type, position) and whose `o` stream holds each control's own properties, laid out as in
[MS-OFORMS]. Sizes are stored in HIMETRIC (2540 to the inch).
"""
import re
import struct
from pathlib import Path

import pytest

olefile = pytest.importorskip(
    "olefile", reason="needs olefile to read the .frx files: sudo apt install python3-olefile")

FORM = "Lp_Spacing_About_Form"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
TEXTBOX, COMMANDBUTTON = 23, 17         # ClsidCacheIndex, [MS-OFORMS] 2.4.1


def pt(himetric):
    return himetric * 72 / 2540


class Block:
    """Reads a DataBlock by its property mask: (bit, size) in stored order, where size is
    1, 2 or 4 bytes, or "s" for a string count (4 bytes; the text is in the ExtraDataBlock)."""

    def __init__(self, data, pos):
        self.data, self.pos, self.base = data, pos, pos

    def read(self, mask, layout):
        out = {}
        for bit, size in layout:
            if not mask & (1 << bit):
                continue
            n = 4 if size == "s" else size
            self.pos += (-(self.pos - self.base)) % n
            out[bit] = int.from_bytes(self.data[self.pos:self.pos + n], "little")
            self.pos += n
        self.pos += (-(self.pos - self.base)) % 4
        return out

    def string(self, count):
        n, compressed = count & 0x7FFFFFFF, count & 0x80000000
        raw = self.data[self.pos:self.pos + n]
        self.pos += n + (-n) % 4
        return raw.decode("latin-1") if compressed else raw.decode("utf-16-le")

    def pair(self):
        a, b = struct.unpack_from("<II", self.data, self.pos)
        self.pos += 8
        return a, b


FORM_LAYOUT = [(1, 4), (2, 4), (3, 4), (6, 4), (7, 1), (8, 1), (9, 1), (13, 4), (15, 2),
               (16, 1), (17, 1), (18, 4), (19, "s"), (20, 2), (21, 2), (22, 4), (23, 1),
               (24, 1), (25, 1), (26, 4), (27, 4)]
SITE_LAYOUT = [(0, "s"), (1, "s"), (2, 4), (3, 4), (4, 4), (5, 4), (6, 2), (7, 2), (9, 2),
               (11, "s"), (12, "s"), (13, "s"), (14, "s")]
MORPH_LAYOUT = [(0, 4), (1, 4), (2, 4), (3, 4), (4, 1), (5, 1), (6, 1), (7, 1), (9, 2),
                (10, 4), (11, 2), (12, 2), (13, 2), (14, 2), (15, 2), (16, 1), (17, 1),
                (18, 1), (20, 1), (21, 1), (22, "s"), (23, "s"), (24, 4), (25, 4), (26, 4),
                (27, 2), (28, 2), (29, 2)]
BUTTON_LAYOUT = [(0, 4), (1, 4), (2, 4), (3, "s"), (4, 4), (6, 1), (7, 2), (8, 2), (10, 2)]


def read_form(frx):
    data = frx.read_bytes()
    ole = olefile.OleFileIO(data[data.find(OLE_MAGIC):])
    f, o = ole.openstream("f").read(), ole.openstream("o").read()
    ole.close()

    cb_form, mask = struct.unpack_from("<HI", f, 2)
    r = Block(f, 8)
    props = r.read(mask, FORM_LAYOUT)
    assert mask & (1 << 10), "the form has no DisplayedSize"
    if 19 in props:
        r.string(props[19])
    form_w, form_h = r.pair()
    assert not mask & ((1 << 15) | (1 << 21)), "this reader does not walk a form picture"
    pos = 4 + cb_form
    if mask & (1 << 20):                                # the form's font
        pos += 16 + 1 + 2 + 1 + 2 + 4
        pos += 1 + f[pos]
    if not props.get(6, 0) & (1 << 15):                 # a class table
        count = struct.unpack_from("<H", f, pos)[0]
        assert count == 0, "this reader does not walk a class table"
        pos += 2
    count_of_sites = struct.unpack_from("<I", f, pos)[0]
    pos += 8
    start, covered = pos, 0
    while covered < count_of_sites:
        t = f[pos + 1]
        if t & 0x80:
            covered += t & 0x7F
            pos += 3
        else:
            covered += 1
            pos += 2
    pos += (-(pos - start)) % 4

    controls, opos = {}, 0
    for _ in range(count_of_sites):
        cb_site, smask = struct.unpack_from("<HI", f, pos + 2)
        s = Block(f, pos + 8)
        sp = s.read(smask, SITE_LAYOUT)
        name = s.string(sp[0])
        if 1 in sp:
            s.string(sp[1])
        left, top = s.pair() if smask & (1 << 8) else (0, 0)
        pos += 4 + cb_site

        size_of_object = sp.get(5, 0)
        obj = o[opos:opos + size_of_object]
        opos += size_of_object
        c = {"type": sp.get(7), "left": pt(left), "top": pt(top)}
        if c["type"] == TEXTBOX:
            omask = struct.unpack_from("<Q", obj, 4)[0]
            b = Block(obj, 12)
            p = b.read(omask, MORPH_LAYOUT)
            assert omask & (1 << 8), "the text box has no Size"
            w, h = b.pair()
            bits = p.get(0, 0x2C80081B)                 # VariousPropertyBits; its default
            c.update(width=pt(w), height=pt(h), scrollbars=p.get(5, 0), maxlength=p.get(3, 0),
                     enabled=bool(bits & (1 << 1)), locked=bool(bits & (1 << 2)),
                     wordwrap=bool(bits & (1 << 23)), autosize=bool(bits & (1 << 28)),
                     multiline=bool(bits & (1 << 31)))
        elif c["type"] == COMMANDBUTTON:
            omask = struct.unpack_from("<I", obj, 4)[0]
            b = Block(obj, 8)
            p = b.read(omask, BUTTON_LAYOUT)
            if 3 in p:
                c["caption"] = b.string(p[3])
            w, h = b.pair()
            c.update(width=pt(w), height=pt(h))
        controls[name] = c
    return pt(form_w), pt(form_h), controls


@pytest.fixture
def layout(repo_root):
    return read_form(Path(repo_root) / "src" / "forms" / f"{FORM}.frx")


def frm_text(repo_root):
    return (Path(repo_root) / "src" / "forms" / f"{FORM}.frm").read_bytes().decode("latin-1")


def test_the_reader_reads_361_as_measured_on_10_28(repo_root):
    """A canary on the reader: 361's box was measured on the build box as 254.4 points high."""
    _, _, controls = read_form(Path(repo_root) / "src" / "forms" / "Sh_Prodnote_Info_Form.frx")
    assert controls["Info_Text"]["type"] == TEXTBOX
    assert round(controls["Info_Text"]["height"], 1) == 254.4


def test_the_frx_size_matches_the_frm_header(repo_root, layout):
    """The form's size lives in the .frx too, and Word draws the .frx's (9/22/2026)."""
    form_w, form_h, _ = layout
    frm = frm_text(repo_root)
    header_w = float(re.search(r"ClientWidth\s+=\s+([\d.]+)", frm).group(1)) / 20
    header_h = float(re.search(r"ClientHeight\s+=\s+([\d.]+)", frm).group(1)) / 20
    assert abs(form_w - header_w) < 0.1 and abs(form_h - header_h) < 0.1


def test_the_note_is_a_text_box_that_can_scroll(layout):
    """A Label never scrolls, AutoSize defeats scrolling, a disabled box hides its bar."""
    box = layout[2]["Info_Text"]
    assert box["type"] == TEXTBOX
    assert box["multiline"] and box["wordwrap"]
    assert box["scrollbars"] in (2, 3)                  # fmScrollBarsVertical or Both
    assert box["locked"] and box["enabled"]
    assert not box["autosize"]
    assert box["maxlength"] == 0


def test_the_box_and_okay_are_inside_the_form(layout):
    """Nothing clipped: the box (and the scroll bar on its right edge) ends inside the form,
    and Okay is below the box, also inside."""
    form_w, form_h, c = layout
    box, okay = c["Info_Text"], c["Close_Me"]
    assert okay["type"] == COMMANDBUTTON and okay["caption"] == "Okay"
    assert box["left"] >= 0 and box["top"] >= 0
    assert box["left"] + box["width"] <= form_w
    assert box["top"] + box["height"] <= okay["top"]
    assert okay["top"] + okay["height"] <= form_h
    assert okay["left"] + okay["width"] <= form_w


def test_the_box_gets_the_focus_once_so_its_scroll_bar_is_drawn(repo_root):
    frm = frm_text(repo_root)
    code = "\r\n".join(l for l in frm.split("\r\n") if not l.strip().startswith("'"))
    m = re.search(r"Private Sub UserForm_Activate\(\)\r\n(.*?)End Sub", code, re.S)
    assert m, "no UserForm_Activate: the scroll bar is not drawn until the box has had the focus"
    body = [l.strip() for l in m.group(1).split("\r\n") if l.strip()]
    assert body == ["With Info_Text", ".SetFocus", ".selStart = 0", ".CurLine = 0", "End With",
                    "Close_Me.SetFocus"]

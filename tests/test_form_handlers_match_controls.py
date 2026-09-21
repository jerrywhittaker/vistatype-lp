"""Every control event handler in a UserForm must name a control that is on the form.

A handler whose control has been renamed or removed does not raise anything. It is simply
never called, so the button does nothing at all, on every machine, with no error and nothing
in the log. Nothing else in the build looks at control names, and nothing here compiles VBA.
That is how "Pictures to Color or Grayscale" on Lp_Bakgrnd_Picture_Menu_Form went dead in
3.0.448 to 3.0.471: the handler was renamed to GrayscaleButton_Click, but the control in the
.frx was still PicturesColorGrayscaleButton. See docs/Reported-Errors.md.

The control names live only in the .frx, which is binary: a 24-byte header, then an OLE
compound file. Each `f` stream in it (the form's own, and one per Frame or MultiPage page) is
a FormControl as laid out in [MS-OFORMS]. This test walks that structure to the list of sites
and reads each site's name by its stored length. It does not search the bytes for text --
a name can turn up inside a picture or a caption, and a substring search would pass a
handler whose control is gone.

Needs olefile (`sudo apt install python3-olefile`), which tools/lib/decompress_vba.py already
uses. Without it this file is skipped, loudly, rather than failing `make test`.
"""
import pathlib
import re
import struct

import pytest

olefile = pytest.importorskip(
    "olefile", reason="needs olefile to read the .frx files: sudo apt install python3-olefile")

ROOT = pathlib.Path(__file__).resolve().parents[1]
FORMS = ROOT / "src" / "forms"

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# Font GUIDs a FormControl's StreamData can carry ([MS-OFORMS] 2.4.5, 2.4.6), in the byte
# order they are stored.
STDFONT_GUID = bytes.fromhex("0352E30B918FCE119DE300AA004BB851")
TEXTPROPS_GUID = bytes.fromhex("2009C2AF4EDACE11B94300AA006887B4")

# Events a control raises in a UserForm. A Sub named <Name>_<one of these> is taken to be a
# control's handler. UserForm_* is the form itself and is left out.
CONTROL_EVENTS = {
    "AfterUpdate", "BeforeDragOver", "BeforeDropOrPaste", "BeforeUpdate", "Change", "Click",
    "DblClick", "DropButtonClick", "Enter", "Error", "Exit", "KeyDown", "KeyPress", "KeyUp",
    "MouseDown", "MouseMove", "MouseUp", "SpinDown", "SpinUp", "Scroll", "Zoom", "AddControl",
    "RemoveControl", "Layout",
}

HANDLER = re.compile(
    r"^[ \t]*(?:Private[ \t]+|Public[ \t]+)?Sub[ \t]+([A-Za-z][A-Za-z0-9_]*)_([A-Za-z]+)[ \t]*\(",
    re.IGNORECASE | re.MULTILINE)


class Reader:
    def __init__(self, data):
        self.data = data
        self.pos = 0

    def take(self, n):
        if self.pos + n > len(self.data):
            raise ValueError(f"ran off the end of the f stream at {self.pos} (+{n})")
        chunk = self.data[self.pos:self.pos + n]
        self.pos += n
        return chunk

    def u8(self):
        return self.take(1)[0]

    def u16(self):
        return struct.unpack("<H", self.take(2))[0]

    def u32(self):
        return struct.unpack("<I", self.take(4))[0]

    def align(self, base, n):
        """Skip padding so the position is a multiple of n counted from base."""
        rem = (self.pos - base) % n
        if rem:
            self.take(n - rem)


def _skip_picture(r):
    r.take(16)                                   # GUID
    preamble, size = r.u32(), r.u32()
    if preamble != 0x0000746C:
        raise ValueError(f"picture preamble {preamble:#x}, expected 0x746c")
    r.take(size)


def _skip_font(r):
    guid = r.take(16)
    if guid == STDFONT_GUID:
        r.take(1 + 2 + 1 + 2 + 4)                # version, charset, flags, weight, height
        r.take(r.u8())                           # face name
    elif guid == TEXTPROPS_GUID:
        r.take(2)                                # minor, major version
        r.take(r.u16())
    else:
        raise ValueError(f"unknown font GUID {guid.hex()}")


def _site_name(site):
    """The Name of one OleSiteConcrete ([MS-OFORMS] 2.2.10.12), or None if it has none."""
    r = Reader(site)
    r.take(2)                                    # version
    r.u16()                                      # cbSite
    mask = r.u32()
    base = r.pos
    name_len = None
    # SiteDataBlock, in order. Only fName's length is kept; the rest is walked past so the
    # ExtraDataBlock, where the name itself is, can be found.
    for bit, size in ((0, 4), (1, 4), (2, 4), (3, 4), (4, 4), (5, 4),
                      (6, 2), (7, 2), (9, 2),
                      (11, 4), (12, 4), (13, 4), (14, 4)):
        if mask & (1 << bit):
            r.align(base, size)
            value = r.u32() if size == 4 else r.u16()
            if bit == 0:
                name_len = value
    r.align(base, 4)
    if name_len is None:
        return None
    count, compressed = name_len & 0x7FFFFFFF, name_len & 0x80000000
    raw = r.take(count)
    return raw.decode("latin-1") if compressed else raw.decode("utf-16-le")


def control_names_in_f_stream(data):
    """The names of every site in one FormControl `f` stream ([MS-OFORMS] 2.2.10.1)."""
    r = Reader(data)
    r.take(2)                                    # minor, major version
    cb_form = r.u16()
    mask = r.u32()
    start = r.pos
    # BooleanProperties follows BackColor, ForeColor and NextAvailableID, each 4 bytes when
    # present. It is needed for one bit: whether a class table was saved.
    offset = start + 4 * sum(1 for bit in (1, 2, 3) if mask & (1 << bit))
    boolean_props = struct.unpack_from("<I", data, offset)[0] if mask & (1 << 6) else 0
    r.pos = 4 + cb_form                          # past DataBlock and ExtraDataBlock

    if mask & (1 << 15):                         # fMouseIcon
        _skip_picture(r)
    if mask & (1 << 20):                         # fFont
        _skip_font(r)
    if mask & (1 << 21):                         # fPicture
        _skip_picture(r)

    if not boolean_props & (1 << 15):            # FORM_FLAG_DONTSAVECLASSTABLE clear
        for _ in range(r.u16()):
            r.take(2)
            r.take(r.u16())

    count_of_sites = r.u32()
    r.u32()                                      # CountOfBytes: depths and sites together
    # SiteDepthsAndTypes: a depth byte, then a type byte -- or, with its high bit set, a
    # repeat count followed by the type. Walked until every site is accounted for, then
    # padded to 4 bytes counted from where the list began.
    depths_start = r.pos
    covered = 0
    while covered < count_of_sites:
        r.u8()                                   # depth
        type_or_count = r.u8()
        if type_or_count & 0x80:
            covered += type_or_count & 0x7F
            r.u8()                               # type
        else:
            covered += 1
    r.align(depths_start, 4)

    names = []
    for _ in range(count_of_sites):
        head = r.pos
        r.take(2)
        cb_site = r.u16()
        r.pos = head
        site = r.take(4 + cb_site)
        name = _site_name(site)
        if name is None or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", name):
            raise ValueError(f"site {len(names)} has no usable name: {name!r}")
        names.append(name)
    return names


def control_names(frx_path):
    data = frx_path.read_bytes()
    start = data.find(OLE_MAGIC)
    if start < 0:
        raise ValueError(f"{frx_path.name}: no OLE compound file inside")
    ole = olefile.OleFileIO(data[start:])
    try:
        names = set()
        for entry in ole.listdir():
            if entry[-1] == "f":
                names.update(control_names_in_f_stream(ole.openstream(entry).read()))
        return names
    finally:
        ole.close()


def control_handlers(frm_path):
    """(line number, control name, event) for each Sub that handles a control's event."""
    text = frm_path.read_bytes().decode("cp1252")
    found = []
    for m in HANDLER.finditer(text):
        name, event = m.group(1), m.group(2)
        if name.lower() == "userform":
            continue
        if event.lower() not in {e.lower() for e in CONTROL_EVENTS}:
            continue
        line = text.count("\n", 0, m.start()) + 1
        found.append((line, name, event))
    return found


FORM_FILES = sorted(FORMS.glob("*.frm"))


def test_there_are_forms_to_check():
    assert len(FORM_FILES) > 40


def test_the_parser_reads_a_known_form():
    """A canary on the parser itself: a form whose controls are known."""
    names = control_names(FORMS / "Lp_Bakgrnd_Picture_Menu_Form.frx")
    assert {"ExitButton", "BackgroundColorButton", "PicturesColorGrayscaleButton"} <= names


def test_nested_controls_are_read_too():
    """Controls inside a Frame or MultiPage live in their own `f` stream, not the form's."""
    top = control_names_in_f_stream(
        olefile.OleFileIO(_ole_bytes(FORMS / "Lp_Table_Convert_Options_Form.frx"))
        .openstream("f").read())
    assert control_names(FORMS / "Lp_Table_Convert_Options_Form.frx") > set(top)


def _ole_bytes(frx_path):
    data = frx_path.read_bytes()
    return data[data.find(OLE_MAGIC):]


@pytest.mark.parametrize("frm", FORM_FILES, ids=lambda p: p.stem)
def test_every_handler_names_a_control_on_its_form(frm):
    frx = frm.with_suffix(".frx")
    handlers = control_handlers(frm)
    if not handlers:
        return
    assert frx.exists(), f"{frm.name} has control handlers but no {frx.name}"
    names = {n.lower() for n in control_names(frx)}
    dead = [f"  line {line}: {name}_{event} -- there is no control named {name}"
            for line, name, event in handlers if name.lower() not in names]
    assert not dead, (
        f"{frm.name}: these handlers will never run, because their control is not on the "
        f"form (a renamed or removed control):\n" + "\n".join(dead))

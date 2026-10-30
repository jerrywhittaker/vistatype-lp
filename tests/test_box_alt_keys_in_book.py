"""A box that stays open answers its own Alt keys from the book, too.

Jerry, 10/30/2026: on Type Fill-In Lines (357), Alt+F did nothing in the box and opened Word's
File tab. A form's Alt keys fire only while the form has the keyboard, and a box that stays open
hands the keyboard to the book (docs/UI-Conventions.md). The fix binds the newest box's Alt keys
in Word as well, the way F6 is bound, each to a Sh_Box_Alt_<key> macro that hands the key to the
box. Resize Pictures in a Selected Range (375) gets the same, and the same day (Jerry, "do all
five") the TOC box (354), Horizontal List to Vertical (348), the two $pg menus (364, 366) and the
question that stays open (402). See docs/Reported-Errors.md.

Everything here fails silently in Word: a letter left off the list is a key that still opens a
ribbon tab; a letter that is not on the form presses nothing visible; a Case that does something
other than the button does the wrong thing to the book; and a binding not cleared on close
leaves Alt+F dead in Word after the box has gone.

THE LETTERS ARE READ FROM THE .frx PROPERTY BLOCKS, not from strings in it ([MS-OFORMS]: the
site list in the `f` stream says each control's type and how many bytes of the `o` stream are
its own, and each control's property mask says whether an Accelerator is stored and where). A
letter set in the form's UserForm_Initialize wins over the .frx, so those are applied on top.
The parser is checked against the 375 letters measured on the build box on 9/29/2026.
"""
import re
import struct

from test_form_handlers_match_controls import (OLE_MAGIC, Reader, _site_name, _skip_font,
                                               _skip_picture, olefile)
from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures

BOXES = {
    # the name in Sh_Box_AltKeys: (who on the shared list, form, the box's _AltKey sub). The two
    # $pg menus are both "pg" on the list and have their own letters (Sh_Box_AltName).
    "fil": ("fil", "Lp_Type_Fill_In_Line_Form", "Lp_Fil_AltKey"),
    "rst": ("rst", "Lp_Same_Pic_Range_Form", "Lp_Rst_AltKey"),
    "tocb": ("tocb", "Lp_TOC_Format_And_Color_Form", "Lp_Tocb_AltKey"),
    "hv": ("hv", "Lp_Horz_To_Vert_List_Form", "Lp_Hvb_AltKey"),
    "pg2": ("pg", "Sh_Valid_Ref_Pg_No_2_Form", "Sh_PgVal_AltKey"),
    "pg4": ("pg", "Sh_Valid_Ref_Pg_No_4_Form", "Sh_PgVal_AltKey"),
    "msg": ("msg", "Sh_Message_Form", "Sh_Ask_Modeless_AltKey"),
}

# 402 is its own instance of Sh_Message_Form, held in a variable: Me on the form is this there.
INSTANCE = {"Sh_Message_Form": "Sh_AskBox_Form"}

# ClsidCacheIndex ([MS-OFORMS] 2.4.1) of the controls whose Alt key presses or ticks something,
# with the bit in the control's property mask that says an Accelerator is stored, the size of
# the mask, and the DataBlock fields stored before the Accelerator as (mask bit, bytes).
_COMMAND_BUTTON = (8, 4, [(0, 4), (1, 4), (2, 4), (3, 4), (4, 4), (6, 1), (7, 2)])
_MORPH = (29, 8, [(0, 4), (1, 4), (2, 4), (3, 4), (4, 1), (5, 1), (6, 1), (7, 1), (9, 2),
                  (10, 4), (11, 2), (12, 2), (13, 2), (14, 2), (15, 2), (16, 1), (17, 1),
                  (18, 1), (20, 1), (21, 1), (22, 4), (23, 4), (24, 4), (25, 4), (26, 4),
                  (27, 2), (28, 2)])
PRESSABLE = {17: ("CommandButton", _COMMAND_BUTTON), 26: ("CheckBox", _MORPH),
             27: ("OptionButton", _MORPH), 28: ("ToggleButton", _MORPH)}


def _site(site):
    """(name, bytes of the o stream, ClsidCacheIndex) for one OleSiteConcrete."""
    r = Reader(site)
    r.take(2)
    r.u16()
    mask = r.u32()
    base = r.pos
    values = {}
    for bit, size in ((0, 4), (1, 4), (2, 4), (3, 4), (4, 4), (5, 4), (6, 2), (7, 2), (9, 2),
                      (11, 4), (12, 4), (13, 4), (14, 4)):
        if mask & (1 << bit):
            r.align(base, size)
            values[bit] = r.u32() if size == 4 else r.u16()
    return _site_name(site), values.get(5, 0), values.get(7, 0x7FFF)


def _sites(data):
    """Every site in one `f` stream, in order - the walk test_form_handlers_match_controls makes."""
    r = Reader(data)
    r.take(2)
    cb_form = r.u16()
    mask = r.u32()
    offset = r.pos + 4 * sum(1 for bit in (1, 2, 3) if mask & (1 << bit))
    boolean_props = struct.unpack_from("<I", data, offset)[0] if mask & (1 << 6) else 0
    r.pos = 4 + cb_form
    if mask & (1 << 15):
        _skip_picture(r)
    if mask & (1 << 20):
        _skip_font(r)
    if mask & (1 << 21):
        _skip_picture(r)
    if not boolean_props & (1 << 15):
        for _ in range(r.u16()):
            r.take(2)
            r.take(r.u16())
    count = r.u32()
    r.u32()
    depths_start = r.pos
    covered = 0
    while covered < count:
        r.u8()
        type_or_count = r.u8()
        if type_or_count & 0x80:
            covered += type_or_count & 0x7F
            r.u8()
        else:
            covered += 1
    r.align(depths_start, 4)
    out = []
    for _ in range(count):
        head = r.pos
        r.take(2)
        cb_site = r.u16()
        r.pos = head
        out.append(_site(r.take(4 + cb_site)))
    return out


def _accelerator(obj, layout):
    bit, mask_size, before = layout
    r = Reader(obj)
    r.take(2)
    r.u16()
    mask = r.u32() if mask_size == 4 else struct.unpack("<Q", r.take(8))[0]
    base = r.pos
    if not mask & (1 << bit):
        return ""
    for field_bit, size in before:
        if mask & (1 << field_bit):
            r.align(base, size)
            r.take(size)
    r.align(base, 2)
    return chr(r.u16())


def frx_accelerators(frx_path):
    """{control name: (type, letter)} for every button, tick box and option button with a letter."""
    data = frx_path.read_bytes()
    ole = olefile.OleFileIO(data[data.find(OLE_MAGIC):])
    try:
        found = {}
        for entry in ole.listdir():
            if entry[-1] != "f":
                continue
            o_path = "/".join(entry[:-1] + ["o"])
            objects = ole.openstream(o_path).read() if ole.exists(o_path) else b""
            pos = 0
            for name, size, kind in _sites(ole.openstream(entry).read()):
                obj = objects[pos:pos + size]
                pos += size
                if kind in PRESSABLE and size:
                    kind_name, layout = PRESSABLE[kind]
                    letter = _accelerator(obj, layout)
                    if letter:
                        found[name] = (kind_name, letter)
        return found
    finally:
        ole.close()


def form_accelerators(repo_root, form):
    """The letters the transcriber sees: the .frx's, with UserForm_Initialize's on top."""
    found = frx_accelerators(repo_root / "src" / "forms" / f"{form}.frx")
    code = "\n".join(form_code(repo_root, form))
    init = re.search(r"Sub UserForm_Initialize\(\)(.*?)End Sub", code, re.S)
    for ctl, letter in re.findall(r'(\w+)\.Accelerator = "(\w)"', init.group(1) if init else ""):
        kind = found.get(ctl, ("CommandButton", ""))[0]
        found[ctl] = (kind, letter)
    return {ctl: (kind, letter.upper()) for ctl, (kind, letter) in found.items()}


def shared_module(repo_root):
    return (repo_root / "src" / "vba" / "ShNonModalMessage.bas").read_bytes().decode("latin-1")


def all_procedures(repo_root):
    """Both modules' procedures: the _AltKey subs live in either."""
    procs = procedures(module_text(repo_root))
    procs.update(procedures(shared_module(repo_root)))
    return procs


def alt_keys(repo_root, who):
    body = "\n".join(procedures(shared_module(repo_root))["Sh_Box_AltKeys"])
    m = re.search(rf'Case "{who}"\s*\n\s*Sh_Box_AltKeys = "(\w*)"', body)
    return m.group(1) if m else ""


def case_body(lines, key):
    """The statements under the single-key Case "<key>" of a Select, up to the next Case."""
    text = "\n".join(lines)
    m = re.search(rf'^\s*Case "{key}"\s*$(.*?)(?=^\s*Case |^\s*End Select)', text, re.S | re.M)
    assert m, f'no Case "{key}"'
    return m.group(1)


def click_body(repo_root, form, ctl):
    """The _Click's statements, comments off (form_code strips them), Dim lines dropped."""
    code = "\n".join(form_code(repo_root, form))
    m = re.search(rf"Sub {ctl}_Click\(\)(.*?)End Sub", code, re.S)
    assert m, f"{form}: {ctl} has no _Click"
    return [line.strip() for line in m.group(1).splitlines()
            if line.strip() and not line.strip().startswith("Dim ")]


def _chain(lines):
    """[(test, target, value)] of an If / ElseIf / Else chain that sets one variable."""
    out = []
    for n, line in enumerate(lines):
        m = re.match(r"(?:If|ElseIf) (\w+) Then$", line) or re.match(r"(Else)$", line)
        if m:
            target, _, value = lines[n + 1].partition(" = ")
            out.append((m.group(1), target.strip(), value.strip()))
    return out


def _reads(body, form):
    """{variable: control} for each read of an option button or tick box off the form."""
    return dict(re.findall(rf"(\w+) = \({form}\.(\w+)\.Value = True\)", body))


# --- the parser itself


def test_the_parser_reads_the_letters_measured_on_the_build_box(repo_root):
    """375's letters, measured in Word on the build box 9/29/2026 (test_375_attach_picture_boxes)."""
    found = frx_accelerators(repo_root / "src" / "forms" / "Lp_Same_Pic_Range_Form.frx")
    for ctl, letter in (("LeftAlignButton", "L"), ("CenterButton", "C"), ("ApplyButton", "A"),
                        ("DoneButton", "D"), ("LastSizeButton", "U")):
        assert found[ctl][1].upper() == letter, (ctl, found.get(ctl))
    assert found["JoinNextParaButton"][0] == "CheckBox"
    assert found["LeftAlignButton"][0] == "OptionButton"


# --- the list of keys is the form's letters, exactly


def test_each_box_binds_exactly_its_own_letters(repo_root):
    for who, (_, form, _) in BOXES.items():
        keys = alt_keys(repo_root, who)
        assert keys, f'Sh_Box_AltKeys has no list for "{who}"'
        letters = {letter for _, letter in form_accelerators(repo_root, form).values()}
        assert len(keys) == len(set(keys)), f"{who}: a letter is listed twice: {keys}"
        assert set(keys) == letters, (
            f"{form}: the box's letters are {sorted(letters)}, the list in Sh_Box_AltKeys is "
            f"{sorted(keys)} - a missing one still goes to Word, an extra one presses nothing")


def test_354_gives_cancel_d_before_format_takes_c(repo_root):
    """The .frx has Cancel on C. Initialize must move Cancel to D before Clean and format the TOC
    takes C, or two controls hold C at once, however briefly."""
    code = "\n".join(form_code(repo_root, "Lp_TOC_Format_And_Color_Form"))
    init = re.search(r"Sub UserForm_Initialize\(\)(.*?)End Sub", code, re.S).group(1)
    assert init.index('CancelButton.Accelerator = "D"') < init.index('FormatTheTOCButton.Accelerator = "C"')


def test_357_has_the_letters_jerry_sees(repo_root):
    found = form_accelerators(repo_root, "Lp_Type_Fill_In_Line_Form")
    assert found["FillToRightMargin"][1] == "F"
    assert found["CancelButton"][1] == "D"
    for n in range(1, 10):
        assert found[f"CommandButton{n}"][1] == str(n)
    assert found["CommandButton10"][1] == "0"


# --- every key reaches the box and does what the control does


def test_the_letters_jerry_was_given(repo_root):
    """Checked by hand against the .frx and UserForm_Initialize, 10/30/2026: 354's Cancel is D in
    code, and Clean and format the TOC is C (Jerry, 10/30/2026: it was F, Word's File tab)."""
    toc = form_accelerators(repo_root, "Lp_TOC_Format_And_Color_Form")
    assert toc["CancelButton"][1] == "D"
    assert toc["FormatTheTOCButton"] == ("OptionButton", "C")
    for name, keys in (("tocb", "CAROD"), ("hv", "ODSALT"), ("pg2", "LDW"), ("pg4", "RDW"),
                       ("msg", "OC")):
        assert set(alt_keys(repo_root, name)) == set(keys), name


def test_every_key_has_its_macro(repo_root):
    shared = procedures(shared_module(repo_root))
    for who in BOXES:
        for key in alt_keys(repo_root, who):
            name = f"Sh_Box_Alt_{key}"
            assert name in shared, f"Alt+{key} is bound to {name}, which does not exist"
            first = shared[name][0]
            assert first.startswith("Public Sub"), f"{name} must be Public: a key binding cannot reach a Private sub"
            assert f'Sh_Box_AltKey "{key}"' in "\n".join(shared[name])


def test_the_key_goes_to_the_newest_box(repo_root):
    body = "\n".join(procedures(shared_module(repo_root))["Sh_Box_AltKey"])
    for who, sub in {(w, s) for w, _, s in BOXES.values()}:
        assert re.search(rf'Case "{who}"\s*\n\s*{sub} key', body), f"Alt keys cannot reach {sub}"
    assert re.search(r'Case ""\s*\n(?:.*\n)*?\s*Sh_Box_UnbindKeys', body), (
        "with no box up, the keys are not given back to Word")
    assert "Sh_Box_Top(Sh_Box_Stack)" in body


def _calls(text, proc):
    """(arguments) of each call to proc in some VBA statements."""
    return [m.group(1).strip() for m in re.finditer(rf"\b{proc}\b(?!_)([^\n]*)", text)]


def _okay_with_choices(sub, letter, form, ctl, click, body):
    """An Okay whose _Click reads the choices on the box into an If chain, then calls one macro."""
    reads = _reads(body, form)
    want = [(t, v) for t, _, v in _chain(click)]
    got = [(reads.get(t, t), v) for t, _, v in _chain([s.strip() for s in body.splitlines()])]
    assert got == want, (f"{sub}: Alt+{letter} does not choose as {ctl} does: the box reads "
                         f"{want}, the key {got}")
    proc, _, args = click[-1].partition(" ")
    calls = _calls(body, proc)
    assert calls, f"{sub}: Alt+{letter} does not run {proc}, as {ctl} does"
    args = [a.strip() for a in args.split(",")]
    have = [a.strip() for a in calls[-1].split(",")]
    assert len(have) == len(args), (sub, letter, click[-1], calls)
    for w, h in zip(args, have):
        m = re.fullmatch(r"\((\w+) = True\)", w)
        if m:
            assert reads.get(h) == m.group(1), f"{sub}: Alt+{letter} does not pass {m.group(1)} as {ctl} does"
        else:
            assert h == w, (sub, letter, click[-1], calls)
    # a failed read presses nothing
    assert re.search(r"readOk = \(Err\.Number = 0\)", body) and "If Not readOk Then Exit Sub" in body, (
        f"{sub}: Alt+{letter} goes on when the box could not be read")
    assert body.index("If Not readOk Then Exit Sub") < body.index(proc + " "), sub


def test_each_key_does_what_its_control_does(repo_root):
    macros = all_procedures(repo_root)
    for name, (_, form, sub) in BOXES.items():
        lines = macros[sub]
        assert lines[0].startswith("Public Sub"), f"{sub} is called from ShNonModalMessage"
        assert re.search(rf"Sub {sub}\(ByVal key As String\)", lines[0])
        for ctl, (kind, letter) in form_accelerators(repo_root, form).items():
            body = case_body(lines, letter)
            if kind == "CommandButton":
                click = click_body(repo_root, form, ctl)
                if any(line.startswith("If ") for line in click):
                    _okay_with_choices(sub, letter, form, ctl, click, body)
                    continue
                for stmt in click:
                    if stmt.startswith("Me."):
                        # 402: Me is the instance Sh_Ask_Modeless waits on
                        inst = INSTANCE.get(form, form)
                        assert f"{inst}.{stmt[3:]}" in body, f"{sub}: Alt+{letter} does not {stmt} as {ctl} does"
                        continue
                    if re.match(r"\w+ = ", stmt):
                        assert stmt in body, f"{sub}: Alt+{letter} does not set {stmt} as {ctl} does"
                        continue
                    proc, _, args = stmt.partition(" ")
                    args = [a.strip() for a in args.split(",")] if args else []
                    calls = _calls(body, proc)
                    assert calls, f"{sub}: Alt+{letter} does not run {proc}, as {ctl} does"
                    got = [a.strip() for a in calls[0].split(",")] if calls[0] else []
                    assert len(got) == len(args), (sub, letter, stmt, calls)
                    for want, have in zip(args, got):
                        if "." in want:
                            # a control on the box: read off the form, directly or into a variable
                            src = f"{form}.{want}"
                            assert have == src or re.search(rf"\b{have}\s*=\s*{re.escape(src)}", body), (
                                f"{sub}: Alt+{letter} does not pass {want} as {ctl} does")
                        else:
                            assert have == want, (sub, letter, stmt, calls)
            elif kind in ("CheckBox", "ToggleButton"):
                assert f".{ctl}.Value = Not (.{ctl}.Value = True)" in body, (
                    f"{sub}: Alt+{letter} does not tick or clear {ctl}")
            elif kind == "OptionButton":
                assert f".{ctl}.Value = True" in body, f"{sub}: Alt+{letter} does not choose {ctl}"


def test_use_last_size_does_nothing_until_it_is_live(repo_root):
    """On the box the button is dead until a size has been used, and so is its Alt key."""
    body = case_body(procedures(module_text(repo_root))["Lp_Rst_AltKey"], "U")
    assert "Lp_Same_Pic_Range_Form.LastSizeButton.Enabled" in body
    assert re.search(r"If live Then Lp_Rst_Apply_Last_Size", body)


def test_the_form_is_touched_only_while_the_box_is_up(repo_root):
    """Touching a member of a form that is down creates it and runs its Initialize."""
    macros = all_procedures(repo_root)
    for sub, guard_line in (("Lp_Fil_AltKey", "If Not Lp_Fil_IsOn Then"),
                            ("Lp_Rst_AltKey", "If Not Lp_Rst_IsOn Then"),
                            ("Lp_Tocb_AltKey", "If Not Lp_Tocb_IsOn Then"),
                            ("Lp_Hvb_AltKey", "If Not Lp_Hvb_IsOn Then"),
                            ("Sh_PgVal_AltKey", "If menu = 0 Then"),
                            ("Sh_Ask_Modeless_AltKey", "If Not Sh_AskBox_Up Then")):
        lines = macros[sub]
        guard = next(n for n, c in enumerate(lines) if guard_line in c)
        touch = [n for n, c in enumerate(lines) if "_Form." in c or "With Lp_" in c]
        assert touch and min(touch) > guard, f"{sub} reaches the form before it asks {flag}"


# --- bound with F6, released on every way out


def test_bound_and_released_with_f6(repo_root):
    procs = procedures(shared_module(repo_root))
    text = shared_module(repo_root)
    assert 'Private Const SH_BOX_ALT_MACRO As String = "LPandBRL.ShNonModalMessage.Sh_Box_Alt_"' in text

    bind = "\n".join(procs["Sh_Box_BindKeys"])
    assert "BuildKeyCode(wdKeyAlt, Asc(ch))" in bind and "SH_BOX_ALT_MACRO & ch" in bind
    assert "Sh_Box_AltKeys(who)" in bind
    assert "who = Sh_Box_AltName(Sh_Box_Top(Sh_Box_Stack), Sh_PgVal_MenuIsOn)" in bind, (
        "the $pg menus' keys must follow the menu on screen")
    # the old box's keys go before the new box's come in - F is on both 357 and 375
    assert bind.index("Sh_Box_Clear_AltKeys") < bind.index("SH_BOX_ALT_MACRO & ch")
    assert "CustomizationContext = ThisDocument" in bind and "ThisDocument.Saved = True" in bind

    unbind = "\n".join(procs["Sh_Box_UnbindKeys"])
    assert "Sh_Box_Clear_AltKeys" in unbind

    clear = "\n".join(procs["Sh_Box_Clear_AltKeys"])
    assert '"Sh_Box_Alt_"' in clear and "KeyBindings(i).Clear" in clear
    assert "Step -1" in clear, "a forward loop steps over the binding after each one cleared"
    assert 'Sh_Box_AltFor = ""' in clear

    closed = "\n".join(procs["Sh_Box_Closed"])
    assert re.search(r"If Len\(Sh_Box_Stack\) = 0 Then\s*\n\s*Sh_Box_UnbindKeys\s*\n\s*Else\s*\n"
                     r"\s*Sh_Box_BindKeys", closed), (
        "a box closing with another still up must hand the Alt keys to that one")

    toggle = "\n".join(procs["Sh_Box_ToggleFocus"])
    assert "Sh_Box_UnbindKeys" in toggle


def test_every_box_leaves_the_list_on_every_way_out(repo_root):
    """Release depends on Sh_Box_Closed being reached; each box's one way out reaches it."""
    macros = all_procedures(repo_root)
    assert any('Sh_Box_Closed "fil"' in c for c in macros["Lp_Fil_Done"])
    assert any(re.search(r"^\s*Lp_Rst_UnbindKeys\s*$", c) for c in macros["Lp_Rst_Finish"])
    assert any(re.search(r"^\s*Lp_Tocb_UnbindKeys\s*$", c) for c in macros["Lp_Tocb_Finish"])
    assert any(re.search(r"^\s*Lp_Tocb_Finish True\s*$", c) for c in macros["Lp_Tocb_Done"])
    assert any('Sh_Box_Closed "hv"' in c for c in macros["Lp_Hvb_Done"])
    assert any(re.search(r"^\s*Sh_PgVal_UnbindKeys\s*$", c) for c in macros["Sh_PgVal_Done"])
    assert any('Sh_Box_Closed "pg"' in c for c in macros["Sh_PgVal_UnbindKeys"])
    for form, done in (("Lp_Type_Fill_In_Line_Form", "Lp_Fil_Done"),
                       ("Lp_Same_Pic_Range_Form", "Lp_Rst_Done"),
                       ("Lp_TOC_Format_And_Color_Form", "Lp_Tocb_Done"),
                       ("Lp_Horz_To_Vert_List_Form", "Lp_Hvb_Done"),
                       ("Sh_Valid_Ref_Pg_No_2_Form", "Sh_PgVal_Done"),
                       ("Sh_Valid_Ref_Pg_No_4_Form", "Sh_PgVal_Done"),
                       ("Sh_Message_Form", "Me.Hide")):
        code = "\n".join(form_code(repo_root, form))
        q = re.search(r"Sub UserForm_QueryClose\(.*?\)(.*?)End Sub", code, re.S)
        assert q and done in q.group(1), f"{form}: the X does not go through {done}"


def test_402_is_left_by_the_wait_and_only_the_wait(repo_root):
    """402's Okay, Cancel, X and Alt keys only hide the box. The wait sees it go, unloads it and
    leaves the list - which binds the keys of the box underneath again (Sh_Box_Closed). Leaving
    the list from inside the Alt key would run inside the wait's DoEvents."""
    procs = procedures(shared_module(repo_root))
    ask = "\n".join(procs["Sh_Ask_Modeless"])
    loop = ask.index("Loop While stillUp")
    tail = ask[loop:]
    assert "Unload Sh_AskBox_Form" in tail and 'Sh_Box_Closed "msg"' in tail
    assert "Exit Function" not in tail.split('Sh_Box_Closed "msg"')[0], "a way out of the wait skips the list"
    alt = procs["Sh_Ask_Modeless_AltKey"]
    guard = next(n for n, c in enumerate(alt) if "If Not Sh_AskBox_Up Then" in c)
    after = "\n".join(alt[guard + 3:])
    assert "Unload" not in after and "Sh_Box_Closed" not in after, (
        "the Alt key must only hide 402, as its buttons do; the wait does the rest")


def test_each_pg_menu_swap_binds_its_own_letters(repo_root):
    """Both $pg menus are "pg" on the list; the letters follow Sh_PgVal_MenuIsOn, so every swap
    binds again AFTER the menu number has changed."""
    procs = procedures(shared_module(repo_root))
    for sub, menu in (("Sh_PgVal_SwapToTempForm", 2), ("Sh_PgVal_SwapToSourceForm", 4)):
        body = "\n".join(procs[sub])
        assert "Sh_Box_BindKeys" in body, f"{sub} leaves the other menu's letters bound"
        assert body.index(f"Sh_PgVal_MenuIsOn = {menu}") < body.index("Sh_Box_BindKeys"), sub
    name = "\n".join(procs["Sh_Box_AltName"])
    assert re.search(r'If who = "pg" Then\s*\n\s*Sh_Box_AltName = "pg" & CStr\(pgMenu\)', name)
    pg = procs["Sh_PgVal_AltKey"]
    assert "If menu = 2 Then Sh_PgVal_LocateInDocument" in "\n".join(pg)
    assert "If menu = 4 Then Sh_PgVal_ReturnToTempAndAdvance" in "\n".join(pg)


# --- an error in the box's macro shows, from the key as from the button


# The macros each Alt key runs on the box, as its buttons' _Click do. None of the buttons' _Click
# has a handler, so an error in one of these shows when the button is clicked; it must show from
# the key too. Review, 10/30/2026: Sh_Box_AltKey ran them all under On Error Resume Next, and
# Sh_PgVal_LocateInDocument and Sh_PgVal_ReturnToTempAndAdvance have no handler of their own, so
# a failure in Locate or Return from Alt+L or Alt+R went by in silence.
_CALL = re.compile(r"^(?:If .+ Then )?((?:Lp|Sh)_\w+)(?!\s*=)(?:\s|$)")


def _resume_next_after(lines, state):
    for line in lines:
        t = line.strip()
        if t.startswith("On Error Resume Next"):
            state = True
        elif t.startswith("On Error GoTo"):
            state = False
    return state


def test_the_dispatcher_swallows_no_error(repo_root):
    procs = procedures(shared_module(repo_root))
    body = [c.strip() for c in procs["Sh_Box_AltKey"]]
    assert not any(c.startswith("On Error Resume Next") for c in body), (
        "Sh_Box_AltKey must not run the box's macro under On Error Resume Next: an error in it "
        "then goes by in silence from the Alt key, while it shows from the button")
    # the two safety nets it calls guard themselves
    for sub in ("Sh_Box_BindKeys", "Sh_Box_UnbindKeys"):
        assert any(c.strip().startswith("On Error Resume Next") for c in procs[sub]), sub


def test_each_key_runs_its_macro_with_errors_showing(repo_root):
    macros = all_procedures(repo_root)
    checked = 0
    for name, (_, form, sub) in BOXES.items():
        lines = macros[sub]
        select = next(n for n, c in enumerate(lines) if "Select Case UCase$(key)" in c)
        before = _resume_next_after(lines[:select], False)
        for ctl, (kind, letter) in form_accelerators(repo_root, form).items():
            state = before
            for line in case_body(lines, letter).splitlines():
                t = line.strip()
                state = _resume_next_after([t], state)
                m = _CALL.match(t)
                if m and not m.group(1).endswith("_Form"):
                    checked += 1
                    assert not state, (
                        f"{sub}: Alt+{letter} runs {m.group(1)} under On Error Resume Next - an "
                        f"error in it is silent from the key and shows from {ctl}")
    assert checked >= 20, f"only {checked} calls found: the walk no longer sees the Case bodies"

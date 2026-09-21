"""Where a macro empties the undo list: after every way out, before the first change.

Jerry, 9/21/2026: "The undo stack should be cleared before 'Resize Pictures Used Throughout',
otherwise pressing Ctrl+Z more than once will undo changes made before the macro was run."
Lp_Resize_Same_Picture_Throughout's own work is one Ctrl+Z (a custom undo record), but a
second press carried on into the transcriber's earlier edits. See docs/Reported-Errors.md.

The clear has two ways to be in the wrong place, and both are silent:

* Too HIGH -- above a dialog the transcriber can still cancel -- and backing out, with nothing
  changed, throws away their undo list for nothing.
* Too LOW -- below the first change to the book -- and that change cannot be taken back.

Nothing here runs VBA, so this reads the source: the sub's code lines with comments taken out,
in order. It checks the real LPandBrlMacros.bas, so moving the line breaks `make test`.
"""
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[1]
BAS = ROOT / "src" / "vba" / "LPandBrlMacros.bas"

SUB = "Lp_Resize_Same_Picture_Throughout"

# A change to the book in this sub: setting a picture's size, its aspect-ratio lock, or its
# paragraph's alignment. `refW = ref.Width` is a read and must not match -- the property is on
# the LEFT of the = here.
EDIT = re.compile(r"^\S+\.(Width|Height|LockAspectRatio|Alignment)\s*=", re.IGNORECASE)


def code_lines(name):
    """The sub's code, one stripped line each, with comment lines and blank lines left out."""
    text = BAS.read_bytes().decode("latin-1")
    lines = text.split("\r\n")
    start = next(i for i, l in enumerate(lines)
                 if re.match(r"^(Public\s+|Private\s+)?Sub\s+%s\s*\(" % name, l))
    end = next(i for i in range(start + 1, len(lines)) if lines[i].startswith("End Sub"))
    out = []
    for l in lines[start + 1:end]:
        s = l.strip()
        if s and not s.startswith("'"):
            out.append(s)
    return out


def index_of(lines, pattern, what):
    for i, l in enumerate(lines):
        if re.search(pattern, l, re.IGNORECASE):
            return i
    raise AssertionError("%s: no %s found -- has the sub been rewritten?" % (SUB, what))


def test_the_undo_list_is_emptied_exactly_once():
    lines = code_lines(SUB)
    clears = [l for l in lines if re.search(r"\.UndoClear\b", l, re.IGNORECASE)]
    assert len(clears) == 1, clears


def test_it_is_the_selected_pictures_document_that_is_cleared():
    """Not whatever is active: the macro opens a hidden reader document of its own."""
    lines = code_lines(SUB)
    clear = lines[index_of(lines, r"\.UndoClear\b", "UndoClear")]
    assert clear == "ref.Range.Document.UndoClear"


def test_every_way_out_comes_before_the_clear():
    """322, Cancel on the alignment question, 323 -- backing out must not cost the undo list."""
    lines = code_lines(SUB)
    clear = index_of(lines, r"\.UndoClear\b", "UndoClear")
    record = index_of(lines, r"\.StartCustomRecord\b", "StartCustomRecord")
    exits = [i for i, l in enumerate(lines[:record]) if l.lower() == "exit sub"]
    assert exits, "no early Exit Sub found -- has the sub been rewritten?"
    late = [lines[i] for i in exits if i > clear]
    assert not late, "a way out comes after the UndoClear"
    assert index_of(lines, r"Lp_Same_Pic_Align_Form\.Show", "alignment question") < clear
    assert index_of(lines, r"VistaType LP \(323\)", "dialog 323") < clear


def test_the_clear_comes_before_the_first_change_to_the_book():
    lines = code_lines(SUB)
    clear = index_of(lines, r"\.UndoClear\b", "UndoClear")
    first_edit = next((i for i, l in enumerate(lines) if EDIT.match(l)), None)
    assert first_edit is not None, "no picture edit found -- has the sub been rewritten?"
    assert clear < first_edit, lines[first_edit]
    assert clear < index_of(lines, r"\.StartCustomRecord\b", "StartCustomRecord")


def test_the_edit_pattern_tells_a_change_from_a_read():
    """Guards the guard: if this pattern matched reads, the test above would prove nothing."""
    assert EDIT.match("s.Width = refW")
    assert EDIT.match("ref.Range.ParagraphFormat.Alignment = alignWanted")
    assert not EDIT.match("refW = ref.Width")
    assert not EDIT.match("oldLock = s.LockAspectRatio")

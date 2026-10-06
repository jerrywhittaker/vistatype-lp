"""No UserForm in src/forms says "Tabel" for "Table".

Issue 6, 10/6/2026. The hover text (ControlTipText) of the ColumnOnlyTable radio button on
Lp_Table_Convert_Options_Form read "Tabel has column heading only radio button". The issue
also named the button's caption, but Word's designer reported the caption as "Table has column
headers only" - already right. The misread came from pulling strings out of the .frx, which
packs one control's name, tip and the next field end to end.

The words a transcriber reads on a dialog are mostly in the BINARY .frx, not in the .frm text,
and Word writes them there as single-byte text or as UTF-16LE. So each file is read as bytes
and searched in both encodings. Case-sensitive on purpose: "tabel" in lower case could be the
inside of a longer name; "Tabel" at a capital is the misspelling.
"""
import pytest

TYPO = "Tabel"
ENCODINGS = ("ascii", "utf-16-le")


def form_files(repo_root):
    files = sorted((repo_root / "src" / "forms").glob("*.fr[mx]"))
    assert files, "no .frm/.frx files found in src/forms"
    return files


def hits(data):
    """The encodings in which TYPO appears in these bytes."""
    return [enc for enc in ENCODINGS if TYPO.encode(enc) in data]


def test_finds_the_typo_in_both_encodings():
    for enc in ENCODINGS:
        data = b"\x00\x01" + "Tabel has column heading only".encode(enc) + b"\x00"
        assert hits(data) == [enc]
    assert hits("Table has column headers only".encode("utf-16-le")) == []


def test_no_form_says_tabel(repo_root):
    bad = {f.name: hits(f.read_bytes()) for f in form_files(repo_root)}
    bad = {k: v for k, v in bad.items() if v}
    assert not bad, f'"{TYPO}" found in: {bad}'

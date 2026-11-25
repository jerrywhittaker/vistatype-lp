"""Importing a file into an LP book uses THE BOOK'S spacing and italics settings (Jerry, 11/24/2026).

Lp_Import_Exported_Selection_File runs Lp_Normalize_Styles on the hidden source file, and
Normalize and its italics step read the document they run on. So the book's three settings
must be put on the source file before Normalize, and the file's own marker read first, so that
italics it had already marked another way are converted to the book's marker.
A check on the source; the import reaches documents, which the headless VBA runner cannot.
"""
import re
from pathlib import Path


def body(repo_root):
    text = (Path(repo_root) / "src/vba/LpExportImportSelectedText.bas").read_text(encoding="utf-8", errors="replace")
    start = text.index("Sub Lp_Import_Exported_Selection_File()")
    return text[start:text.index("End Sub", start)]


def at(b, pattern):
    m = re.search(pattern, b)
    assert m, f"missing: {pattern}"
    return m.start()


def test_book_settings_go_on_the_file_before_normalize(repo_root):
    b = body(repo_root)
    src = at(b, r"srcMarker = Lp_Pref_Italics_Marker\(sourceDoc\)")
    book = at(b, r"bookMarker = Lp_Pref_Italics_Marker\(destDoc\)")
    put = at(b, r"Lp_Pref_Set_For_Doc sourceDoc, Lp_Pref_Line_Multiple\(destDoc\), _\s+Lp_Pref_Letter_Spacing\(destDoc\), bookMarker")
    conv = at(b, r"Lp_Italics_Convert_Marker\(srcMarker, bookMarker\)")
    norm = at(b, r"Call Lp_Normalize_Styles")
    assert src < put and book < put < conv < norm
    assert "sourceDoc.Close SaveChanges:=False" in b, "the file must not be saved with the book's settings"

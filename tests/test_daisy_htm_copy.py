"""DAISY/NIMAS conversion: the copy beside the book is an .htm that starts at <dtbook.

Jerry, 11/27/2026: Convert DAISY/NIMAS to Word Document (308) wrote a .txt beside the book's
.xml, the same as the .html - VistaType LP's own <html><head>...<body> wrapper and all. It is now
an .htm, written before that wrapper goes on, so nothing comes before <dtbook. Word can open it
by hand and gives the same book as the conversion, without the red on the $pg tags.

These are checks on the source: the order of the lines in the converter.
"""
from test_toc_never_deletes_a_line import module_text, procedures

CONVERTER = "Sh_Convert_XML_File_To_Word_Document"


def _lines(repo_root):
    return procedures(module_text(repo_root))[CONVERTER]


def _index(lines, text):
    hits = [i for i, l in enumerate(lines) if text in l]
    assert len(hits) == 1, f"expected one line holding {text!r}, found {len(hits)}"
    return hits[0]


def test_the_copy_is_an_htm_and_no_txt_is_written(repo_root):
    lines = _lines(repo_root)
    path = lines[_index(lines, "htmCopy = folderPath & ")]
    assert path.rstrip().endswith('& ".htm"')
    _index(lines, "outStream.SaveToFile htmCopy, 2")
    code = [l for l in lines if not l.lstrip().startswith("'")]
    assert not any('".txt"' in l or "txtCopy" in l for l in code)


def test_the_htm_is_written_before_the_wrapper_and_after_the_tagging(repo_root):
    lines = _lines(repo_root)
    prolog = _index(lines, "fileContent = Mid$(fileContent, dtPos)")
    pagenums = _index(lines, "fileContent = Sh_Tag_Empty_Pagenums(fileContent)")
    prodnotes = _index(lines, "fileContent = Sh_Tag_Prodnotes_As_Prodnote_Style(fileContent)")
    htm = _index(lines, "outStream.SaveToFile htmCopy, 2")
    wrapper = _index(lines, 'fileContent = "<html><head>')
    html = _index(lines, "outStream.SaveToFile htmlPath, 2")
    assert prolog < pagenums < prodnotes < htm < wrapper < html
    # nothing else is done to fileContent between the .htm and the wrapper
    assert not any(l.lstrip().startswith("fileContent =") for l in lines[htm + 1:wrapper])


def test_message_308_and_the_bar_say_htm(repo_root):
    lines = _lines(repo_root)
    assert any("An .htm file holding the book's XML code, which Word can open" in l for l in lines)
    assert any('Sh_Progress_Say 3, "Creating .htm file' in l for l in lines)
    assert not any(".txt format" in l for l in lines)

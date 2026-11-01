"""DAISY or NIMAS to Word needs no document open (Jerry, 11/1/2026).

The conversion opens its book from a folder, so message 293 ("Cannot continue... no document is
open.", from Sh_Is_Doc_Open) only got in the way. With none open, a blank document is added
first, because Step 1's ActiveWindow and the progress bar's home document need one.
"""
import re

from test_toc_box_stays_open import module_text, procedures


def test_the_converter_does_not_stop_with_293(repo_root):
    procs = procedures(module_text(repo_root))
    for name in ("DN_Menu_Starter", "Sh_Convert_XML_File_To_Word_Document"):
        code = [c.split("'")[0] for c in procs[name]]
        assert not any(re.search(r"\bSh_Is_Doc_Open\b", c) for c in code), (
            f"{name} calls Sh_Is_Doc_Open again, which stops with message 293 when no document is open")


def test_a_blank_document_comes_before_the_first_window_read(repo_root):
    code = [c.split("'")[0] for c in procedures(module_text(repo_root))["Sh_Convert_XML_File_To_Word_Document"]]
    add = next(n for n, c in enumerate(code) if re.search(r"Documents\.count\s*=\s*0\s+Then\s+Application\.Documents\.Add", c))
    window = next(n for n, c in enumerate(code) if "ActiveWindow" in c)
    assert add < window, "the blank document must be added before Step 1 reads ActiveWindow"

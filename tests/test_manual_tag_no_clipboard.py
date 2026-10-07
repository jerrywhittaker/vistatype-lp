"""Braille Manual Tag Ref Page reads the line without the clipboard.

Reported by Jerry 10/6/2026 on 3.0.539: Manual Tag Ref Page on a line "64A", in a book with the
Normal template, raised -2147221040, "DataObject:GetFromClipboard OpenClipboard Failed", twice
(VistaType-Errors.log on vistabuild). The sub copied the line with Selection.Copy and read it
back through an MSForms DataObject; when another program had the clipboard open, the read failed
(inferred - the same steps ran 20 times clean in a separate Word). The code was there from the
first commit; it became reachable on a Normal-template book when issue #8 took the template check
out of this button on 10/6/2026.

The line is now read straight from Selection.Text. Not unit-testable in VBA: the sub works on the
Selection, and anything that reaches a Document hangs the headless runner (tests/vba/README.md).
So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

SUB = "Dx_Manual_Tag_with_Dollar_pg"
CLIPBOARD = re.compile(r"\bDataObject\b|\bGetFromClipboard\b|\bSelection\.Copy\b", re.IGNORECASE)


def _code(repo_root):
    procs = procedures(module_text(repo_root))
    assert SUB in procs, f"{SUB} is missing from LPandBrlMacros.bas"
    return [strip_comment(c) for c in procs[SUB]]


def test_manual_tag_does_not_use_the_clipboard(repo_root):
    hits = [c.strip() for c in _code(repo_root) if CLIPBOARD.search(c)]
    assert not hits, (
        f"{SUB} reads the line through the clipboard again - it fails with -2147221040 when "
        f"another program has the clipboard open: {hits}")


def test_manual_tag_reads_the_selection_text(repo_root):
    assert any(re.search(r"Possible_LC_Roman\s*=.*\bSelection\.Text\b", c, re.IGNORECASE)
               for c in _code(repo_root)), (
        f"{SUB} must read the line from Selection.Text.")

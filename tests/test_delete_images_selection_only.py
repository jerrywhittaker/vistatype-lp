"""Remove Pictures stays inside the selection (issue #31).

Dx_Delete_Images runs from Braille Selection Clean Up ("Remove Pictures") and from Full File
Cleanup (Dx_Ffc). Until 10/11/2026 it worked on the whole book even with text selected: a For
Each over ActiveDocument.Shapes selected and converted every floating shape, the count was
ActiveDocument.InlineShapes, and both Finds were Selection.Find with wdFindContinue.

Measured on vistabuild, 10/11/2026, in a hidden Word with the old and new subs copied out of the
source: with paragraph 3 of 5 selected (an inline and a floating picture in paragraphs 1, 3 and
5), the old sub said "Found 5 images in the document" and deleted every picture in the book
except one floating picture its For Each had skipped; the new one said "Found 2 images in the
selection", deleted the two in paragraph 3 only, and put the selection back. With nothing
selected the new one found and deleted all 6.

Not unit-testable in VBA: the sub works on the Selection and the active document, which hang the
headless runner (tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

SUB = "Dx_Delete_Images"


def _code(repo_root):
    procs = procedures(module_text(repo_root))
    assert SUB in procs, f"{SUB} is missing from LPandBrlMacros.bas"
    return [c for c in procs[SUB] if c.strip()]


def _joined(repo_root):
    return "\n".join(_code(repo_root))


def test_scope_is_the_selection_when_there_is_one(repo_root):
    body = _joined(repo_root)
    assert re.search(r"Selection\.Type\s*=\s*wdSelectionNormal", body, re.IGNORECASE), (
        f"{SUB} must test for a selection, as the other Selection Clean Up passes do")
    assert re.search(r"Set\s+rng\s*=\s*Selection\.Range\b", body, re.IGNORECASE), (
        f"{SUB} must work on Selection.Range when text is selected")
    assert re.search(r"Set\s+rng\s*=\s*ActiveDocument\.Content\b", body, re.IGNORECASE), (
        f"{SUB} must keep the whole book when nothing is selected - Full File Cleanup relies on it")


def test_nothing_reaches_past_the_range(repo_root):
    body = _joined(repo_root)
    assert not re.search(r"\bwdFindContinue\b", body, re.IGNORECASE), (
        f"{SUB} uses wdFindContinue - a Replace All that wraps runs through the rest of the book")
    assert not re.search(r"\bSelection\.Find\b", body, re.IGNORECASE), (
        f"{SUB} uses Selection.Find - the Finds must run on the range (rng)")
    assert not re.search(r"ActiveDocument\.InlineShapes\b", body, re.IGNORECASE), (
        f"{SUB} counts ActiveDocument.InlineShapes - the whole book, not the selection")
    assert re.search(r"\brng\.InlineShapes\.count\b", body, re.IGNORECASE), (
        f"{SUB} must count rng.InlineShapes")
    wraps = re.findall(r"\.Wrap\s*=\s*(\w+)", body)
    assert wraps and all(w.lower() == "wdfindstop" for w in wraps), (
        f"{SUB}: every Find must use .Wrap = wdFindStop; found {wraps}")


def test_floating_shapes_converted_in_range_backwards_without_selecting(repo_root):
    body = _joined(repo_root)
    assert not re.search(r"For\s+Each\s+\w+\s+In\s+ActiveDocument\.Shapes", body, re.IGNORECASE), (
        f"{SUB} walks Shapes with For Each - converting takes a shape out of the collection, so "
        "For Each skips the next one (issue #27's fault)")
    assert re.search(r"For\s+i\s*=\s*ActiveDocument\.Shapes\.count\s+To\s+1\s+Step\s+-1", body,
                     re.IGNORECASE), f"{SUB} must convert floating shapes backwards by index"
    assert re.search(r"If\s+Sh_Shape_In_Range\(\s*shp\s*,\s*rng\s*\)\s+Then\s+shp\.ConvertToInlineShape",
                     body, re.IGNORECASE), (
        f"{SUB} must convert only shapes anchored in the range (Sh_Shape_In_Range), directly")
    assert not re.search(r"\.Select\b(?!ion)", body.replace("rng.Select", ""), re.IGNORECASE), (
        f"{SUB} selects something other than putting the selection back - selecting a shape "
        "moves the transcriber's selection")


def test_message_says_selection_and_selection_is_put_back(repo_root):
    body = _joined(repo_root)
    assert '"in the selection"' in body and '"in the document"' in body, (
        f"{SUB}'s count message must say \"in the selection\" when there is one")
    assert re.search(r"If\s+hadSelection\s+Then\s+rng\.Select\b", body, re.IGNORECASE), (
        f"{SUB} must put the selection back - the passes after it in Selection Clean Up need it")

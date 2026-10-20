"""Braille Compress Linear Math stays inside the selection.

Found by reading, 10/18/2026: the x-to-times-sign pass ran with wdFindContinue, so with a
selection Replace All went on through the whole book - every x became a times sign ("next"
became "ne×t"). Every pass in the sub must stop at the end of the selection. This test lived in
test_braille_cleanup_keeps_mathml.py until that file went with the MathML park, 10/20/2026.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures


def test_compress_linear_math_stays_in_the_selection(repo_root):
    body = procedures(module_text(repo_root))["Dx_Compress_Linear_Math"]
    assert not any("wdFindContinue" in l for l in body), "a Compress Linear Math pass still runs past the selection"
    x = [i for i, l in enumerate(body) if re.search(r'^\s*\.Text = "x"', l)]
    assert x, "the x-to-times-sign pass is not where it was"
    assert re.search(r"\.Wrap = wdFindStop", "\n".join(body[x[0]:x[0] + 4]))

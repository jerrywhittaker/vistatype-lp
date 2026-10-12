"""Fix Common File Errors makes every picture inline before anything else runs.

Jerry, 10/12/2026: attaching the large print template deleted pictures that were not inline,
and he narrowed it to Lp_Fix_Common_File_Errors. A floating picture hangs off an anchor
paragraph, and the passes delete and merge paragraphs. So the FIRST pass now converts every
floating picture to inline.

Issue #27, the same day: Lp_SetPicturesToInlineAndLockAspectRatio walked ActiveDocument.Shapes
with For Each, and converting a picture takes it out of that collection, so every other one was
stepped over (measured 10/9/2026: 3 floating, 1 left after one run). It walks backwards now.

A Document hangs the headless VBA suite (tests/vba/README.md), so this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import procedures
from test_no_end_while_box_open import module_text


def _code(body):
    return [c.split("'")[0] for c in body if c.split("'")[0].strip()]


def test_inline_is_the_first_pass(repo_root):
    body = _code(procedures(module_text(repo_root))["Lp_Fix_Common_File_Errors"])
    # The passes: the first line after the first step announcement. (Sh_Save_User_Position comes
    # before it and only records where the cursor is.)
    first_step = next(i for i, c in enumerate(body) if re.match(r"\s*Lp_Ffc_Step\s+stepNo", c))
    runs = body[first_step + 1:first_step + 2]
    assert runs and "Lp_SetPicturesToInlineAndLockAspectRatio" in runs[0], (
        "The first pass of Lp_Fix_Common_File_Errors is no longer the inline conversion - "
        f"a later pass could delete a floating picture with its paragraph. First: {runs[:1]}")


def test_every_floating_picture_is_converted(repo_root):
    body = "\n".join(_code(procedures(module_text(repo_root))["Lp_SetPicturesToInlineAndLockAspectRatio"]))
    assert not re.search(r"For\s+Each\s+\w+\s+In\s+ActiveDocument\.Shapes", body, re.I), (
        "For Each over Shapes while converting skips every other picture (issue #27)")
    assert re.search(r"For\s+i\s*=\s*ActiveDocument\.Shapes\.count\s+To\s+1\s+Step\s+-1", body, re.I)
    assert len(re.findall(r"ConvertToInlineShape", body)) == 1


def test_step_total_is_one_more_than_the_passes(repo_root):
    procs = procedures(module_text(repo_root))
    passes = sum(1 for c in _code(procs["Lp_Fix_Common_File_Errors"])
                 if re.match(r"\s*Lp_Ffc_Step\s+stepNo", c))
    total = int(re.search(r"Const LP_FFC_STEPS As Long = (\d+)",
                          "\n".join(procs["Lp_Ffc_Step"])).group(1))
    assert total == passes + 1, (
        f"LP_FFC_STEPS is {total} for {passes} passes - it must be one more, or the bar reads "
        "100% while the last pass is still running")

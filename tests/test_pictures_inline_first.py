"""Pictures on attach: inline on a first attach only, and before the cleanup.

Jerry, 10/12/2026: attaching the large print template deleted pictures that were not inline,
and he narrowed it to Lp_Fix_Common_File_Errors - a floating picture hangs off an anchor
paragraph, and the cleanup deletes and merges paragraphs. His rules, the same day:

  * a FIRST attach (a book that is not large print yet) makes every picture inline, and it
    must happen BEFORE the cleanup;
  * a RE-ATTACH, and Fix Common File Errors run on its own, leave floating pictures as they are.

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


INLINE = "Lp_SetPicturesToInlineAndLockAspectRatio"


def test_cleanup_on_its_own_leaves_floating_pictures(repo_root):
    body = "\n".join(_code(procedures(module_text(repo_root))["Lp_Fix_Common_File_Errors"]))
    assert INLINE not in body, (
        "Lp_Fix_Common_File_Errors makes pictures inline - run on its own it must leave floating "
        "pictures as they are (Jerry, 10/12/2026)")


def test_first_attach_only_and_before_the_cleanup(repo_root):
    body = _code(procedures(module_text(repo_root))["Lp_Attach_The_Template"])
    calls = [i for i, c in enumerate(body) if re.match(rf"\s*(Call\s+)?{INLINE}\b", c)]
    assert len(calls) == 1, (
        f"Lp_Attach_The_Template should make pictures inline exactly once, found {len(calls)} - "
        "a re-attach must leave floating pictures as they are")
    i = calls[0]
    assert re.match(r"\s*Lp_Fix_Common_File_Errors\s*$", body[i + 1]), (
        "The inline conversion must come straight before the cleanup, which deletes floating pictures")
    opener = max(j for j in range(i) if re.match(r"\s*If\b.*\bThen\s*$", body[j]))
    assert re.match(r"\s*If\s+Not\s+Lp_Doc_Was_Already_LP\s+Then\s*$", body[opener]), (
        "The inline conversion is no longer inside 'If Not Lp_Doc_Was_Already_LP' - a re-attach "
        "would make floating pictures inline")


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

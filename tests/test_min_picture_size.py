"""No picture smaller than the text: the LP template attach enlarges tiny pictures.

Jerry, 11/1/2026: "I've encountered some pictures that are just a dot on the screen even in a file
which is already large print." Any picture whose shorter side is under the Normal style's font
size is enlarged to it, keeping its shape, then fitted inside the margins. Attach only, no button
(Jerry). It must run after the text is at its point size, which is final only after
Lp_Normalize_Styles and Lp_Restore_Style_Sizes.

The arithmetic is tested in tests/vba/TestMinPictureSize.bas. Everything else reaches a picture,
which hangs the headless runner (tests/vba/README.md), so this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures


def _code(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    # drop comments so a mention in a comment does not count
    return "\n".join(c.split("'")[0] for c in procs[name] if c.split("'")[0].strip())


def test_the_attach_enlarges_after_the_sizes_are_final(repo_root):
    body = _code(repo_root, "Lp_Attach_The_Template")
    restore = body.find("Lp_Restore_Style_Sizes")
    normalize = body.find("Lp_Normalize_Styles")
    enlarge = body.find("Lp_Enlarge_Tiny_Pictures")
    assert enlarge >= 0, "the attach no longer enlarges pictures smaller than the text"
    assert 0 <= normalize < enlarge and 0 <= restore < enlarge, (
        "Lp_Enlarge_Tiny_Pictures must run after Lp_Normalize_Styles and Lp_Restore_Style_Sizes - "
        "before them the Normal style's size is not the book's point size yet")


def test_pictures_only_from_the_normal_size_and_inside_the_margins(repo_root):
    body = _code(repo_root, "Lp_Enlarge_Tiny_Pictures")
    assert "Lp_Same_Pic_Is_Picture" in body, (
        "only pictures may be enlarged - not MathType objects, charts or other OLE objects")
    assert re.search(r"Styles\s*\(\s*wdStyleNormal\s*\)\.Font\.Size", body), (
        "the minimum is the Normal style's size, read from the style")
    assert "Lp_Base_Font_Size" not in body, "Lp_Base_Font_Size can be stale; read the style"
    assert "Lp_Enlarge_One_Picture" in body
    assert not re.search(r"\.Shapes\b", body), "floating pictures are left alone"


def test_the_final_size_is_worked_out_before_the_picture_is_touched(repo_root):
    plan = _code(repo_root, "Lp_Enlarge_Plan")
    assert "Lp_Min_Pic_Factor" in plan and "Lp_Fit_Factor" in plan, (
        "the margin fit must be worked into the size BEFORE Width is set - enlarging first asked "
        "Word for sizes past its largest and left a rule line as a thick, stretched bar")
    assert not re.search(r"\.(Width|Height|LockAspectRatio)\s*=", plan), "the plan only reads"
    one = _code(repo_root, "Lp_Enlarge_One_Picture")
    assert "Lp_Fit_Inline_Shape_To_Margins" not in one
    assert "On Error GoTo" in one, "one picture Word will not resize must not end the attach"


def test_everything_is_read_before_anything_is_resized(repo_root):
    """Jerry 11/2/2026: 742 s on a math book. Reading a picture's margins straight after the last
    one was resized laid the book out again each time (107 pictures x 6.9 s a layout). Part 1
    reads and lists, Part 2 only resizes, and the margins are asked once when sections agree."""
    body = _code(repo_root, "Lp_Enlarge_Tiny_Pictures")
    plan = body.find("Lp_Enlarge_Plan")
    resize = body.find("Lp_Enlarge_One_Picture")
    assert 0 <= plan < resize, "every picture must be read before any is resized"
    loop2 = body[resize - 200:resize]
    assert "For k = 1 To n" in loop2, "the resizing must be a second loop over the list"
    assert "Lp_Same_Room_All_Sections" in body, "the margins must be asked once, not per picture"
    assert "Lp_Pic_Room" not in body

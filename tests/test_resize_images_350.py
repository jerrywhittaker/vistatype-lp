"""Resize Selected Pictures (350): a selection is required, pictures only, inside the margins.

Jerry, 10/14/2026: "require a range selection. resize only inline pictures. final picture sizes
must not exceed any margins... if so reduce the size to fit within the margins." One clicked
picture still counts (Jerry, the same day).

Before: an "All Pictures" choice scaled ActiveDocument.InlineShapes; a bare cursor in a table
resized the whole table; OLE objects, equations and charts were scaled with the pictures; nothing
stopped a picture growing past the margins; and Okay with a percent out of range showed message
300 and then ran the resize anyway with the last percent used.

The arithmetic of the fit is tested in tests/vba/TestResizePictures.bas. Everything else here
reaches the Selection or a picture, which hangs the headless runner (tests/vba/README.md), so
this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures


def _body(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return "\n".join(c for c in procs[name] if c.strip())


def _form_code(repo_root, form):
    text = (repo_root / "src" / "forms" / f"{form}.frm").read_text(encoding="latin-1")
    # drop comments so a mention in a comment does not count
    return "\n".join(line.split("'")[0] for line in text.splitlines())


def test_the_all_pictures_route_is_gone(repo_root):
    body = _body(repo_root, "Lp_Resize_Images")
    assert not re.search(r"ActiveDocument\.InlineShapes", body, re.I), (
        "Lp_Resize_Images reaches ActiveDocument.InlineShapes again - the whole book without a "
        "selection, which Jerry removed on 10/14/2026")
    assert "Lp_Pic_All_Selectd" not in module_text(repo_root)
    assert "Lp_Resize_Images_Has_Pictures" in body and "Lp_Resize_Images_Pics" in body, (
        "Lp_Resize_Images must ask the same rule as the menu button")


def test_a_bare_cursor_is_refused(repo_root):
    body = _body(repo_root, "Lp_Resize_Images_Pics")
    for allowed in ("wdSelectionInlineShape", "wdSelectionNormal", "wdSelectionRow"):
        assert allowed in body, f"Lp_Resize_Images_Pics no longer takes {allowed}"
    assert "wdSelectionIP" not in body, "a bare cursor must not count as a selection"
    assert not re.search(r"Information\s*\(\s*wdWithInTable", body, re.I), (
        "a cursor in a table must not mean the whole table any more - it is not a selection")


def test_only_pictures_and_each_fitted_after_scaling(repo_root):
    body = _body(repo_root, "Lp_Scale_Inline_Shapes")
    pic = body.find("Lp_Same_Pic_Is_Picture")
    scale = body.find(".ScaleWidth")
    fit = body.find("Lp_Fit_Inline_Shape_To_Margins")
    assert pic != -1, "OLE objects, equations and charts must be left alone"
    assert scale != -1 and fit != -1, "every picture must be scaled and then fitted"
    assert pic < scale < fit, "the picture test, then the scaling, then the fit - in that order"


def test_fit_uses_the_pictures_own_section_and_sets_points(repo_root):
    body = _body(repo_root, "Lp_Fit_Inline_Shape_To_Margins")
    assert re.search(r"\.Sections\(1\)\.PageSetup", body), (
        "margins must come from the picture's own section, not ActiveDocument.PageSetup")
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin", "Gutter"):
        assert m in body, f"the fit ignores {m}"
    assert "Lp_Fit_Factor" in body
    assert not re.search(r"\.Scale(Width|Height)\s*=", body), (
        "set Width and Height in points: Word keeps the scale as a whole percent, which could "
        "round a picture back past the margin")


def test_menu_button_checks_before_the_percent_box(repo_root):
    code = _form_code(repo_root, "Lp_Bakgrnd_Picture_Menu_Form")
    m = re.search(r"Sub ImageSizeButton_Click\(\)(.*?)End Sub", code, re.S)
    assert m, "ImageSizeButton_Click is missing"
    body = m.group(1)
    check = body.find("Lp_Resize_Images_Has_Pictures")
    show = body.find("Lp_Resize_Images_Form.Show")
    assert check != -1 and show != -1 and check < show, (
        "the selection must be checked before the percent box (350) opens")


def test_out_of_range_percent_stops(repo_root):
    code = _form_code(repo_root, "Lp_Resize_Images_Form")
    m = re.search(r"Sub CmdOkay_Click\(\)(.*?)End Sub", code, re.S)
    assert m, "CmdOkay_Click is missing"
    body = m.group(1)
    assert re.search(r"\(300\)\"\s*\n\s*Exit Sub", body), (
        "after message 300 Okay must stop, not go on and resize with the last percent")
    assert "Resize_All" not in code and "Resize_Selected" not in code


def test_fit_reads_both_sizes_before_setting_either(repo_root):
    """Found by review, 10/14/2026: with the aspect ratio locked, setting Width moved Height too,
    and Height was then read again and multiplied a second time - the picture shrank twice."""
    body = _body(repo_root, "Lp_Fit_Inline_Shape_To_Margins")
    assert not re.search(r"s\.(Width|Height)\s*\*\s*f", body), (
        "the new size must come from sizes read BEFORE either is set, not from s.Width/s.Height")
    read = body.find("h = s.Height")
    first_set = body.find("s.Width = ")
    assert read != -1 and first_set != -1 and read < first_set


def test_a_clicked_floating_picture_gets_397_not_142(repo_root):
    body = _body(repo_root, "Lp_Resize_Images_Pics")
    assert "wdSelectionShape" in body, (
        "a clicked floating picture must hand back an empty collection, so the user is told to "
        "use All Pictures to Inline (397) rather than to click a picture (142)")


def test_right_edge_comes_from_the_end_of_the_picture(repo_root):
    """Jerry, 10/14/2026: an indented or second-column picture is kept inside the right margin.
    Measured on the build box the same day: only the collapsed END of the picture's range gives
    its right edge correctly in every layout; the start is wrong for a centered picture, the whole
    range for an indented one, and the text-boundary position is measured from a table cell."""
    body = _body(repo_root, "Lp_Fit_Inline_Shape_To_Margins")
    assert re.search(r"\.Collapse\s+wdCollapseEnd", body), "measure at the END of the picture"
    assert "wdHorizontalPositionRelativeToPage" in body
    assert "wdHorizontalPositionRelativeToTextBoundary" not in body, (
        "the text-boundary position is measured from the cell in a table")
    assert "Lp_Fit_Right_Room" in body

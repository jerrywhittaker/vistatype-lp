"""Normalize Styles sets line spacing to At least, type size + 4 points (Jerry, 11/23/2026).

The Normal style only: it gets its own size + 4. Heading 1-5, TOC Heading, TOC 1-5, List Bullet
1-5, List 1-5, List Number 1-5, List Paragraph and the Para color styles would inherit that, so
they are set to Single, with 4 pt after instead of the base size - 0 for Heading 1-5 (Jerry,
11/23/2026).
At least, not Exactly: Exactly was tried first and cut off anything taller than the line -
pictures, equations, and the superscripts the attach enlarges afterward.

What this checks, on the source (Normalize Styles reaches a Document, which the headless VBA
runner cannot do - tests/vba/README.md):

  1. Normal is set At least at its own Font.Size + 4.
  2. Those styles are set to Single.
  3. Nothing in the pass uses Exactly.
  4. The pass comes after the size passes (so the sizes read are final) and before
     Sh_Color_Dollar_PG_Red, which must stay last.
  5. The progress bar's total is one more than the number of steps.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

SUB = "Lp_Normalize_Styles"

# Single line spacing and 4 pt after (Jerry, 11/23/2026).
ALL_SINGLE = (
    {f"Heading {n}" for n in range(1, 6)} | {"TOC Heading"} | {f"TOC {n}" for n in range(1, 6)}
    | {"List", "List Bullet", "List Number", "List Paragraph"}
    | {f"List {n}" for n in range(2, 6)} | {f"List Bullet {n}" for n in range(2, 6)}
    | {f"List Number {n}" for n in range(2, 6)}
    | {f"Para {c}" for c in ("Aqua", "Black", "Black Inverted", "Blue", "Green", "Tan", "Yellow")}
)
STEP = "Lp_Ns_Step"


def lines_of(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return procs[name]


def first_line(lines, pattern):
    return next((i for i, c in enumerate(lines) if re.search(pattern, c)), None)


def test_normal_is_at_least_its_size_plus_four(repo_root):
    body = "\n".join(lines_of(repo_root, SUB))
    m = re.search(r"With ActiveDocument\.Styles\(wdStyleNormal\)(.*?)End With", body, re.S)
    assert m, "Normal's line spacing is not set"
    assert re.search(r"\.ParagraphFormat\.LineSpacingRule = wdLineSpaceAtLeast", m.group(1))
    assert re.search(r"\.ParagraphFormat\.LineSpacing = \.Font\.Size \+ 4", m.group(1))


def test_normal_only_the_named_styles_stay_single(repo_root):
    """Jerry, 11/23/2026: points are added to the Normal style only. The headings, lists, TOC and
    Para color styles are based on Normal, so they are set to Single rather than inheriting it."""
    lines = lines_of(repo_root, SUB)
    start = first_line(lines, r'"Setting line spacing"')
    end = next(i for i in range(start, len(lines)) if "Sh_Color_Dollar_PG_Red" in lines[i])
    body = "\n".join(lines[start:end])
    assert body.count("wdLineSpaceAtLeast") == 1, "only Normal gets At least"
    names = re.search(r"StyleNames = Array\((.*?)\)\s*\n", body, re.S)
    assert names, "no list of heading styles kept Single"
    want = ALL_SINGLE
    got = set(re.findall(r'"([^"]+)"', names.group(1)))
    assert want <= got, f"not kept Single: {sorted(want - got)}"
    assert re.search(r"Sh_Style_Exists\(ActiveDocument, CStr\(StyleNames\(i\)\)\)", body)
    assert re.search(r"\.ParagraphFormat\.LineSpacingRule = wdLineSpaceSingle", body)
    assert not re.search(r"LineSpacing = .*\+", body.split("End With", 1)[1]), "no points added to headings"


def test_nothing_in_the_pass_is_exactly(repo_root):
    """Exactly (tried first, 11/23/2026) cut off pictures, equations and enlarged superscripts."""
    lines = lines_of(repo_root, SUB)
    start = first_line(lines, r'"Setting line spacing"')
    body = "\n".join(lines[start:])
    assert "wdLineSpaceExactly" not in body


def test_space_after_on_the_named_styles(repo_root):
    """Jerry, 11/23/2026: the lists, TOC and Para colors get 4 pt after, Heading 1-5 get 0. The earlier pass puts the base size after every paragraph as DIRECT formatting, so the
    style alone would never show - their paragraphs must be given it too, after that pass."""
    lines = lines_of(repo_root, SUB)
    story = first_line(lines, r"ActiveDocument\.Content\.ParagraphFormat\.SpaceAfter = Lp_Base_Font_Size")
    style_set = first_line(lines, r"sty\.ParagraphFormat\.SpaceAfter = spaceAfterPts")
    direct = first_line(lines, r"\.Replacement\.ParagraphFormat\.SpaceAfter = spaceAfterPts")
    assert story is not None and style_set is not None and direct is not None
    assert story < style_set < direct, "the 4 pt must come after the story-wide space after"
    lists = re.findall(r"StyleNames = Array\((.*?)\)\s*\n", "\n".join(lines[style_set - 30:style_set]), re.S)
    got = set(re.findall(r'"([^"]+)"', lists[-1]))
    want = ALL_SINGLE
    assert want == got, f"4 pt after: expected {sorted(want)}, got {sorted(got)}"
    pick = "\n".join(lines[style_set - 12:style_set])
    assert re.search(r'Left\$\(StyleNames\(i\), 8\) = "Heading "\s*Then\s*spaceAfterPts = 0\s*Else\s*spaceAfterPts = 4', pick), \
        "Heading 1-5 get 0 after, the rest 4"
    body = "\n".join(lines[style_set:direct + 20])
    assert ".Style = sty" in body and "Replace:=wdReplaceAll" in body and ".Format = True" in body
    # Auto space after shows as space even when SpaceAfter reads 0 (measured 11/23/2026).
    assert "sty.ParagraphFormat.SpaceAfterAuto = False" in "\n".join(lines[style_set - 3:style_set])
    # A Find and Replace cannot clear Auto (measured), so the paragraphs get it directly, story-wide,
    # before the story-wide space after.
    auto = first_line(lines, r"ActiveDocument\.Content\.ParagraphFormat\.SpaceAfterAuto = False")
    assert auto is not None and auto < story, "Auto space after must be cleared on every paragraph first"


def test_the_pass_sits_after_the_sizes_and_before_the_red(repo_root):
    lines = lines_of(repo_root, SUB)
    sizes = first_line(lines, r'"Ensuring base font size for Normal and color paragraph styles"')
    spacing = first_line(lines, r'"Setting line spacing"')
    red = first_line(lines, r'"Sh_Color_Dollar_PG_Red"')
    assert sizes is not None and spacing is not None and red is not None
    assert sizes < spacing < red


def test_the_bar_total_is_one_more_than_the_steps(repo_root):
    steps = sum(1 for c in lines_of(repo_root, SUB) if re.match(rf"\s*{STEP} stepNo,", c))
    body = "\n".join(lines_of(repo_root, STEP))
    total = int(re.search(r"Const LP_NS_STEPS As Long = (\d+)", body).group(1))
    assert total == steps + 1, f"{steps} steps, so LP_NS_STEPS should be {steps + 1}, not {total}"

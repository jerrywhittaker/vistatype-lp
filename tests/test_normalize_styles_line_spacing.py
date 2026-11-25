"""Normalize Styles sets line spacing to Multiple 1.25 (Jerry, 11/24/2026; At least size + 4 from 11/23).

The Normal style only: it gets Multiple 1.25 (LP_LINE_MULTIPLE). Heading 1-5, TOC Heading, TOC 1-5, List Bullet
1-5, List 1-5, List Number 1-5, List Paragraph and the Para color styles would inherit that, so
Heading 1-5 are set to Single with 0 after (Jerry, 11/23/2026), and the others take Normal's
Normal's Multiple 1.25 with 0 after (Jerry, 11/24/2026: one rule, one number to offer the user).
At least, not Exactly: Exactly was tried first and cut off anything taller than the line -
pictures, equations, and the superscripts the attach enlarges afterward.

What this checks, on the source (Normalize Styles reaches a Document, which the headless VBA
runner cannot do - tests/vba/README.md):

  1. Normal is set to Multiple nsLine - the book's own line spacing, from Lp_Pref_Line_Multiple
     (Spacing and Italics Settings, 11/24/2026); LP_LINE_MULTIPLE is the built-in value.
  2. Heading 1-5 Single; the others Multiple 1.25, on the style and on the paragraphs.
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


def test_normal_is_multiple_1_25(repo_root):
    body = "\n".join(lines_of(repo_root, SUB))
    m = re.search(r"With ActiveDocument\.Styles\(wdStyleNormal\)(.*?)End With", body, re.S)
    assert m, "Normal's line spacing is not set"
    assert re.search(r"\.ParagraphFormat\.LineSpacingRule = wdLineSpaceMultiple", m.group(1))
    assert re.search(r"\.ParagraphFormat\.LineSpacing = LinesToPoints\(nsLine\)", m.group(1))
    # The book's own line spacing (Spacing and Italics Settings, Jerry, 11/24/2026), 1.25 built in.
    assert re.search(r"^\s*nsLine = Lp_Pref_Line_Multiple\(ActiveDocument\)\s*$", body, re.M)


def test_headings_single_the_rest_multiple(repo_root):
    """Jerry, 11/24/2026: TOC, List and Para color styles take Normal's Multiple 1.25;
    Heading 1-5 stay Single. No points are added to the headings."""
    lines = lines_of(repo_root, SUB)
    start = first_line(lines, r'"Setting line spacing"')
    end = next(i for i in range(start, len(lines)) if "Sh_Color_Dollar_PG_Red" in lines[i])
    body = "\n".join(lines[start:end])
    names = re.search(r"StyleNames = Array\((.*?)\)\s*\n", body, re.S)
    assert names, "no list of styles"
    got = set(re.findall(r'"([^"]+)"', names.group(1)))
    assert ALL_SINGLE <= got, f"missing: {sorted(ALL_SINGLE - got)}"
    assert re.search(r"Sh_Style_Exists\(ActiveDocument, CStr\(StyleNames\(i\)\)\)", body)
    assert re.search(r"bodyLinePts = LinesToPoints\(nsLine\)", body)
    assert re.search(r'isHeadingStyle = \(Left\$\(StyleNames\(i\), 8\) = "Heading "\)', body)
    assert re.search(r"If isHeadingStyle Then\s+sty\.ParagraphFormat\.LineSpacingRule = wdLineSpaceSingle\s+"
                     r"Else\s+sty\.ParagraphFormat\.LineSpacingRule = wdLineSpaceMultiple\s+"
                     r"sty\.ParagraphFormat\.LineSpacing = bodyLinePts", body)
    # Paragraphs carrying direct line spacing get it too, through the same Find and Replace.
    assert re.search(r"If Not isHeadingStyle Then\s+(?:'.*\s+)*"
                     r"\.LineSpacingRule = wdLineSpaceMultiple\s+"
                     r"\.LineSpacing = bodyLinePts", body)


def test_nothing_in_the_pass_is_exactly(repo_root):
    """Exactly (tried first, 11/23/2026) cut off pictures, equations and enlarged superscripts."""
    lines = lines_of(repo_root, SUB)
    start = first_line(lines, r'"Setting line spacing"')
    body = "\n".join(lines[start:])
    assert "wdLineSpaceExactly" not in body


def test_space_after_on_the_named_styles(repo_root):
    """Jerry, 11/24/2026: all of these styles get 0 pt after (the lists, TOC and Para colors had 4). The earlier pass puts the base size after every paragraph as DIRECT formatting, so the
    style alone would never show - their paragraphs must be given it too, after that pass."""
    lines = lines_of(repo_root, SUB)
    story = first_line(lines, r"ActiveDocument\.Content\.ParagraphFormat\.SpaceAfter = Lp_Base_Font_Size")
    style_set = first_line(lines, r"sty\.ParagraphFormat\.SpaceAfter = spaceAfterPts")
    direct = first_line(lines, r"^\s*\.SpaceAfter = spaceAfterPts")
    assert story is not None and style_set is not None and direct is not None
    assert story < style_set < direct, "the 0 pt must come after the story-wide space after"
    lists = re.findall(r"StyleNames = Array\((.*?)\)\s*\n", "\n".join(lines[style_set - 45:style_set]), re.S)
    got = set(re.findall(r'"([^"]+)"', lists[-1]))
    want = ALL_SINGLE
    assert want == got, f"0 pt after: expected {sorted(want)}, got {sorted(got)}"
    pick = "\n".join(lines[style_set - 14:style_set])
    assert re.search(r"spaceAfterPts = 0\s*$", pick, re.M), "every one of these styles gets 0 after (11/24/2026)"
    assert not re.search(r"spaceAfterPts = [1-9]", pick), "no 4 pt after any more"
    body = "\n".join(lines[style_set:direct + 20])
    assert ".Style = sty" in body and ".Format = True" in body
    assert re.search(r"Do While spRng\.Find\.Execute", body), "each run of the style is found in turn"
    assert "With spRng.ParagraphFormat" in body
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


def test_the_multiple_is_aph_built_in_and_wcag_on_its_button(repo_root):
    """Jerry, 11/25/2026: built in is VTLP V2 (APH)'s Multiple 1.00; W3C WCAG's 1.25 (1.5 times the
    type size in Tahoma, measured 18-28 pt, 11/24/2026) is the W3C WCAG button's value."""
    text = module_text(repo_root)
    assert re.search(r"^Private Const LP_LINE_MULTIPLE As Single = 1(\.0+)?#?\s*$", text, re.M)
    assert re.search(r"^Private Const LP_WCAG_LINE_MULTIPLE As Single = 1\.25\b", text, re.M)


def test_no_replace_all_in_normalize_styles(repo_root):
    """11/24/2026: the line-spacing pass sets each found run of a style directly instead of a
    Find and Replace with replacement paragraph formatting. Tried while chasing issue #56 and
    NOT a cure for it; this only holds the pass to the way it is written now."""
    body = "\n".join(lines_of(repo_root, SUB))
    assert "wdReplaceAll" not in body, "no replace-all in Normalize Styles"
    assert ".Replacement.ParagraphFormat" not in body, "no replacement paragraph formatting"

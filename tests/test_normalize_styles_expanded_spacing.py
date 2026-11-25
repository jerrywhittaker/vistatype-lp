"""Normalize Styles' expanded character spacing comes out exactly as its tables say (11/23/2026).

MEASURED on vistabuild at every size from 14 to 42: Word keeps Font.Spacing in twentieths of a
point and drops any fraction of one, and Font.Spacing is a Single - so 1.05, stored as 1.0499999,
came out 1.00. 14 of the 15 body values and 8 of the 15 heading values were 0.05 point under the
table. Lp_Spacing_For_Word rounds to the nearest twentieth and adds a quarter of one.

Jerry, 11/24/2026: the tables are gone. Letter spacing is a share of the type size
(LP_LETTER_SPACING: 0.12, WCAG 1.4.12, until 11/25/2026; now 0.07, VTLP V2 (APH)) - for body text at the base size, and for each heading at its
own size (base + 10 down to base + 2).

What this checks, on the source (it reaches a Document, which the headless VBA runner cannot do):
  1. Both spacing assignments go through Lp_Spacing_For_Word.
  2. The helper's formula, run here with Single precision, lands in the intended twentieth for
     every value 0.00 to 0.25 x size gives from 14 to 52 pt - whether Word drops the fraction or rounds it.
  3. The built-in constant is 0.07 (W3C WCAG's 0.12 is its own constant), and both assignments use the book's own share (nsLetter, from
     Lp_Pref_Letter_Spacing, which falls back to it - 11/24/2026), headings on their own size.
"""
import math
import re
import struct

from test_toc_never_deletes_a_line import module_text, procedures

SUB = "Lp_Normalize_Styles"
HELPER = "Lp_Spacing_For_Word"


def single(x):
    return struct.unpack("f", struct.pack("f", x))[0]


def procs(repo_root):
    p = procedures(module_text(repo_root))
    for name in (SUB, HELPER):
        assert name in p, f"{name} is missing from LPandBrlMacros.bas"
    return p


def sizes():
    # every base 14-42 and every heading size above it (base + 2 .. base + 10)
    return sorted({float(s) for b in range(14, 43, 2) for s in [b] + [b + o for o in (2, 4, 6, 8, 10)]})


def test_both_assignments_use_the_helper(repo_root):
    body = "\n".join(procs(repo_root)[SUB])
    assigns = re.findall(r"sty\.Font\.spacing = (.+)", body, re.I)
    assert assigns == ["Lp_Spacing_For_Word(targetSpacing)",
                       "Lp_Spacing_For_Word(nsLetter * (base + SizeOffsets(i)))"]
    assert re.search(r"targetSpacing = nsLetter \* CDbl\(Lp_Base_Font_Size\)", body)
    # The book's own letter spacing (Spacing and Italics Settings, Jerry, 11/24/2026), 0.12 built in.
    assert re.search(r"^\s*nsLetter = Lp_Pref_Letter_Spacing\(ActiveDocument\)\s*$", body, re.M)
    assert "SpacingTable" not in body and "headingSpacing" not in body, "the old tables are gone"


def test_letter_spacing_is_aph_built_in_and_wcag_on_its_button(repo_root):
    """Jerry, 11/25/2026: built in is VTLP V2 (APH)'s 0.07 times the type size; W3C WCAG's 0.12
    (11/24/2026) is the W3C WCAG button's value."""
    text = module_text(repo_root)
    assert re.search(r"^Private Const LP_LETTER_SPACING As Single = 0\.07\b", text, re.M)
    assert re.search(r"^Private Const LP_WCAG_LETTER_SPACING As Single = 0\.12\b", text, re.M)


def test_the_helper_lands_in_the_right_twentieth(repo_root):
    helper = "\n".join(procs(repo_root)[HELPER])
    assert re.search(r"CSng\(\(Round\(pts \* 20#, 0\) \+ 0\.25\) / 20#\)", helper)
    # every share the box offers, 0.00 to 0.25 (11/25/2026: the built-in is 0.07, WCAG's 0.12)
    for size, share in ((z, n / 100) for z in sizes() for n in range(0, 26)):
        value = single(single(share) * size)
        want = round(value * 20)
        sent = single((round(value * 20) + 0.25) / 20) * 20
        assert math.floor(sent) == want, f"{value}: Word would keep {math.floor(sent)}/20"
        assert round(sent) == want, f"{value}: Word would round to {round(sent)}/20"
        # and the old way really did lose a twentieth for values like this one
    assert math.floor(single(1.05) * 20) == 20, "the measured fault: 1.05 as a Single drops to 1.00"

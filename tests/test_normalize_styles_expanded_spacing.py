"""Normalize Styles' expanded character spacing comes out exactly as its tables say (11/23/2026).

MEASURED on vistabuild at every size from 14 to 42: Word keeps Font.Spacing in twentieths of a
point and drops any fraction of one, and Font.Spacing is a Single - so 1.05, stored as 1.0499999,
came out 1.00. 14 of the 15 body values and 8 of the 15 heading values were 0.05 point under the
table. Lp_Spacing_For_Word rounds to the nearest twentieth and adds a quarter of one.

Also Jerry, 11/23/2026: the body table's 38 pt value is 3.8 (the 0.25 step), not 3.65.

What this checks, on the source (it reaches a Document, which the headless VBA runner cannot do):
  1. Both spacing assignments go through Lp_Spacing_For_Word.
  2. The helper's formula, run here with Single precision, lands in the intended twentieth for
     every table value - whether Word drops the fraction or rounds it.
  3. The body table steps by exactly 0.25 per size from 14 to 42, so 38 is 3.8.
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


def body_table(lines):
    text = "\n".join(lines)
    m = re.search(r"SpacingTable = Array\((.*?)\)\s*\n", text, re.S)
    nums = [float(n) for n in re.findall(r"-?\d+(?:\.\d+)?", m.group(1).replace("_", " "))]
    return dict(zip(nums[0::2], nums[1::2]))


def heading_table(lines):
    text = "\n".join(lines)
    return {float(a): float(b) for a, b in re.findall(r"Case (\d+): headingSpacing = ([\d.]+)", text)}


def test_both_assignments_use_the_helper(repo_root):
    body = "\n".join(procs(repo_root)[SUB])
    assigns = re.findall(r"sty\.Font\.spacing = (.+)", body, re.I)
    assert assigns == ["Lp_Spacing_For_Word(targetSpacing)", "Lp_Spacing_For_Word(headingSpacing)"]
    assert re.search(r"Dim targetSpacing As Double", body)
    assert re.search(r"Dim headingSpacing As Double", body)


def test_the_helper_lands_in_the_right_twentieth(repo_root):
    helper = "\n".join(procs(repo_root)[HELPER])
    assert re.search(r"CSng\(\(Round\(pts \* 20#, 0\) \+ 0\.25\) / 20#\)", helper)
    lines = procs(repo_root)[SUB]
    for value in list(body_table(lines).values()) + list(heading_table(lines).values()):
        want = round(value * 20)
        sent = single((round(value * 20) + 0.25) / 20) * 20
        assert math.floor(sent) == want, f"{value}: Word would keep {math.floor(sent)}/20"
        assert round(sent) == want, f"{value}: Word would round to {round(sent)}/20"
        # and the old way really did lose a twentieth for values like this one
    assert math.floor(single(1.05) * 20) == 20, "the measured fault: 1.05 as a Single drops to 1.00"


def test_body_table_steps_by_a_quarter_point(repo_root):
    table = body_table(procs(repo_root)[SUB])
    assert sorted(table) == [float(n) for n in range(14, 43, 2)]
    assert table[38.0] == 3.8, "Jerry, 11/23/2026: 38 pt is 3.8, not 3.65"
    sizes = sorted(table)
    for a, b in zip(sizes, sizes[1:]):
        assert abs(table[b] - table[a] - 0.25) < 1e-9, f"{a} to {b} steps by {table[b] - table[a]:.2f}"

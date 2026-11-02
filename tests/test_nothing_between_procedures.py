"""Only comments and blank lines may stand between one procedure and the next.

vba-review, 11/1/2026: a scripted edit on 9/4/2026 (9705ff3) cut the first line of a comment so its
tail - "the whole of what Change Picture Color does." - stuck to an End Function line, harmless as
part of that line's comment. An insert on 11/1/2026 split the line and left the tail on its own
after End Sub, where VBA raises "Only comments may appear after End Sub". Nothing caught it: VBA
compiles each procedure only when it is first run (no build-time compile is possible).
"""
import re

import pytest

END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
START = re.compile(r"^\s*((Public|Private|Friend)\s+)?(Static\s+)?(Sub|Function|Property)\s", re.IGNORECASE)


def _modules(repo_root):
    src = repo_root / "src"
    return sorted(list((src / "vba").glob("*.bas")) + list((src / "vba").glob("*.cls"))
                  + list((src / "forms").glob("*.frm")))


@pytest.mark.parametrize("name", ["all"])
def test_only_comments_between_procedures(repo_root, name):
    bad = []
    for path in _modules(repo_root):
        between = False
        for n, line in enumerate(path.read_bytes().decode("latin-1").splitlines(), 1):
            if END.match(line):
                between = True
                continue
            if START.match(line):
                between = False
                continue
            if between and line.strip() and not line.lstrip().startswith("'"):
                bad.append(f"{path.name}:{n}: {line.strip()[:80]}")
    assert not bad, "Code or stray text between procedures:\n" + "\n".join(bad)

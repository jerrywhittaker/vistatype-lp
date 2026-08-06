#!/usr/bin/env python3
"""Refuse to build when a Styles("name") lookup can raise run-time error 5941.

ActiveDocument.Styles("X") raises 5941 -- "the requested member of the collection does not
exist" -- when the document does not carry that style, and a document need not: RefPageNemeth
comes from the Nemeth braille templates, Print Pg Num from the LP template, and so on. The
macro dies where it stands, often before it has done anything the user can see. Jerry hit it
on 8/5/2026 running AutoTag Ref Pages over an already-tagged braille file.

Every such lookup must therefore sit behind Sh_Style_Exists / Sh_Style_In_Use, or under an
On Error Resume Next that assigns into a variable and checks it for Nothing. "Normal" is
exempt: every Word document has it.
"""
import re, sys, pathlib

STYLE  = re.compile(r'ActiveDocument\.Styles\("([^"]+)"\)')
GUARD  = re.compile(r'Sh_Style_(?:Exists|In_Use)\(ActiveDocument, "([^"]+)"\)')
IFTHEN = re.compile(r'^\s*If\b.*\bThen\s*(?:\'.*)?$', re.I)
ENDIF  = re.compile(r'^\s*End If\b', re.I)
SETVAR = re.compile(r'^\s*Set \w+ = ActiveDocument\.Styles\(')
ALWAYS = {"Normal"}


def code(line: str) -> bool:
    return bool(line.strip()) and not line.strip().startswith("'")


def check(path: pathlib.Path):
    text = path.read_bytes().decode("utf-8", errors="replace").replace("\r\n", "\n")
    stack, onerr, bad = [], False, []
    for n, line in enumerate(text.split("\n"), 1):
        if not code(line):
            continue
        if re.match(r"^\s*On Error Resume Next", line, re.I):
            onerr = True
        elif re.match(r"^\s*On Error GoTo 0", line, re.I):
            onerr = False
        # Set st = Styles(...) under On Error, followed by an Is Nothing test, is the older
        # hand-written form of the same guard. Accept it.
        if onerr and SETVAR.match(line):
            continue
        opening = bool(IFTHEN.match(line))
        if opening:
            stack.append(set(GUARD.findall(line)))
        active = set().union(*stack) if stack else set()
        for m in STYLE.finditer(line):
            name = m.group(1)
            if name not in ALWAYS and name not in active:
                bad.append((n, name, line.strip()[:100]))
        if not opening and ENDIF.match(line) and stack:
            stack.pop()
    return bad


def main() -> int:
    root = pathlib.Path(__file__).resolve().parents[2] / "src" / "vba"
    failures = []
    for f in sorted(root.glob("*.bas")) + sorted(root.glob("*.cls")):
        for n, name, src in check(f):
            failures.append(f'  {f.relative_to(root.parents[1])}:{n}  Styles("{name}") is not guarded\n      {src}')
    if failures:
        print("Unguarded style lookups -- these raise run-time error 5941 on a document")
        print("that does not carry the style. Wrap with Sh_Style_Exists / Sh_Style_In_Use.")
        print("\n".join(failures))
        return 1
    print("style guards OK (every Styles(name) lookup is guarded)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

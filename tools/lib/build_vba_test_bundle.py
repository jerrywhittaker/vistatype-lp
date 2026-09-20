#!/usr/bin/env python3
"""Gather the VBA procedures a test module exercises into one importable module.

WHY THIS EXISTS
---------------
Most of the helpers worth unit-testing are `Private`, so a test document that merely
references the built add-in cannot call them. Referencing the add-in also tests the LAST
build rather than the source in front of you.

So the tests carry the source instead. Each test module names what it exercises:

    '@uses Sh_IsValidRomanNumeral

and this pulls that procedure out of src/vba, follows what it calls, brings the module-level
declarations those need, and writes one module the test document can import. `Private` becomes
`Public`, because the test modules are separate modules in the same project.

That is the recipe the headless work here has used since 8/2/2026 -- lift the sub's text out
of src/ and inject it -- with the lifting done on Linux, where it can be tested, instead of by
hand in PowerShell each time.

WHAT IT DOES NOT DO
-------------------
It does not compile anything, and it cannot: nothing on Linux compiles VBA. A procedure that
reaches a Document, a Selection or a MsgBox will be bundled quite happily and then hang or die
when the test calls it. Pick pure helpers; see tests/vba/README.md.

Usage:  python3 tools/lib/build_vba_test_bundle.py [outfile]
        (default build/vbatests/VtFunctionsUnderTest.bas)
"""
import glob
import os
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
TESTS = "tests/vba/*.bas"
SOURCES = ("src/vba/*.bas",)
DEFAULT_OUT = "build/vbatests/VtFunctionsUnderTest.bas"
MODULE_NAME = "VtFunctionsUnderTest"

PROC = re.compile(
    r"^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.IGNORECASE)
END_PROC = re.compile(r"^End\s+(Sub|Function|Property)\b", re.IGNORECASE)
USES = re.compile(r"^'\s*@uses\s+(\w+)\s*$", re.IGNORECASE)
# Any identifier. It is intersected with what src/vba actually defines, so matching Word's
# own names or a keyword costs nothing -- and unlike a prefix list it still finds a helper
# that does not carry one, and a CONSTANT, whose name is upper case (LP_FONT_SANS does not
# match the prefix "Lp_", which is how a missing declaration got through the first time).
IDENT = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b")
CONST = re.compile(
    r"^(?:Public\s+|Private\s+|Global\s+)?Const\s+(\w+)", re.IGNORECASE)
VAR = re.compile(
    r"^(?:Public|Private|Dim|Global)\s+(?:WithEvents\s+)?(\w+)\s", re.IGNORECASE)
BLOCK_OPEN = re.compile(r"^(?:Public\s+|Private\s+)?(Type|Enum)\s+(\w+)", re.IGNORECASE)
BLOCK_CLOSE = re.compile(r"^End\s+(Type|Enum)\b", re.IGNORECASE)


def read_lines(path):
    """A VBA source file as a list of lines, whatever it was written with."""
    data = pathlib.Path(path).read_bytes()
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode("cp1252", errors="replace")
    return text.replace("\r\n", "\n").split("\n")


def strip_comment(line):
    """The code half of a line: no trailing comment, no string literals.

    Walking it rather than splitting on the first apostrophe, because an apostrophe inside a
    string is not a comment -- and a procedure name mentioned in a comment is not a call.
    """
    out = []
    in_string = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_string = not in_string
        elif ch == "'" and not in_string:
            break
        elif not in_string:
            out.append(ch)
        i += 1
    return "".join(out)


def index_procedures(paths):
    """name -> (path, first line index, last line index). First definition wins."""
    found = {}
    for path in paths:
        lines = read_lines(path)
        i = 0
        while i < len(lines):
            m = PROC.match(lines[i].strip())
            if not m:
                i += 1
                continue
            name = m.group(3)
            start = i
            while i < len(lines) and not END_PROC.match(lines[i].strip()):
                i += 1
            end = min(i, len(lines) - 1)
            found.setdefault(name, (path, start, end))
            i += 1
    return found


def declarations(path):
    """The module's declarations section: name -> the lines that declare it."""
    lines = read_lines(path)
    decls = {}
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()
        if PROC.match(stripped):
            break
        block = BLOCK_OPEN.match(stripped)
        if block:
            start = i
            while i < len(lines) and not BLOCK_CLOSE.match(lines[i].strip()):
                i += 1
            decls[block.group(2)] = lines[start:i + 1]
            i += 1
            continue
        m = CONST.match(stripped) or VAR.match(stripped)
        if m:
            decls[m.group(1)] = [lines[i]]
        i += 1
    return decls


def seeds(test_paths):
    """Every '@uses name the test modules declare, in the order they appear."""
    wanted = []
    for path in sorted(test_paths):
        for line in read_lines(path):
            m = USES.match(line.strip())
            if m and m.group(1) not in wanted:
                wanted.append(m.group(1))
    return wanted


def make_public(line):
    """A Private procedure has to be Public to be callable from the test module."""
    return re.sub(r"^(\s*)Private\s+(?=(?:Static\s+)?(?:Sub|Function|Property)\b)",
                  r"\1Public ", line, count=1, flags=re.IGNORECASE)


def collect(wanted, procs, decls_by_file):
    """The procedures asked for, plus everything they call, plus the declarations they use."""
    chosen = []
    seen = set()
    queue = list(wanted)
    missing = []
    used_names = set()

    while queue:
        name = queue.pop(0)
        if name in seen:
            continue
        seen.add(name)
        if name not in procs:
            missing.append(name)
            continue
        path, start, end = procs[name]
        body = read_lines(path)[start:end + 1]
        chosen.append((name, path, body))
        for line in body:
            for ref in IDENT.findall(strip_comment(line)):
                used_names.add(ref)
                if ref not in seen and ref in procs:
                    queue.append(ref)

    needed_decls = []
    for path, decls in decls_by_file.items():
        for decl_name, lines in decls.items():
            if decl_name in used_names and decl_name not in seen:
                needed_decls.append((decl_name, lines))
    return chosen, needed_decls, missing


def render(chosen, decls):
    out = [
        'Attribute VB_Name = "%s"' % MODULE_NAME,
        "' GENERATED by tools/lib/build_vba_test_bundle.py - do not edit, and do not commit.",
        "'",
        "' The procedures tests/vba exercises, lifted out of src/vba as it stands now.",
        "' Private has become Public so the test modules can call them; nothing else is changed.",
        "",
        "Option Explicit",
        "",
    ]
    if decls:
        out.append("' --- module-level declarations these depend on ---")
        for name, lines in decls:
            out.extend(lines)
        out.append("")
    for name, path, body in chosen:
        out.append("' --- %s, from %s ---" % (name, os.path.relpath(path, ROOT)))
        out.append(make_public(body[0]))
        out.extend(body[1:])
        out.append("")
    return "\r\n".join(out) + "\r\n"


def build(root=ROOT, out_path=None):
    root = pathlib.Path(root)
    test_paths = sorted(glob.glob(str(root / TESTS)))
    source_paths = []
    for pattern in SOURCES:
        source_paths.extend(sorted(glob.glob(str(root / pattern))))

    wanted = seeds(test_paths)
    if not wanted:
        raise SystemExit("ERROR: no '@uses lines found in %s - nothing to bundle" % TESTS)

    procs = index_procedures(source_paths)
    decls_by_file = {p: declarations(p) for p in source_paths}
    chosen, decls, missing = collect(wanted, procs, decls_by_file)

    if missing:
        raise SystemExit(
            "ERROR: %d procedure(s) named by a '@uses line are not in src/vba:\n%s"
            % (len(missing), "".join("  - %s\n" % m for m in sorted(missing))))

    text = render(chosen, decls)
    out_path = pathlib.Path(out_path or (root / DEFAULT_OUT))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(text.encode("cp1252"))
    return out_path, [n for n, _, _ in chosen], [n for n, _ in decls]


def main():
    out_path, names, decls = build(out_path=sys.argv[1] if len(sys.argv) > 1 else None)
    asked = len(seeds(sorted(glob.glob(str(ROOT / TESTS)))))
    print("bundled %d procedure(s) into %s" % (len(names), out_path))
    print("  %d named by a test, %d pulled in because they are called"
          % (asked, len(names) - asked))
    if decls:
        print("  %d module-level declaration(s): %s" % (len(decls), ", ".join(decls)))
    return 0


if __name__ == "__main__":
    sys.exit(main())

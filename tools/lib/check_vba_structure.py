#!/usr/bin/env python3
"""Refuse to build on the VBA structure mistakes that ship silently.

Nothing in this build compiles VBA. `make build` imports the text into Word and Word
writes the .dotm; the code is not compiled until it loads on the transcriber's machine,
where a structural mistake appears as

    Compile error in hidden module: <name>

with no line number and nothing pointing at the cause. It cost a build on 8/18/2026: three
`Private Const` lines were written beside the procedures that used them instead of in the
module's declarations section, which VBA does not allow.

Three checks, all of them things a human reading a diff would not notice:

  1. A module-level declaration (Dim / Private / Public / Const / Type / Enum) below the
     first procedure. VBA requires the whole declarations section above every Sub.
  2. Two procedures with the same name in one module - a compile error where they sit.
  3. `With` without its `End With` inside a procedure.

Covers src/vba/*.bas, src/vba/*.cls and the code section of src/forms/*.frm (everything
after the designer header - see DEVELOPMENT.md for why that header is sacred).

Exit 0 = clean, 1 = something would fail on the user's machine.
"""

import collections
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

PROC = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.IGNORECASE)
END_PROC = re.compile(r"^End\s+(Sub|Function|Property)\b", re.IGNORECASE)
DECL = re.compile(r"^(Public|Private|Dim|Global|Const|Type|Enum|Declare)\b", re.IGNORECASE)
WITH_OPEN = re.compile(r"^With\b", re.IGNORECASE)
WITH_CLOSE = re.compile(r"^End\s+With\b", re.IGNORECASE)


def code_lines(path):
    """The lines that are VBA code, and their real line numbers.

    A .frm starts with the form designer's own block. Word parses that header itself and it
    must never be treated as code; the code section begins after the last Attribute line.
    """
    with open(path, "rb") as fh:
        data = fh.read()
    # Word exports forms in the system codepage, and src/ holds a mix; only the line
    # structure matters here, so fall back rather than fail on a curly apostrophe.
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode("cp1252", errors="replace")
    raw = text.split("\r\n")
    start = 0
    if path.lower().endswith(".frm"):
        for i, line in enumerate(raw):
            if line.startswith("Attribute "):
                start = i + 1
    return [(i + 1, raw[i]) for i in range(start, len(raw))]


def check(path):
    problems = []
    rel = os.path.relpath(path, ROOT)
    seen = collections.Counter()
    first_proc = None
    in_proc = False
    withs = 0
    proc_name = None
    proc_line = 0

    for lineno, line in code_lines(path):
        text = line.strip()
        if not text or text.startswith("'"):
            continue

        m = PROC.match(text)
        if m and not in_proc:
            in_proc = True
            withs = 0
            proc_name = m.group(2)
            proc_line = lineno
            seen[proc_name] += 1
            if first_proc is None:
                first_proc = (lineno, proc_name)
            continue

        if END_PROC.match(text):
            if in_proc and withs != 0:
                problems.append(
                    "%s:%d  %s has %d unclosed With (missing End With)"
                    % (rel, proc_line, proc_name, withs))
            in_proc = False
            withs = 0
            continue

        if in_proc:
            if WITH_OPEN.match(text):
                withs += 1
            elif WITH_CLOSE.match(text):
                withs -= 1
        elif first_proc is not None and DECL.match(text):
            problems.append(
                "%s:%d  module-level declaration below the first procedure "
                "(%s, line %d) - move it to the declarations section at the top:\n"
                "           %s"
                % (rel, lineno, first_proc[1], first_proc[0], text[:90]))

    for name, count in sorted(seen.items()):
        if count > 1:
            problems.append("%s  procedure %s is declared %d times in one module"
                            % (rel, name, count))
    return problems


def main():
    targets = []
    for pattern in ("src/vba/*.bas", "src/vba/*.cls", "src/forms/*.frm"):
        targets.extend(sorted(glob.glob(os.path.join(ROOT, pattern))))

    problems = []
    for path in targets:
        problems.extend(check(path))

    if problems:
        sys.stderr.write("VBA structure check FAILED - this would ship as "
                         "\"Compile error in hidden module\":\n\n")
        for p in problems:
            sys.stderr.write("  %s\n" % p)
        sys.stderr.write("\n")
        return 1

    print("VBA structure OK (%d files: declarations at top, no duplicate "
          "procedures, With/End With balanced)" % len(targets))
    return 0


if __name__ == "__main__":
    sys.exit(main())

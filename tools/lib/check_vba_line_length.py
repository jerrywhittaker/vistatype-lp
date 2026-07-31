#!/usr/bin/env python3
"""Refuse to build if any VBA source line is too long for Word to import.

WHY THIS EXISTS
---------------
A VBA source line may be at most 1023 characters. Go over it and `make build` does not fail
-- it HANGS. Word opens the .dotm, chokes silently on the line, and sits there forever with
no error and no dialog you can see, exactly like the stale ~$ lock-file trap. The lock file
Word leaves behind while stuck then looks like the cause, which sends you chasing the wrong
thing.

That happened on 7/30/2026: a 1110-character changelog line in the LPandBrlMacros header cost
three hung builds before the real cause turned up. The check is two seconds; the failure is
twenty minutes and a wrong diagnosis.

Checks .bas and .cls under src/vba, and .frm under src/forms -- forms carry code too.

Usage:  python3 tools/lib/check_vba_line_length.py
"""
import glob
import sys

LIMIT = 1023           # Word's hard limit for one physical line of VBA source
WARN_AT = 900          # close enough to be worth mentioning before it bites


def main():
    paths = sorted(glob.glob("src/vba/*.bas") + glob.glob("src/vba/*.cls") +
                   glob.glob("src/forms/*.frm"))
    if not paths:
        sys.exit("ERROR: no VBA source found - run from the repo root")

    too_long, near = [], []
    for path in paths:
        with open(path, "r", encoding="utf-8", errors="replace", newline="") as fh:
            for n, line in enumerate(fh, 1):
                length = len(line.rstrip("\r\n"))
                if length > LIMIT:
                    too_long.append((path, n, length, line.strip()[:60]))
                elif length > WARN_AT:
                    near.append((path, n, length))

    for path, n, length in near:
        print(f"  note: {path}:{n} is {length} chars (limit {LIMIT})")

    if too_long:
        print(f"\nERROR: {len(too_long)} VBA line(s) longer than {LIMIT} characters.")
        print("Word does not report this - it HANGS on import, with no error and no visible")
        print("dialog. Wrap the line before building.\n")
        for path, n, length, preview in too_long:
            print(f"  {path}:{n}  {length} chars")
            print(f"      {preview}...")
        sys.exit(1)

    print(f"VBA line lengths OK ({len(paths)} files, limit {LIMIT})")


if __name__ == "__main__":
    main()

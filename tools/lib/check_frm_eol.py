#!/usr/bin/env python3
"""Fail the build if any UserForm .frm has bare LF line endings.

Word parses the designer header (VERSION / Begin {GUID} ... End) at the top of a .frm before
the code, and that parser needs CRLF. With LF the header is not recognized, Word treats those
lines as VBA source and drops them into the form's code module. The build still succeeds -- the
failure surfaces only on the user's machine, as "Compile error in hidden module: <FormName>".
Editing a .frm from Linux in text mode is enough to cause it. See DEVELOPMENT.md.
"""
import glob, sys

bad = []
for path in sorted(glob.glob("src/forms/*.frm")):
    data = open(path, "rb").read()
    bare_lf = data.count(b"\n") - data.count(b"\r\n")
    if bare_lf:
        bad.append((path, bare_lf))

if bad:
    print("ERROR: these .frm files have bare LF line endings and will not import correctly:", file=sys.stderr)
    for path, n in bad:
        print(f"  {path}  ({n} bare LF)", file=sys.stderr)
    print('\nFix: read the file as bytes, replace b"\\n" with b"\\r\\n", write it back as bytes.', file=sys.stderr)
    sys.exit(1)

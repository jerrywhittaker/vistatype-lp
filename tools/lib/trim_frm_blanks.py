#!/usr/bin/env python3
"""Collapse the blank lines Word piles up in an exported .frm.

Every time Word exports a UserForm it inserts one more blank line between the Attribute
block and the first line of code, and one more at the end of the file. Nothing reads them,
but the two About dialogs are re-exported on EVERY build (the version stamper rewrites their
captions), so they had grown to 50 and 43 blank lines by 8/1/2026 and gained four more on a
five-build day. Left alone they make a one-word caption change look like a real edit.

This collapses both runs to a single blank line and leaves everything else exactly as it is,
including blank lines between subs -- those are the author's formatting, not export noise.

CRLF is preserved: a .frm rewritten with bare LF makes Word dump the designer header into the
form's code module, which fails only on the user's machine (see check_frm_eol.py).
"""
import sys


def trim(path):
    with open(path, 'rb') as fh:
        raw = fh.read()
    if b'\r\n' not in raw:
        sys.exit('%s: expected CRLF line endings, refusing to touch it' % path)

    lines = raw.decode('cp1252').split('\r\n')

    # The attribute block sits at the top and ends at the last line starting "Attribute ".
    last_attr = None
    for i, line in enumerate(lines):
        if line.startswith('Attribute '):
            last_attr = i
    if last_attr is None:
        return path, 0          # not a form module; nothing to do

    head = lines[:last_attr + 1]
    body = lines[last_attr + 1:]

    # Leading run: exactly one blank line before the first line of code.
    first_code = 0
    while first_code < len(body) and body[first_code].strip() == '':
        first_code += 1
    body = [''] + body[first_code:]

    # Trailing run: exactly one blank line at the end of the file.
    while len(body) > 1 and body[-1].strip() == '' and body[-2].strip() == '':
        body.pop()

    out = head + body
    removed = len(lines) - len(out)
    if removed:
        with open(path, 'wb') as fh:
            fh.write('\r\n'.join(out).encode('cp1252'))
    return path, removed


def main(paths):
    for p in paths:
        path, removed = trim(p)
        if removed:
            print('trimmed %d blank line(s) from %s' % (removed, path))


if __name__ == '__main__':
    main(sys.argv[1:])

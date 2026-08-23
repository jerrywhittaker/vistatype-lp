#!/usr/bin/env python3
"""Build the no-installer way of giving somebody the typeface: a .zip they unpack and right-click.

Why this exists beside the font-only Setup.exe. An unsigned installer can be deleted by antivirus
on sight - every build of the main Setup.exe was, for two days in August 2026, by a
machine-learning guess with no malware in it anywhere (docs/Code-Signing.md). Nothing can promise
an unsigned .exe will not be flagged. A .zip of four font files can make that promise, because
there is no program in it to judge: Windows installs a font from right-click > Install, for the
current user, with no administrator rights.

The three SIL Open Font License texts go in with them. That is not politeness - condition 2 of
the OFL wants each license on disk beside the font it covers, and a font handed to someone
without it is a breach nobody would ever notice.

Usage:  python3 tools/lib/build_font_zip.py <out.zip>
"""
import os
import sys
import zipfile

FONT_DIR = 'assets/fonts/vistatypelp-sans'
FACES = [
    'VistaTypeLPSans-Regular.ttf',
    'VistaTypeLPSans-Bold.ttf',
    'VistaTypeLPSans-Italic.ttf',
    'VistaTypeLPSans-BoldItalic.ttf',
]
LICENSES = ['OFL.txt', 'OFL-NotoSansMath.txt', 'OFL-NotoSansSymbols.txt']

README = """VistaTypeLP Sans - how to install it
====================================

This is the typeface VistaType LP sets large print books in. There is no program to
run here, and you do not need administrator rights.

TO INSTALL IT

  1. Close Microsoft Word. A font that arrives while Word is open does not show up
     in it until Word is started again.

  2. Select all four files whose names end in .ttf:

        VistaTypeLPSans-Regular.ttf
        VistaTypeLPSans-Bold.ttf
        VistaTypeLPSans-Italic.ttf
        VistaTypeLPSans-BoldItalic.ttf

  3. Right-click them and choose "Install".

     If you are offered "Install" and "Install for all users", choose plain
     "Install" - the other one needs administrator rights and you do not need it.

  4. Open Word. "VistaTypeLP Sans" is now in the font list.

INSTALL ALL FOUR, NOT JUST THE FIRST

Word does not fall back within a font family. If the Bold file is missing, bold text
comes out of some other typeface at a different size, without saying so - and in a
large print book, a character at the wrong size is the one thing that must never
happen.

ABOUT THE THREE LICENSE FILES

The typeface is built from Google's Noto Sans, with Noto Sans Math and Noto Sans
Symbols folded into it so that mathematics, Greek and the phonetic characters are all
there at the right size. Each of those carries the SIL Open Font License, and the
three OFL files here are their terms. Keep them with the fonts.

TO REMOVE IT LATER

Settings > Personalization > Fonts, find VistaTypeLP Sans, and uninstall it. Windows
will not let go of a font while it is in use, so close Word first.

Jerry Whittaker - jerry@thewhittakers.org
"""


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    out = sys.argv[1]

    missing = [f for f in FACES + LICENSES if not os.path.isfile(os.path.join(FONT_DIR, f))]
    if missing:
        print('MISSING from %s: %s' % (FONT_DIR, ', '.join(missing)))
        return 1

    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in FACES + LICENSES:
            z.write(os.path.join(FONT_DIR, f), 'VistaTypeLP Sans/' + f)
        # CRLF: it is read on Windows, very likely in Notepad.
        z.writestr('VistaTypeLP Sans/READ ME FIRST.txt', README.replace('\n', '\r\n'))

    print('wrote %s (%d faces, %d licenses, and the instructions)'
          % (out, len(FACES), len(LICENSES)))
    return 0


if __name__ == '__main__':
    sys.exit(main())

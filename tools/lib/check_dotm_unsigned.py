#!/usr/bin/env python3
"""Refuse to build from, or promote, a .dotm whose VBA project carries a digital signature.

The repo-root LPandBRL.dotm is the BASE every build starts from: push-src sends it to the build
box and Import-Vba.ps1 copies it and imports the modules into that copy. Once signing is wired
in, the file in dist/ is signed -- and `make deploy` copies dist/LPandBRL.dotm to the repo root.
A signed file reaching the root therefore seeds every later build with a signature that no
longer matches the code once the modules are re-imported.

Nothing else in this repo would notice. See docs/Code-Signing.md, "A signed .dotm must never
travel back to this machine".

WHERE THE SIGNATURE ACTUALLY LIVES -- measured on the build box 9/20/2026, by signing a real
copy of the base with a throwaway certificate and comparing the two files. It is NOT a stream
inside vbaProject.bin: that part came back byte-identical, all 302 streams the same size. The
signature is its own part in the Office package, alongside it:

    word/vbaProjectSignature.bin          the legacy format
    word/vbaProjectSignatureAgile.bin     the agile format
    word/vbaProjectSignatureV3.bin        the 2020 format

Word wants all three and signtool writes one per pass, which is why offsign.bat exists. The
first version of this guard searched the OLE streams inside vbaProject.bin -- the layout the
old binary .doc format used -- and passed a genuinely signed file. Matching is by PREFIX so
that a fourth format, whenever it appears, is caught rather than quietly ignored.

Usage:
    python3 tools/lib/check_dotm_unsigned.py [file.dotm ...]      # default: LPandBRL.dotm
"""
import sys
import zipfile

DEFAULT = ["LPandBRL.dotm"]

# Package parts that mean "this VBA project is signed".
VBA_SIG_PREFIX = "word/vbaprojectsignature"
# A signature over the whole package rather than the VBA project.
PACKAGE_SIG_PREFIX = "_xmlsignatures/"


def signature_parts(dotm_path):
    """Return the names of any signature-bearing parts in the .dotm, as the file spells them."""
    with zipfile.ZipFile(dotm_path) as z:
        return [
            name
            for name in z.namelist()
            if name.lower().startswith((VBA_SIG_PREFIX, PACKAGE_SIG_PREFIX))
        ]


def main(argv):
    paths = argv[1:] or DEFAULT
    bad = []

    for path in paths:
        try:
            parts = signature_parts(path)
        except FileNotFoundError:
            continue  # nothing to check; other targets report a missing .dotm
        except zipfile.BadZipFile:
            print(f"ERROR: {path} is not a readable Word file.", file=sys.stderr)
            return 1
        if parts:
            bad.append((path, parts))

    if not bad:
        return 0

    print("ERROR: this .dotm carries a VBA digital signature and must not be used as a",
          file=sys.stderr)
    print("       build base. Every later build would start from it.", file=sys.stderr)
    print("", file=sys.stderr)
    for path, parts in bad:
        print(f"  {path}", file=sys.stderr)
        for p in parts:
            print(f"      {p}", file=sys.stderr)
    print("", file=sys.stderr)
    print("What to do:", file=sys.stderr)
    print("  - If it was promoted by mistake, put the previous one back:", file=sys.stderr)
    print("        git checkout HEAD -- LPandBRL.dotm", file=sys.stderr)
    print("  - Sign the file in dist/ on its way into the installer, never the repo root.",
          file=sys.stderr)
    print("    The signing step belongs between the SHIPFILES copy and ISCC in installer-build.",
          file=sys.stderr)
    print("  - To strip a signature deliberately, offsign.bat clears one before writing a new", file=sys.stderr)
    print("    one; see docs/Code-Signing.md.", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))

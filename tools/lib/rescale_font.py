#!/usr/bin/env python3
"""Rescale a font so Word's point size and the printed letter size finally agree.

WHY THIS EXISTS
---------------
A point size in Word sets the *em* -- the notional body the type sits on -- not the height
of any letter. Every typeface decides for itself how much of that body the letters fill, and
Atkinson Hyperlegible fills less of it than Tahoma does: measured against Tahoma at the same
nominal size, its cap height is 8.1% smaller, its x-height 9.1%, a real sentence 8.6%.

Jerry measured it with a font ruler on 8/8/2026 and read it as roughly two points short at
20 and 22 point. He is right, and it matters more here than in ordinary typesetting: large
print is specified in points, and a transcriber who sets 18 point believes they have produced
18 point. In unmodified Atkinson they have produced something that measures about 16.5.

Two ways to fix that. Bump every size in the template -- but then Word says 20 where the
transcriber means 18, and everyone downstream is confused. Or scale the font itself, which is
what this script does: after rescaling, the transcriber sets 18, Word reports 18, and the
ruler reads 18.

HOW
---
Scale every outline and every metric up by FACTOR while leaving unitsPerEm at 1000. The
glyphs then occupy more of the same em, so they render larger at the same point size.
fontTools' scale_upem does the real work -- it moves outlines, advance widths, kerning,
vertical metrics and the OS/2 values together, which is the part that is easy to get wrong by
hand. We then put unitsPerEm back to 1000 (and pin the CFF FontMatrix to match) so the scale
sticks instead of cancelling out.

RENAMING IS NOT OPTIONAL
------------------------
Atkinson Hyperlegible is licensed under the SIL Open Font License 1.1, which reserves two
words: **ATKINSON** and **HYPERLEGIBLE**. Condition 3 forbids a modified version from using
either in the name presented to users, so this script REFUSES a --family containing them.
Two independent reasons, and each alone would be enough:

  * License. Using a reserved name in a modified version breaches condition 3, and the
    license terminates outright if any condition is not met.
  * Practice. Jerry and his transcribers have the real Atkinson installed. Two different
    fonts with one name is not a supported state -- Word picks whichever it happens to find,
    and the same document sets differently on different machines.

The naming restriction covers only the primary font name, so the origin is still credited in
the description record, which is where the OFL expects that acknowledgment to live.

WHAT ELSE THE LICENSE MAKES US DO
---------------------------------
Condition 2: every copy must carry the copyright notice and the license. The upstream 1.02
files predate the OFL grant -- their license record still reads "without derivatives or
alteration", which is superseded but would make our derivative contradict itself. So this
script REWRITES the copyright, license and license-URL records to the OFL -- in the name table
AND in the CFF table, which keeps its own copy of the names and copyright. Shipping the font
also means shipping OFL.txt beside it; the font's own records are not enough on their own.

Condition 4: the Braille Institute's name may not be used to promote a modified version, only
to acknowledge it. The vendor record therefore points at this project, not at them, while the
designer credit stays exactly as it was.

Condition 5: the font stays under the OFL. VistaType LP is GPLv3; the bundled font is not,
and the installer has to say so. The OFL expressly allows bundling with any software.

Any DSIG is dropped: a digital signature over the old outlines is meaningless once they move,
and a stale one is worse than none.

FEED IT THE TTF, NOT THE OTF
----------------------------
The output keeps the outline format of the input, and that choice decides whether the font can
be embedded in a Word document at all. Word embeds TrueType outlines and skips PostScript/CFF
ones without a word -- the file simply saves without the font. So the shipping faces are
rescaled from upstream/ttf/, not upstream/otf/.

Usage:
    python3 tools/lib/rescale_font.py --factor 1.094 \
        --family "VistaTypeLP Legible" \
        --out assets/fonts/atkinson-hyperlegible/scaled \
        assets/fonts/atkinson-hyperlegible/upstream/ttf/*.ttf
"""
import argparse
import array
import pathlib
import sys

from fontTools.ttLib import TTFont
from fontTools.ttLib.scaleUpem import scale_upem

WINDOWS = (3, 1, 0x409)   # platformID, encodingID, langID -- the record Word reads
MAC = (1, 0, 0)

# name IDs that carry the family identity and must be rewritten together
COPYRIGHT = 0
FAMILY = 1
SUBFAMILY = 2
UNIQUE = 3
FULL = 4
VERSION = 5
POSTSCRIPT = 6
MANUFACTURER = 8
DESCRIPTION = 10
VENDOR_URL = 11
LICENSE = 13
LICENSE_URL = 14
TYPO_FAMILY = 16
TYPO_SUBFAMILY = 17
COMPATIBLE_FULL = 18   # Mac-only legacy full name, capped at 31 characters by the spec

# Reserved Font Names from the OFL grant (assets/fonts/atkinson-hyperlegible/OFL.txt).
# Condition 3: a Modified Version may not use these in the name shown to users.
RESERVED_NAMES = ("atkinson", "hyperlegible")

# Condition 2: each copy carries the copyright notice and the license. These replace the
# upstream 1.02 records, which still carry the superseded "no derivatives" wording.
OFL_COPYRIGHT = (
    'Copyright 2020, Braille Institute of America, Inc. '
    '(https://www.brailleinstitute.org/), with Reserved Font Names: '
    '"ATKINSON" and "HYPERLEGIBLE". '
    'Modifications copyright 2026, Jerry Whittaker, for VistaType LP.'
)
OFL_LICENSE = (
    'This Font Software is licensed under the SIL Open Font License, Version 1.1. '
    'This license is available with a FAQ at: https://openfontlicense.org'
)
OFL_LICENSE_URL = 'https://openfontlicense.org'

PROJECT = 'VistaType LP'
PROJECT_URL = 'https://github.com/jerrywhittaker/vistatype-lp'


def ps_safe(s):
    """PostScript names allow no spaces and a limited character set."""
    return "".join(c for c in s if c.isalnum() or c == "-")


def rescale(path, out_dir, factor, new_family, dry_run=False):
    # recalcTimestamp=False keeps upstream's "modified" date instead of stamping now. Without
    # it, two runs of the same command produce different bytes, so re-running this would show
    # up as a change to four tracked binaries that are in fact identical -- and no checksum
    # could ever tell a stale font from a fresh one.
    f = TTFont(path, recalcTimestamp=False)
    name = f["name"]
    old_family = name.getDebugName(FAMILY)
    subfamily = name.getDebugName(SUBFAMILY) or "Regular"

    # --- the scale itself -------------------------------------------------
    # scale_upem moves everything to a larger em; putting the em back to 1000 without
    # rescaling again is what actually makes the glyphs bigger relative to the body.
    target = round(1000 * factor)
    scale_upem(f, target)
    f["head"].unitsPerEm = 1000
    if "CFF " in f:
        top = f["CFF "].cff[f["CFF "].cff.fontNames[0]]
        top.rawDict["FontMatrix"] = [0.001, 0, 0, 0.001, 0, 0]

    # scale_upem leaves the TrueType hinting alone, so the control-value table still describes
    # the ORIGINAL cap height and x-height. The hinting then pulls letters back toward the size
    # we just moved them away from. It makes no difference at large-print sizes or in print --
    # the gap is far wider than the snap tolerance -- but at small sizes on screen, a zoomed-out
    # page or Word's own font menu, it would undo part of the rescale.
    # cvt values are an array of int16, and the table will not compile if that is replaced with
    # a plain list -- so rebuild it as an array of the same typecode.
    if "cvt " in f:
        cvt = f["cvt "].values
        f["cvt "].values = array.array(cvt.typecode,
                                       [round(v * factor) for v in cvt])

    # Make two rescales at different factors distinguishable. Nothing else varies between
    # builds, so without this neither Windows, nor an installer, nor a support call can tell
    # which rescale a transcriber is running.
    f["head"].fontRevision = round(f["head"].fontRevision + (factor - 1.0), 4)

    # --- make the two upstream formats agree on line spacing --------------
    # Upstream's OTF and TTF carry IDENTICAL typo and win metrics but disagree on which of
    # them applies: the OTF sets USE_TYPO_METRICS and gives hhea a 150-unit line gap, the TTF
    # does neither. Word honors USE_TYPO_METRICS, so the same design would set lines about
    # 17% further apart as a TTF than as an OTF -- a change nobody asked for, arriving purely
    # because we switched format to make embedding possible.
    #
    # So pin the TrueType build to the PostScript build's behavior. This normalizes an
    # upstream inconsistency; it does not invent a new line height. Jerry evaluated the OTF,
    # and what ships now spaces its lines the same way.
    #
    # The version bump is required, not tidying: USE_TYPO_METRICS is bit 7 of fsSelection,
    # which is only DEFINED from OS/2 version 4. Upstream's OTF is already v4 (hence the flag
    # is set there); the TTF is v3, where setting the bit is meaningless and fontTools warns.
    # v4 adds no fields over v3, so the bump is free.
    if f["OS/2"].version < 4:
        f["OS/2"].version = 4
    f["OS/2"].fsSelection |= (1 << 7)                            # USE_TYPO_METRICS
    f["hhea"].ascender = f["OS/2"].sTypoAscender
    f["hhea"].descender = f["OS/2"].sTypoDescender
    f["hhea"].lineGap = f["OS/2"].sTypoLineGap

    # --- rename -----------------------------------------------------------
    full = f"{new_family} {subfamily}".strip()
    ps = f"{ps_safe(new_family)}-{ps_safe(subfamily)}"
    if len(ps) > 63:
        ps = ps[:63]

    for plat in (WINDOWS, MAC):
        pid, eid, lid = plat
        def put(nid, value):
            name.setName(value, nid, pid, eid, lid)
        if name.getName(FAMILY, pid, eid, lid) is None and plat is MAC:
            continue
        put(FAMILY, new_family)
        put(FULL, full)
        put(POSTSCRIPT, ps)
        # The unique ID is not the primary name, but keep the reserved words out of it too:
        # some tools surface it, and there is nothing to gain by being close to the line.
        put(UNIQUE, f"{full}; rescaled x{factor}; {PROJECT}")
        if name.getName(TYPO_FAMILY, pid, eid, lid) is not None:
            put(TYPO_FAMILY, new_family)
        # The upstream TTFs carry a Mac "compatible full name" that the OTFs do not, and the
        # post-write check refuses the file if a reserved word survives in it. The spec caps
        # it at 31 characters, which is why upstream reads "Atkinson Hyperlegible Bold It".
        if name.getName(COMPATIBLE_FULL, pid, eid, lid) is not None:
            put(COMPATIBLE_FULL, full[:31])
        ver = name.getName(VERSION, pid, eid, lid)
        if ver is not None:
            put(VERSION, f"{ver.toUnicode()}; rescaled x{factor} for {PROJECT}")

        # --- license records: condition 2 (see module docstring) ----------
        put(COPYRIGHT, OFL_COPYRIGHT)
        put(LICENSE, OFL_LICENSE)
        put(LICENSE_URL, OFL_LICENSE_URL)

        # This file was produced here, so it says so; the designer credit (name 9) and the
        # designer's URL (name 12) stay untouched -- that is the acknowledgment condition 4
        # allows. The vendor URL must NOT keep pointing at the Braille Institute, which would
        # read as their endorsement of a version they did not make.
        put(MANUFACTURER, PROJECT)
        put(VENDOR_URL, PROJECT_URL)
        put(DESCRIPTION,
            f"{new_family} is {PROJECT}'s rescaled version of {old_family}, the typeface "
            "created by Applied Design Works for Braille Institute of America, Inc. to "
            "increase legibility for readers with low vision. Every outline and metric is "
            f"scaled by {factor} while the em stays at 1000 units, so that a point size set "
            "in Word matches the size of the printed letters -- large print is specified in "
            "points, and the unmodified face sets about 9% smaller than that size implies. "
            "The letterforms are otherwise unaltered. Not endorsed by, or connected with, "
            "Braille Institute of America, Inc.")

    # --- the CFF table carries its OWN copy of the names ------------------
    # Easy to miss, and it is not an internal detail: the CFF FontName is what Word writes
    # into a PDF as /BaseFont. Leave it alone and a transcriber's finished large-print PDF
    # reports "AtkinsonHyperlegible-Regular" -- our altered font wearing the Braille
    # Institute's name (condition 3), and worse, a name a print shop's RIP can resolve to the
    # REAL Atkinson, substituting the 9%-too-small face back in at the very last step.
    if "CFF " in f:
        cff = f["CFF "].cff
        cff.fontNames[0] = ps
        top = cff[ps]
        top.rawDict["FullName"] = full
        top.rawDict["FamilyName"] = new_family
        top.rawDict["Copyright"] = OFL_COPYRIGHT

    # --- a signature over moved outlines is meaningless -------------------
    if "DSIG" in f:
        del f["DSIG"]

    # --- write it out in the format it came in ---------------------------
    # TrueType outlines get .ttf, PostScript/CFF outlines get .otf. This is not cosmetic and
    # it is why we rescale the upstream ttf/ rather than the otf/: Word's EmbedTrueTypeFonts
    # means exactly what it says. It embeds TrueType outlines and skips CFF SILENTLY -- no
    # error, no warning, the document just saves without the font, and a reader who does not
    # have it installed gets a substitute at the wrong size. That is the whole defect this
    # font exists to cure, reappearing at the last step.
    ext = ".ttf" if "glyf" in f else ".otf"
    out = pathlib.Path(out_dir) / f"{ps}{ext}"
    if dry_run:
        print(f"  would write {out}")
        return out
    out.parent.mkdir(parents=True, exist_ok=True)
    f.save(out)
    verify_no_reserved_names(out)
    return out


# Name IDs a user or a tool may see AS the font's name. The OFL restricts the reserved words
# here; IDs 0 (copyright) and 10 (description) are excluded on purpose -- the license wants
# the reserved names declared in the first and permits acknowledgment in the second.
PRIMARY_NAME_IDS = (1, 3, 4, 6, 16, 18, 20, 21, 25)


def verify_no_reserved_names(path):
    """Re-open what was just written and prove the reserved words are gone.

    The --family guard checks the ARGUMENT. This checks the RESULT, which is the only thing
    that ships. That distinction is not academic: the first version of this script passed its
    own guard and still wrote "AtkinsonHyperlegible-Regular" into every file's CFF table.
    """
    f = TTFont(path)
    hits = []
    for rec in f["name"].names:
        if rec.nameID in PRIMARY_NAME_IDS:
            v = rec.toUnicode().lower()
            hits += [f"name ID {rec.nameID}: {rec.toUnicode()!r}"
                     for w in RESERVED_NAMES if w in v]
    if "CFF " in f:
        cff = f["CFF "].cff
        candidates = [("CFF FontName", cff.fontNames[0])]
        top = cff[cff.fontNames[0]]
        for k in ("FullName", "FamilyName"):
            if k in top.rawDict:
                candidates.append((f"CFF {k}", top.rawDict[k]))
        hits += [f"{label}: {value!r}"
                 for label, value in candidates
                 for w in RESERVED_NAMES if w in str(value).lower()]
    if hits:
        raise SystemExit(
            f"ERROR: {path} still carries Reserved Font Names in records users can see:\n  "
            + "\n  ".join(hits)
            + "\nThis breaches OFL condition 3. The file was written; delete it.")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("fonts", nargs="+")
    ap.add_argument("--factor", type=float, required=True,
                    help="scale for outlines and metrics, e.g. 1.094")
    ap.add_argument("--family", required=True,
                    help="new family name -- must differ from the original, see module docstring")
    ap.add_argument("--out", required=True)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    if a.factor <= 0:
        sys.exit("factor must be positive")

    # Reserved Font Names, checked ONCE before anything is written. Matching on the bare
    # word is deliberate: "Atkinson Hyperlegible (VT-LP)" is a different string from the
    # original and would sail past an equality test, but it still uses both reserved words
    # and still breaches condition 3.
    hit = [w for w in RESERVED_NAMES if w in a.family.lower()]
    if hit:
        sys.exit(
            f"ERROR: --family {a.family!r} contains the Reserved Font Name(s) "
            f"{', '.join(repr(w.upper()) for w in hit)}.\n"
            "The SIL Open Font License forbids a modified version using them in the name "
            "shown to users, and the license terminates if any condition is not met.\n"
            "Choose a family name with neither word. See the RENAMING IS NOT OPTIONAL "
            "section of this script, and assets/fonts/atkinson-hyperlegible/OFL.txt.")

    for p in a.fonts:
        src = pathlib.Path(p)
        old = TTFont(src)["name"].getDebugName(FAMILY)
        if old and old.strip().lower() == a.family.strip().lower():
            sys.exit(f"ERROR: --family must differ from the original name {old!r}. "
                     "See the RENAMING IS NOT OPTIONAL section of this script.")
        out = rescale(src, a.out, a.factor, a.family, a.dry_run)
        print(f"  {src.name}  ->  {out.name}")


if __name__ == "__main__":
    main()

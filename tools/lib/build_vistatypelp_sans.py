#!/usr/bin/env python3
"""
build-vistatypelp-sans.py — build VistaTypeLP Sans from upstream Noto Sans.

Pipeline, all of it repeatable from an upstream release:

  1. fetch    Noto Sans variable TTF from google/fonts (OFL)
  2. slash    repoint cmap U+0030 -> zero.slash, so the slashed zero is the DEFAULT
              glyph and needs no OpenType feature (Word cannot reach the `zero` feature)
  3. instance cut static Regular (wght 400) and Bold (wght 700) — Word handles
              variable fonts poorly and will not give real Bold
  4. scale    unitsPerEm 1000 -> 960, an exact 25/24 enlargement, so nominal point
              sizes match the VistaType ruler (measured: 37.5pt matched the ruler's 36)
  5. symbols  fold in Noto Sans Math and Noto Sans Symbols, into EVERY face -- Word does
              not fall back inside a family (tested: a bold character the Bold face lacked
              came out of Cambria Math, another typeface at another size)
  6. metrics  retune vertical metrics so Word's "Single" line spacing matches Tahoma
  7. rename   per the OFL Reserved Font Name clause

Output: TrueType (glyf) .ttf, fsType 0, ready for Word embedding.

Usage:
    python3 build-vistatypelp-sans.py [--out DIR] [--offline SRC.ttf] [--no-verify]
"""
from __future__ import annotations
import argparse, math, os, sys, urllib.request

try:
    from fontTools.ttLib import TTFont
    from fontTools.varLib import instancer
    from fontTools.pens.boundsPen import BoundsPen
    from fontTools.ttLib.tables.O_S_2f_2 import OS2_UNICODE_RANGES
except ImportError:
    sys.exit("fontTools is required:  pip install fonttools")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from merge_symbol_glyphs import merge_glyphs

DONORS = {
    "math":    ("https://raw.githubusercontent.com/google/fonts/main/ofl/notosansmath/"
                "NotoSansMath-Regular.ttf"),
    "symbols": ("https://raw.githubusercontent.com/google/fonts/main/ofl/notosanssymbols/"
                "NotoSansSymbols%5Bwght%5D.ttf"),
}

SOURCES = {
    "roman":  ("https://raw.githubusercontent.com/google/fonts/main/ofl/notosans/"
               "NotoSans%5Bwdth%2Cwght%5D.ttf"),
    "italic": ("https://raw.githubusercontent.com/google/fonts/main/ofl/notosans/"
               "NotoSans-Italic%5Bwdth%2Cwght%5D.ttf"),
}

FAMILY   = "VistaTypeLP Sans"
PSPREFIX = "VistaTypeLPSans"
RULER_NUM, RULER_DEN = 25, 24          # 37.5pt measured == 36pt on the VistaType ruler
TARGET_LINE_EM = 2472 / 2048           # Tahoma's "Single" line height, 1.20703 em

# subfamily, weight, which source, is-italic.  All four so a book can be set properly:
# Word fakes an italic it has not got, and a faked italic of a rescaled face is the wrong size.
INSTANCES = [
    ("Regular",     400, "roman",  False),
    ("Bold",        700, "roman",  False),
    ("Italic",      400, "italic", True),
    ("Bold Italic", 700, "italic", True),
]

# Characters VistaTypeLP Sans must set. Drives the no-clipping check on win metrics.
REQUIRED = (
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝàáâãäåçèéêëìíîïñòóôõöøùúûüýÿ"
    "ĄĆĘŁŃŚŹŻąćęłńśźżČŠŽčšžŐŰőűĞİışğ"
    "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρσςτυφχψωάέήίόύώΐΰϊϋ"
    "ɞəɛɔɪʊʌæŋʃʒθðɑɒøœɜɐɘɵɤɯˈˌː"
    "!\"#$%&'()*+,-./:;<=>?@[]^_{|}~«»“”‘’–—…"
    # Mathematics, logic and arrows. This is the third leg of the coverage test the installer's
    # own note demands -- "Greek, the IPA, the math operators" -- and the one VistaTypeLP
    # Legible failed. It is listed here so the BUILD fails on a face that cannot set it, rather
    # than a transcriber finding out when Word substitutes at the wrong size.
    "≤≥≠≈∞√∑∏∫∂∆∇∀∃∈∉⊂⊆∪∩±×÷←→↑↓⇌⇔"
    "ℝℕℤℚℂℵℏℓ∅"          # letterlike: real numbers, aleph, Planck, empty set
    "☉♀♂♁♃♄♭♮♯"           # astronomy and music, from Noto Sans Symbols
)


def log(msg): print(msg, flush=True)


def fetch_donor(which: str, offline_dir: str | None, workdir: str) -> str:
    """Noto Sans Math / Noto Sans Symbols, whose characters get folded into every face."""
    name = f"NotoSans{which.capitalize()}-source.ttf"
    if offline_dir:
        local = os.path.join(offline_dir, name)
        if not os.path.exists(local):
            raise SystemExit(f"FAIL: --offline given but {local} is not there")
        return local
    dst = os.path.join(workdir, name)
    if not os.path.exists(dst):
        log(f"      fetching {which} donor")
        urllib.request.urlretrieve(DONORS[which], dst)
    return dst


def fetch(which: str, offline_dir: str | None, workdir: str) -> str:
    """Return a path to the Noto Sans source for `which` ('roman' or 'italic')."""
    name = "NotoSans-source.ttf" if which == "roman" else "NotoSans-Italic-source.ttf"
    if offline_dir:
        local = os.path.join(offline_dir, name)
        if not os.path.exists(local):
            raise SystemExit(f"FAIL: --offline given but {local} is not there")
        log(f"[1/8] using local source {local}")
        return local
    dst = os.path.join(workdir, name)
    if os.path.exists(dst):
        log(f"[1/8] reusing {name}")
        return dst
    log(f"[1/8] fetching {which} source")
    urllib.request.urlretrieve(SOURCES[which], dst)
    log(f"      {os.path.getsize(dst)//1024} KB")
    return dst


def slash_zero(font: TTFont) -> None:
    """Make the slashed zero the default glyph for U+0030."""
    zero = font.getBestCmap()[0x30]
    alt = None
    gsub = font["GSUB"].table
    lookups = gsub.LookupList.Lookup
    for fr in gsub.FeatureList.FeatureRecord:
        if fr.FeatureTag != "zero":
            continue
        for li in fr.Feature.LookupListIndex:
            for st in lookups[li].SubTable:
                m = getattr(st, "mapping", {}) or {}
                if zero in m:
                    alt = m[zero]
    if alt is None:
        raise SystemExit("FAIL: upstream has no slashed zero behind the 'zero' feature")
    n = 0
    for t in font["cmap"].tables:
        if 0x30 in t.cmap:
            t.cmap[0x30] = alt
            n += 1
    log(f"[2/8] slashed zero: U+0030 -> {alt} in {n} cmap subtable(s)")


def ink_bounds(font: TTFont) -> tuple[int, int]:
    """Highest and lowest ink across REQUIRED, in whole font units.

    Rounded outwards. Italic outlines are slanted, so their bounds come back
    fractional, and the OS/2 metrics these feed are unsigned 16-bit integers.
    """
    gs, cm = font.getGlyphSet(), font.getBestCmap()
    hi, lo = 0, 0
    for ch in REQUIRED:
        g = cm.get(ord(ch))
        if not g or g not in gs:
            continue
        bp = BoundsPen(gs)
        try:
            gs[g].draw(bp)
        except Exception:
            continue
        if bp.bounds:
            hi, lo = max(hi, bp.bounds[3]), min(lo, bp.bounds[1])
    return math.ceil(hi), math.floor(lo)


def retune_metrics(font: TTFont) -> None:
    """Make Word's 'Single' spacing match Tahoma, without clipping any required glyph."""
    upm = font["head"].unitsPerEm
    os2, hhea = font["OS/2"], font["hhea"]
    before = (os2.usWinAscent + os2.usWinDescent + hhea.lineGap) / upm

    total = round(TARGET_LINE_EM * upm)
    span = os2.sTypoAscender - os2.sTypoDescender          # keep the font's own balance
    asc = round(os2.sTypoAscender * total / span)
    desc = asc - total                                     # negative

    os2.sTypoAscender, os2.sTypoDescender, os2.sTypoLineGap = asc, desc, 0
    hhea.ascent, hhea.descent, hhea.lineGap = asc, desc, 0
    os2.fsSelection |= (1 << 7)                            # USE_TYPO_METRICS

    # win metrics bound the ink so nothing clips, and so apps that ignore
    # USE_TYPO_METRICS still land close to Tahoma rather than 31% over.
    hi, lo = ink_bounds(font)
    os2.usWinAscent, os2.usWinDescent = hi, abs(lo)

    typo_em = (asc - desc) / upm
    win_em = (os2.usWinAscent + os2.usWinDescent) / upm
    log(f"[6/8] vertical metrics: {before:.4f} em -> typo {typo_em:.4f} em / win {win_em:.4f} em"
        f"  (Tahoma {TARGET_LINE_EM:.4f})")
    if hi > asc:
        log(f"      note: tallest ink {hi} exceeds ascender {asc} by {hi-asc} units "
            f"({(hi-asc)/upm:.3f} em) — same as Tahoma, accents sit slightly proud")


# The OS/2 unicode-range flags say which Unicode blocks the font claims. WORD BUILDS THE
# "Subset" LIST IN Insert > Symbol FROM THESE FLAGS, not from the character map, so a block
# whose flag is clear cannot be BROWSED even though every character in it is present and
# reachable by typing its code point. Measured 8/24/2026: 16 flags were clear and 2,807 of
# 6,450 characters - 43.5% - could not be reached through that list.
#
# WHY NOT JUST recalcUnicodeRanges AND BE DONE, which is what the documentation used to
# suggest: it sets a flag for ANY character in a block, however few, and that makes the font
# claim scripts it cannot actually set. Measured 9/6/2026 on the shipping Regular:
#
#     Georgian                     1 character  of 96
#     Control Pictures             2            of 64
#     CJK Symbols And Punctuation  2            of 64
#     Halfwidth And Fullwidth      2            of 240
#     Arabic                      62            of 256
#
# The first four are a token presence. Arabic is the dangerous one: 62 characters is enough
# to look like coverage, and Arabic is a JOINING script - the font has no 'arab' entry in
# GSUB or GPOS at all (its script tags are DFLT, cyrl, dev2, deva, grek, latn), so Windows
# would be told this face sets Arabic and it would come out as unjoined isolated forms.
# Devanagari IS claimed, and that is not an inconsistency: the block is complete, 128 of 128,
# and the shaping tables are there.
#
# So the flags are set from the character map and then the refused ones are cleared, and a
# block that is neither claimed nor refused STOPS THE BUILD. That is deliberate. Changing a
# donor changes what is in the font, and the failure this whole typeface exists to prevent is
# a character silently coming from somewhere else - the same reasoning as REQUIRED above.
# A new block is a decision for a person, not something to wave through.

# Blocks the font covers well enough to claim. The count is what was measured on 9/6/2026;
# it is a note to the next reader, not a test.
CLAIMED = {
    37: "Arrows (330)",
    38: "Mathematical Operators (688)",
    39: "Miscellaneous Technical (214)",
    42: "Enclosed Alphanumerics (160)",
    43: "Box Drawing (10)",
    44: "Block Elements (8)",
    46: "Miscellaneous Symbols (161)",
    47: "Dingbats (43)",
    57: "Non-Plane 0 - the math alphanumerics live above U+FFFF",
    89: "Mathematical Alphanumeric Symbols (996)",
    15: "Devanagari (128 of 128, and deva/dev2 shaping is present)",
}

# Blocks present in the character map that must NOT be claimed, and why.
REFUSED = {
    13: "Arabic: 62 of 256 and NO 'arab' shaping - it would render unjoined",
    26: "Georgian: 1 character",
    40: "Control Pictures: 2 characters",
    48: "CJK Symbols And Punctuation: 2 characters",
    68: "Halfwidth And Fullwidth Forms: 2 characters",
}


WORDS = ("ulUnicodeRange1", "ulUnicodeRange2", "ulUnicodeRange3", "ulUnicodeRange4")


def _range_bits(os2) -> set[int]:
    return {i * 32 + b
            for i, w in enumerate(WORDS)
            for b in range(32) if getattr(os2, w) & (1 << b)}


def set_unicode_ranges(font: TTFont) -> None:
    """Claim the blocks the font really covers, so Word's Subset list can browse them."""
    os2 = font["OS/2"]

    # Only the blocks the MERGES added are ours to decide. Whatever the donor already claimed
    # is Noto's own call about its own coverage, and nothing here has taken characters away.
    before = _range_bits(os2)
    os2.recalcUnicodeRanges(font, pruneOnly=False)
    added = _range_bits(os2) - before

    unknown = sorted(added - set(CLAIMED) - set(REFUSED))
    if unknown:
        names = "; ".join(f"bit {b} "
                          + ", ".join(n for n, _ in OS2_UNICODE_RANGES[b])
                          for b in unknown)
        raise SystemExit(
            f"FAIL: the character map reaches {len(unknown)} Unicode block(s) that this "
            f"script neither claims nor refuses -> {names}. Count the characters in each "
            f"and add it to CLAIMED or REFUSED with the reason. See the note above CLAIMED.")

    for bit in REFUSED:
        i, b = divmod(bit, 32)
        setattr(os2, WORDS[i], getattr(os2, WORDS[i]) & ~(1 << b))

    kept = sorted(added & set(CLAIMED))
    log(f"[7/8] unicode ranges: claimed {len(kept)} newly browsable block(s), "
        f"refused {len(set(REFUSED) & added)} the font only touches")


def rename(font: TTFont, subfamily: str) -> None:
    full = FAMILY if subfamily == "Regular" else f"{FAMILY} {subfamily}"
    ps = f"{PSPREFIX}-{subfamily.replace(' ', '')}"
    for rec in font["name"].names:
        if   rec.nameID == 1:  rec.string = FAMILY
        elif rec.nameID == 2:  rec.string = subfamily
        elif rec.nameID == 3:  rec.string = f"{ps};VistaTypeLP"
        elif rec.nameID == 4:  rec.string = full
        elif rec.nameID == 6:  rec.string = ps
        elif rec.nameID == 16: rec.string = FAMILY
        elif rec.nameID == 17: rec.string = subfamily


def verify(path: str, subfamily: str, italic: bool = False) -> bool:
    f = TTFont(path, lazy=True)
    upm = f["head"].unitsPerEm
    cm = f.getBestCmap()
    hm = f["hmtx"]
    os2, hhea = f["OS/2"], f["hhea"]
    digits = {hm[cm[ord(d)]][0] for d in "0123456789"}
    greek = sum(1 for c in "αβγδεζηθικλμνξοπρσςτυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ"
                if ord(c) in cm)
    missing = [c for c in REQUIRED if ord(c) not in cm]
    checks = [
        ("TrueType outlines (glyf)",     "glyf" in f),
        ("static, no fvar",              "fvar" not in f),
        ("unitsPerEm == 960",            upm == 960),
        ("U+0030 is a slashed zero",     cm[0x30].endswith(".slash")),
        ("digits tabular",               len(digits) == 1),
        ("fsType 0 (embeddable)",        os2.fsType == 0),
        ("USE_TYPO_METRICS set",         bool(os2.fsSelection & (1 << 7))),
        ("Greek complete (49/49)",       greek == 49),
        ("U+025E present",               0x25E in cm),
        ("no required glyph missing",    not missing),
        ("line ~= Tahoma",               abs((hhea.ascent - hhea.descent + hhea.lineGap) / upm
                                             - TARGET_LINE_EM) < 0.005),
        ("win metrics cover ink",        os2.usWinAscent >= max(ink_bounds(f)[0], 0)),
        ("italic bits agree",            bool(os2.fsSelection & 0x01) == italic
                                         and bool(f["head"].macStyle & 2) == italic),
        # REGULAR means neither bold nor italic. Checked because the italic test above passes
        # happily on a face that claims to be regular AND italic at once - which shipped.
        ("regular bit only when plain",  bool(os2.fsSelection & 0x40)
                                         == (not italic and "Bold" not in subfamily)),
        # Word builds Insert > Symbol's Subset list from these flags, not from the character
        # map, so a cleared one hides a whole block from browsing. 43.5% of the font was
        # unreachable that way until 9/6/2026.
        ("symbol blocks browsable",      _range_bits(os2) >= set(CLAIMED)),
        ("Arabic NOT claimed",           13 not in _range_bits(os2)),
    ]
    ok = all(c[1] for c in checks)
    log(f"\n  verify {os.path.basename(path)} [{subfamily}]")
    for name, good in checks:
        log(f"    {'PASS' if good else 'FAIL'}  {name}")
    if missing:
        log(f"    missing: {' '.join(missing[:24])}")
    f.close()
    return ok


def main() -> int:
    ap = argparse.ArgumentParser(description="Build VistaTypeLP Sans from upstream Noto Sans.")
    ap.add_argument("--out", default="dist/fonts", help="output directory")
    ap.add_argument("--offline", metavar="DIR",
                    help="directory holding NotoSans-source.ttf and NotoSans-Italic-source.ttf "
                         "instead of fetching them")
    ap.add_argument("--no-verify", action="store_true")
    a = ap.parse_args()

    os.makedirs(a.out, exist_ok=True)
    sources = {w: fetch(w, a.offline, a.out) for w in sorted({i[2] for i in INSTANCES})}

    all_ok = True
    for subfamily, wght, which, italic in INSTANCES:
        log(f"\n=== {FAMILY} {subfamily} ===")
        font = TTFont(sources[which])
        slash_zero(font)

        log(f"[3/8] instancing {subfamily} (wght {wght}, wdth 100)")
        font = instancer.instantiateVariableFont(
            font, {"wght": wght, "wdth": 100}, inplace=False, updateFontNames=False)

        upm = font["head"].unitsPerEm
        if upm * RULER_DEN % RULER_NUM:
            raise SystemExit(f"FAIL: upm {upm} will not scale exactly by "
                             f"{RULER_NUM}/{RULER_DEN}; refusing to round")
        font["head"].unitsPerEm = upm * RULER_DEN // RULER_NUM
        log(f"[4/8] ruler scale: upm {upm} -> {font['head'].unitsPerEm} "
            f"({RULER_NUM}/{RULER_DEN} = {RULER_NUM/RULER_DEN:.6f}x)")

        # Every face gets the symbols, not just Regular. Word does NOT fall back inside a
        # family: tested on vistabuild 8/22/2026 with a Regular that had U+2264 and a Bold that
        # did not, and the bold one came out of CAMBRIA MATH -- another typeface at another
        # size, which is the exact failure this whole font exists to prevent. A heading is bold.
        for which, label in (("math", "Noto Sans Math"), ("symbols", "Noto Sans Symbols")):
            rep = merge_glyphs(font, fetch_donor(which, a.offline, a.out), label)
            log(f"[5/8] {label}: added {rep['added']} characters"
                + (f", skipped {rep['skipped']}" if rep["skipped"] else ""))

        retune_metrics(font)
        set_unicode_ranges(font)

        bold = "Bold" in subfamily
        os2 = font["OS/2"]
        os2.usWeightClass = wght
        # bits 0 italic, 5 bold, 6 regular. REGULAR means "neither bold NOR italic" - it is not
        # the opposite of bold. The line here read `0x20 if bold else 0x40`, which handed the
        # regular bit to the Italic face as well as its italic bit, and fontTools warned about
        # it on every save from 8/22/2026. Nothing decided that; it fell out of the way the line
        # was written, and verify's "italic bits agree" looked only at the italic bit so it never
        # caught it. What it risks is style matching: an application choosing a family member
        # from these flags can read that face as a plain one. Found 9/6/2026 while setting the
        # unicode ranges; fixed on Jerry's say-so the same day.
        os2.fsSelection &= ~0x61
        os2.fsSelection |= ((0x20 if bold else 0)
                            | (0x01 if italic else 0)
                            | (0x40 if not (bold or italic) else 0))
        font["head"].macStyle = (1 if bold else 0) | (2 if italic else 0)
        rename(font, subfamily)

        out = os.path.join(a.out, f"{PSPREFIX}-{subfamily.replace(' ', '')}.ttf")
        font.save(out)
        font.close()
        log(f"[8/8] wrote {out}  ({os.path.getsize(out)//1024} KB)")

        if not a.no_verify:
            all_ok &= verify(out, subfamily, italic)

    log("\nBUILD OK" if all_ok else "\nBUILD FAILED VERIFICATION")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())

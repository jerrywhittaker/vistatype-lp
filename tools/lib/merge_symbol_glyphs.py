#!/usr/bin/env python3
"""
merge_symbol_glyphs.py -- copy characters from a donor font into a VistaTypeLP face.

Used by build_vistatypelp_sans.py to fold Noto Sans Math and Noto Sans Symbols into
VistaTypeLP Sans, so a book can set mathematics, logic, astronomy and music without Word
substituting the character out of some other typeface at some other size.

Why this exists rather than fontTools' own merger: that merger fails outright on these two
donors. It cannot cope with Noto Sans Math's MATH table, and even with every layout table
stripped it raises deep in its own code on both of them. So this copies glyphs one at a
time, which is also the safer operation -- nothing of ours can be overwritten, because a
character we already carry is never copied.

Two things it takes care of that a naive copy does not:

  * Components. A donor glyph may be built from pieces of other glyphs, and those pieces are
    named and numbered in the DONOR. Copying such a glyph by reference would point it at
    whatever happens to sit at that number in OUR font, which draws recognizable garbage.
    Every glyph is therefore flattened to plain outlines on the way across.

  * Units per em. Our faces are 960 to the em (the 25/24 ruler rescaling); the donors are
    1000. Their units-per-em is set to 960 without touching a coordinate, exactly as the
    ruler rescaling does, so their symbols arrive at the same scale as our letters with
    nothing rounded.
"""
from __future__ import annotations

from fontTools.ttLib import TTFont
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.ttGlyphPen import TTGlyphPen

TARGET_UPM = 960

# Layout and colour tables the donors bring that we neither want nor can merge. Their
# typographic rules are not ours; we are taking characters.
DONOR_DROP = (
    "MATH", "SVG ", "COLR", "CPAL", "DSIG",
    "fvar", "STAT", "avar", "HVAR", "VVAR", "MVAR", "gvar",
    "GSUB", "GPOS", "GDEF", "kern",
)


def _prepare(path: str) -> TTFont:
    """Open a donor, pin any variable axes to Regular, drop what we cannot use, rescale."""
    font = TTFont(path)
    if "fvar" in font:
        from fontTools.varLib import instancer
        axes = {a.axisTag: (400 if a.axisTag == "wght" else a.defaultValue)
                for a in font["fvar"].axes}
        font = instancer.instantiateVariableFont(font, axes, inplace=False,
                                                 updateFontNames=False)
    for tag in DONOR_DROP:
        if tag in font:
            del font[tag]

    upm = font["head"].unitsPerEm
    if upm != TARGET_UPM:
        if upm != 1000:
            raise SystemExit(f"donor {path} is {upm} units per em; only 1000 can be "
                             f"rescaled to {TARGET_UPM} without rounding coordinates")
        font["head"].unitsPerEm = TARGET_UPM
    return font


def _wide_subtable(target: TTFont):
    """Return a character map able to hold codepoints above U+FFFF, adding one if needed.

    The mathematical alphabets -- the double-struck, script and fraktur letters a maths book
    sets -- all live above U+FFFF, and the ordinary character map format tops out there. It
    is stored as an unsigned short, so writing one into it fails on save with an OverflowError
    that names no character. A second, wider map is the standard way to carry them, and it is
    built by copying what the ordinary one already holds so the two never disagree.
    """
    from fontTools.ttLib.tables._c_m_a_p import CmapSubtable

    for table in target["cmap"].tables:
        if table.isUnicode() and table.format in (12, 13):
            return table

    base = {}
    for table in target["cmap"].tables:
        if table.isUnicode():
            base.update(table.cmap)

    wide = CmapSubtable.newSubtable(12)
    wide.platformID, wide.platEncID, wide.format = 3, 10, 12
    wide.reserved, wide.length, wide.language, wide.nGroups = 0, 0, 0, 0
    wide.cmap = dict(base)
    target["cmap"].tables.append(wide)
    return wide


def _map_codepoint(target: TTFont, cp: int, name: str) -> None:
    """Point cp at name in every character map that can express it."""
    if cp > 0xFFFF:
        _wide_subtable(target).cmap[cp] = name
        return
    for table in target["cmap"].tables:
        if table.isUnicode():
            table.cmap[cp] = name


def merge_glyphs(target: TTFont, donor_path: str, label: str) -> dict:
    """Copy every character the donor has and the target lacks. Returns a small report."""
    donor = _prepare(donor_path)

    target_cmap = target.getBestCmap()
    donor_cmap = donor.getBestCmap()
    wanted = sorted(set(donor_cmap) - set(target_cmap))
    if not wanted:
        donor.close()
        return {"donor": label, "added": 0, "skipped": 0}

    donor_glyphs = donor.getGlyphSet()
    donor_hmtx = donor["hmtx"]
    target_glyf = target["glyf"]
    target_hmtx = target["hmtx"]
    existing = set(target.getGlyphOrder())

    added, skipped, new_names = 0, 0, []
    for cp in wanted:
        src_name = donor_cmap[cp]

        # A name of ours must never be reused - that would silently replace one of our own
        # letters with a donor's idea of it.
        name = src_name
        if name in existing:
            name = f"{src_name}.u{cp:04X}"
        if name in existing:
            skipped += 1
            continue

        # Flatten: draw the donor glyph, resolving any components into plain outlines.
        pen = DecomposingRecordingPen(donor_glyphs)
        try:
            donor_glyphs[src_name].draw(pen)
        except Exception:
            skipped += 1
            continue

        tt_pen = TTGlyphPen(None)
        pen.replay(tt_pen)
        target_glyf[name] = tt_pen.glyph()
        target_hmtx[name] = donor_hmtx[src_name]

        existing.add(name)
        new_names.append(name)
        added += 1

        _map_codepoint(target, cp, name)

    if new_names:
        # The outline table keeps its OWN glyph order and ALREADY appended each name as it was
        # inserted above. Appending them again to the font's copy makes the two disagree, and
        # the save then fails on a bare AssertionError that names neither table. So take the
        # outline table's order as the truth and put the font's copy in step with it.
        order = list(target["glyf"].glyphOrder)
        target.setGlyphOrder(order)
        target["maxp"].numGlyphs = len(order)

    donor.close()
    return {"donor": label, "added": added, "skipped": skipped}


def merge_all(target: TTFont, donors: list[tuple[str, str]]) -> list[dict]:
    """donors is [(path, label), ...], applied in order; earlier donors win a tie."""
    return [merge_glyphs(target, path, label) for path, label in donors]

# Atkinson Hyperlegible — imported, and rescaled for large print

Typeface by the **Braille Institute of America**, designed for low-vision readers — the same
audience this add-in serves.

**This ships.** From 3.0.101 the installer installs the four rescaled faces as
**VistaTypeLP Legible**, and the transcriber chooses between it and Tahoma on the attach
dialog. `make stage` copies `scaled/*.ttf` and `OFL.txt` into `dist/`, and
`installer/vistatype.iss` installs them.

## Where this came from

| | |
|---|---|
| Source archive | `Atkinson-Hyperlegible-Font-Print-and-Web-Final.zip` |
| Found at | `E:\Downloads\` on Jerry's Windows machine, 8 Aug 2026 |
| Archive SHA-256 | `9e1a65d53ddadc8253791e958a07bf5aba210ef6155ea9c91b2b4c46bbce53e3` |
| Release folder inside | `Atkinson-Hyperlegible-Font-Print-and-Web-2020-0514` |
| Font version | 1.02 (14 May 2020) |

## What is here

`upstream/` holds the files **exactly as they came out of the archive** — never edit them.

- `upstream/ttf/` — TrueType outlines. **These are the ones we rescale and ship.**
- `upstream/otf/` — the same four faces with PostScript outlines. Kept for reference only.

**It has to be the TTF, and this is not a preference.** Word's `EmbedTrueTypeFonts` means
what it says: it embeds TrueType outlines and skips PostScript ones **silently** — no error,
no warning, the document simply saves without the font. A book handed to a reader who hasn't
got the font installed would then set in a substitute at the wrong size, which is the exact
fault the rescale exists to cure, arriving at the last possible moment.

`scaled/` is **derived** from `upstream/ttf/` by `tools/lib/rescale_font.py` and is a build
output — regenerate it, never hand-edit it:

```bash
python3 tools/lib/rescale_font.py --factor 1.094 \
    --family "VistaTypeLP Legible" \
    --out assets/fonts/atkinson-hyperlegible/scaled \
    assets/fonts/atkinson-hyperlegible/upstream/ttf/*.ttf
```

`OFL.txt` is the license, and must travel with the font wherever it goes — see below.

Left in the archive as useless to a Word add-in: the EOT, SVG, WOFF and WOFF2 web formats.
The archive itself was not copied in; the checksum above identifies it if it is needed again.

## Why the font is rescaled at all

A point size in Word sets the **em** — the notional body the type sits on — not the height of
any letter, and every typeface decides for itself how much of that body the letters fill.
Atkinson fills less of it than Tahoma: 8.1% less cap height, 9.1% less x-height, 8.6% on a
real sentence. Jerry measured it with a font ruler on 8 Aug 2026 and read 18 pt large print as
about 16.5 pt.

That matters more here than in ordinary typesetting, because large print is *specified* in
points. A transcriber who sets 18 pt has to get 18 pt. `rescale_font.py` scales every outline
and metric by **1.094** while leaving the em at 1000 units, so the transcriber sets 18, Word
reports 18, and the ruler reads 18. Jerry confirmed the rescaled faces against the ruler on
8 Aug 2026. Letterforms are otherwise untouched; all 352 glyphs survive the round trip.

One trap worth knowing, because the first attempt fell into it: a font carries its names in
**two** places. Renaming only the `name` table leaves the CFF table's PostScript name intact —
and that is the name Word writes into a PDF. A large-print PDF would have reported
`AtkinsonHyperlegible-Regular`, which is both a license problem and a practical one: a print
shop's system resolves fonts by that name, so it could substitute the genuine Atkinson and put
the 9%-too-small face back on the paper, at the very last step, in front of the reader.
`rescale_font.py` now re-opens every file it writes and refuses one that still carries a
reserved word in any name a user or a tool can see.

### Line spacing, and why the script touches the vertical metrics

Upstream's OTF and TTF carry **identical** typographic and window metrics but disagree about
which of them applies: the OTF is OS/2 version 4 with `USE_TYPO_METRICS` set and a 150-unit
`hhea` line gap; the TTF is version 3, where that flag isn't even defined, and its line gap is
zero. Word honors `USE_TYPO_METRICS`, so the same design would set lines about **17% further
apart** as a TTF than as an OTF.

Switching format to make embedding possible must not change how the type sits on the page, so
the script bumps the TTF's OS/2 table to version 4 and pins its `hhea` values to the
typographic ones. That normalizes an upstream inconsistency; it doesn't invent a line height.

The result sets lines **8.5% taller than Tahoma** at the same point size, which is the font's
own design and is left alone. It's not a problem in practice: a book's typeface is chosen once
when the template is attached, and re-attaching at a different size or page size repaginates
anyway.

Measured against Tahoma at the same point size: cap height **1.005**, x-height **0.996**,
running text **1.065** wider, bullet glyph **1.054** wider. That last figure is where
`Lp_Indent_Factor_For_Font` gets its number.

## License — settled, with conditions we have to keep meeting

**This was an open question and is now closed.** The fonts' own 2020 name table still says the
typeface is offered *"without derivatives or alteration"*, which would have forbidden the
rescale. That statement is **superseded**. Jerry obtained the Braille Institute's license
document (prepared by their counsel, dated December 2024), which places the typeface under the
**SIL Open Font License, Version 1.1** under the copyright line *"Copyright 2020, Braille
Institute of America, Inc."* — the same 2020 that our 1.02 files carry.

`OFL.txt` in this folder is that license, transcribed from Jerry's PDF and verified against it
character for character. It reads very slightly differently from the generic SIL boilerplate —
the Braille Institute flattened the optional "(s)" plurals, being a single copyright holder —
so **this** wording is what we ship, not a copy fetched from elsewhere.

The PDF itself is deliberately **not** in the repository — it is Jerry's own copy of a third
party's legal document, and the repo needs the text, not the file. It is kept at
`~/reference/vistatype-lp/`, with a note there recording what it settled. `.gitignore` refuses
a file of that name, so a stray second copy dropped into the project cannot be committed.

The OFL permits the rescale, permits shipping the result in the installer, and permits
bundling it with software under a different license. It attaches five conditions. Four bear on
us, and here is how each is met — **re-check these whenever the font or the installer
changes**:

| Condition | What it requires | How it is met |
|---|---|---|
| 1 — not sold alone | The font may not be sold by itself | VistaType LP is free; the font only ever ships inside the installer |
| 2 — license travels with it | Every copy carries the copyright notice and the license | The font's own copyright/license records are rewritten to the OFL by `rescale_font.py` (in the CFF table too, when there is one), **and** `OFL.txt` is installed to `%AppData%\VistaType LP` |
| 3 — Reserved Font Names | A modified version may not use **ATKINSON** or **HYPERLEGIBLE** in the name shown to users | The family is renamed **VistaTypeLP Legible**; `rescale_font.py` refuses any `--family` containing either word, and re-opens each file it writes to prove the words are gone from every name a user or tool can see |
| 4 — no endorsement | The Braille Institute's name may not promote a modified version, only be acknowledged | The vendor record points at this project, not at them; the designer credit and the origin statement stay, as acknowledgment |
| 5 — stays under the OFL | The font must be distributed entirely under the OFL and no other license | The font is **not** covered by VistaType's GPLv3; the installer's welcome page says so, and `docs/Installation-Guide.md` explains it |

Two consequences worth stating outright:

- **The font is not GPL.** VistaType LP is GPLv3; this font is OFL, and the two simply travel
  together. The OFL expressly allows that. Do not let a future tidy-up sweep the font under
  the project license — condition 5 forbids it, and breaching any condition voids the grant.
- **"VistaTypeLP Legible" cannot casually be renamed back.** Anything with "Atkinson" or
  "Hyperlegible" in it breaches condition 3. (The name also avoids a clash with VistaType, an
  unrelated existing font producer — hence "VistaTypeLP", not "VistaType".)

## Embedding in documents, and what transcribers may do

`fsType = 0` on all four faces: no embedding restriction. A transcriber can embed the font in
a `.docx` and send it to someone who does not have it installed.

From 3.0.101 the add-in does this automatically: `Lp_Attach_The_Template` turns on
`EmbedTrueTypeFonts` when the transcriber chose VistaTypeLP Legible, and leaves it off for
Tahoma, which is on every Windows machine anyway. `SaveSubsetFonts` is deliberately **off** —
a subset holds only the characters the document already contains, so a teacher who edits the
book and types a new one would get a substitute for it.

Better still, the OFL says outright that **the requirement for fonts to remain under this
license does not apply to any document created using the font**. So a transcriber's finished
large-print document carries no license obligation of any kind — they can hand it to a reader,
a school or a publisher without a thought. That is the answer to the obvious worry, and it is
in `docs/Installation-Guide.md`.

## How the choice reaches the document

| Where | What it does |
|---|---|
| `LP_Attach_An_Lp_Template_Form` | The Typeface buttons, beside the point size. Legible is the default, grayed out when it isn't installed. Sets the public `Lp_Base_Font_Name` |
| `Lp_Apply_Base_Font_To_Styles` | Sets Normal — which almost everything follows — plus the 13 styles that name a face of their own: the 12 colored character styles and `No Spacing` |
| `Sh_Set_Whole_Document_Font` | Lays the face over the text as direct formatting, so a messy imported document actually conforms |
| `Lp_Get_Doc_Setup_Params` | Reads it back off the Normal style. Nothing stores the choice separately, so it survives closing and reopening, and a book made before this feature answers Tahoma correctly |
| `Lp_Indent_Factor_For_Font` | Scales the hand-tuned bulleted-list indents by 1.054 for Legible, 1.0 for Tahoma |

**`LargePrintTemplate.dotx` is deliberately not touched.** It stays the Tahoma baseline — its
`docDefaults` still names Tahoma — and the macros override per document. That is exactly what
makes a book produced before this existed read back as Tahoma without any migration.


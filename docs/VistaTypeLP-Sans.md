# VistaTypeLP Sans

The typeface VistaType LP bundles and installs, from build 3.0.224 onward.
Every figure in this document was measured from the shipping font files on
8/24/2026, against Tahoma 7.05 as it comes on Windows 11. Nothing here is
estimated.

---

## What it is, in one paragraph

VistaTypeLP Sans is Google's **Noto Sans** with three changes made for large
print: a **slashed zero as the ordinary zero**, an **exact 25/24 enlargement** so
a nominal point size matches the VistaType ruler, and the whole of **Noto Sans
Math** and **Noto Sans Symbols** folded into all four faces. It carries **6,450
characters** where Tahoma carries 3,772. It is released under the SIL Open Font
License, so VistaType LP may bundle it, embed it in a book, and hand it to a
transcriber without asking anyone.

---

## The four faces

| Face | File | Characters | Glyphs |
|---|---|---|---|
| Regular | `VistaTypeLPSans-Regular.ttf` | 6,450 | 7,871 |
| Bold | `VistaTypeLPSans-Bold.ttf` | 6,450 | 7,871 |
| Italic | `VistaTypeLPSans-Italic.ttf` | 6,448 | 7,883 |
| Bold Italic | `VistaTypeLPSans-BoldItalic.ttf` | 6,448 | 7,883 |

All four report version 2.015 and all four carry the *same* symbol set. That is
deliberate and it is not free — the math and symbol characters had to be merged
into each face separately. Word does **not** fall back inside a family: when a
Bold face lacks a character, Word does not quietly reach for the Regular one, it
goes to a different typeface entirely at a different size. That was tested, and a
bold character came back out of Cambria Math.

The Italic is short two characters against the Regular — U+10FB (a Georgian
paragraph separator) and U+20C0 (a proposed currency sign). Neither matters.

**Tahoma has only two faces.** Windows ships `tahoma.ttf` and `tahomabd.ttf` and
nothing else. There is no italic Tahoma and there never was — what a transcriber
sees when she presses Ctrl+I in a Tahoma book is Word slanting the roman by
machine. VistaTypeLP Sans has four drawn faces, so italic and bold italic are
real.

---

## What was changed from stock Noto Sans

Seven steps, all of them repeatable from an upstream release by
`tools/lib/build_vistatypelp_sans.py`.

| Step | What it does | Why |
|---|---|---|
| fetch | Pulls the Noto Sans variable fonts from Google | The source of truth is upstream, not a file in this repo |
| **slash** | Repoints U+0030 to the `zero.slash` glyph | Word cannot reach the OpenType `zero` feature, so the slashed zero has to *be* the ordinary zero. Confirmed: U+0030 maps to `zero.slash` |
| instance | Cuts static Regular (400) and Bold (700) | Word handles variable fonts poorly and will not give a real Bold from one |
| **scale** | Units-per-em 1000 → 960, an exact 25/24 enlargement | Measured: 37.5 point matched the VistaType ruler's 36. Rescaling makes the nominal size honest |
| **symbols** | Folds Noto Sans Math and Noto Sans Symbols into every face | See the note above about Word not falling back inside a family |
| metrics | Retunes vertical metrics to 1.20729 em | So Word's "Single" line spacing lands on Tahoma's 1.20703 em. The two differ by 0.0003 em — about a quarter of a thousandth of a line |
| rename | Renames per the OFL Reserved Font Name clause | Required by the license once the font has been modified |

---

## Languages

Anything a **Latin-script** language needs is there. Both typefaces cover the
whole of Latin-1, Latin Extended-A and B, Latin Extended Additional, and the
Vietnamese range; VistaTypeLP Sans adds Latin Extended-F and G on top (88 more
characters Tahoma has not got).

| Language | VistaTypeLP Sans | Tahoma |
|---|---|---|
| English, Indonesian, Malay, Swahili | yes | yes |
| Spanish, French, German, Italian, Portuguese, Dutch | yes | yes |
| Danish, Norwegian, Swedish, Finnish, Icelandic | yes | yes |
| Polish, Czech, Slovak, Hungarian, Romanian | yes | yes |
| Croatian, Slovene, Serbian (Latin), Latvian, Lithuanian, Estonian | yes | yes |
| Turkish, Azerbaijani, Maltese, Welsh, Irish, Esperanto | yes | yes |
| **Vietnamese** (all tone marks) | yes | yes |
| Filipino, Hawaiian, Navajo, Yoruba, Hausa, pan-Nigerian | yes | yes |
| **Greek** — modern, accented and polytonic | yes | yes |
| **Russian, Ukrainian, Bulgarian, Serbian (Cyrillic), Macedonian** | yes | yes |
| Kazakh, Mongolian (Cyrillic) | yes | yes |
| **Hindi, Marathi, Nepali** (Devanagari) | **yes** | no |
| Hebrew | **no** | yes |
| Yiddish | **no** | yes |
| Arabic, Persian, Urdu | **no** — see below | yes |
| Thai | **no** | yes |
| Armenian | **no** | yes |
| Georgian, Bengali, Tamil, Thaana, Khmer | no | no |
| Chinese, Japanese, Korean | no | no |
| Cherokee, Inuktitut | no | no |

**Arabic needs a footnote, because a bare character count lies about it.**
VistaTypeLP Sans holds 62 Arabic letters, but Arabic letters change shape
depending on what sits either side of them, and that shaping is done by
OpenType rules the font has to carry. VistaTypeLP Sans has no `arab` script
in its layout tables at all — no initial, medial, final or isolated forms —
so Arabic would come out as a row of disconnected letters. Treat it as not
supported. Tahoma does it properly: the `arab` script, all four positional
features, and 319 presentation forms besides.

**Devanagari is the reverse case.** VistaTypeLP Sans carries 128 Devanagari
characters *and* the full set of shaping rules for them — conjuncts, half
forms, the reph, below-base forms. Hindi, Marathi and Nepali genuinely set.
Tahoma has no Devanagari whatever.

The scripts VistaTypeLP Sans can actually lay out are: **Latin, Greek,
Cyrillic and Devanagari.** Tahoma's are **Latin, Greek, Cyrillic, Hebrew,
Arabic and Thai.**

---

## Mathematics, science and medicine

This is where the two part company, and it is the reason VistaTypeLP Sans
exists. The counts are characters found out of characters tested.

| Area | VistaTypeLP Sans | Tahoma |
|---|---|---|
| Arithmetic — `+ − × ÷ = ≠ ≈ ≡ ± ∓ ≤ ≥ ∞ % ‰ ° ′ ″` | 22/22 | 19/22 |
| Algebra and calculus — `√ ∛ ∜ ∑ ∏ ∫ ∬ ∭ ∮ ∂ ∇ ∆ ∝ ∴ ∵` | 20/20 | 11/20 |
| Set theory and logic — `∀ ∃ ∈ ∉ ⊂ ⊆ ∪ ∩ ∅ ¬ ∧ ∨ ⇒ ⇔ ⊕ ⊗ ⊥` | **27/27** | **2/27** |
| Relations and geometry — `∠ ∡ ∢ ≅ ≌ ∼ ≃ △ ▽ ⊿` | 21/21 | 8/21 |
| Arrows — `← → ↑ ↓ ↔ ⇐ ⇒ ⇑ ⇓ ⇔ ⇌ ⇋ ↦ ⟶ ↺ ↻` | **23/23** | **6/23** |
| Superscript digits and signs | 17/17 | 17/17 |
| Subscript digits and letters | 27/27 | 27/27 |
| Vulgar fractions — `½ ⅓ ⅔ ¼ ¾ ⅕ ⅙ ⅛ ⅜ ⅝ ⅞` | 20/20 | 20/20 |
| Letterlike and blackboard — `ℝ ℕ ℤ ℚ ℂ ℍ ℵ ℏ ℓ ℘ ℑ ℜ №` | 29/29 | 29/29 |
| **Math italic letters** `𝑎 𝑏 𝑐 𝐴 𝐵 𝐶` | **12/12** | **0/12** |
| **Math bold letters** `𝐚 𝐛 𝐜 𝟎 𝟏 𝟐` | **9/9** | **0/9** |
| **Math script and fraktur** `𝒜 𝒞 𝔄 𝔅 𝕬 𝔸` | **11/11** | **1/11** |
| Chemistry — `⇌ ⌬ ⚛ ⚗ µ Å ℓ ‰` | 17/17 | 10/17 |
| Physics and SI — `Ω µ Å ° ℃ ℉ ℧ ℏ ∂ ∇` | 16/16 | 14/16 |
| **Astronomy and planets** — `☉ ☽ ☿ ♀ ♁ ♂ ♃ ♄ ♅ ♆ ♇ ★` | **16/16** | **2/16** |
| **Zodiac** — `♈ ♉ ♊ ♋ ♌ ♍ ♎ ♏ ♐ ♑ ♒ ♓` | **12/12** | **0/12** |
| Medicine and pharmacy — `℞ ☤ ⚕ ⚚ ℈ ℥ ℔ ʒ † ☩` | 13/14 | 8/14 |
| **Genealogy** — `♀ ♂ ⚥ ⚭ ⚮ ⚯ ✝ †` | **12/12** | **7/12** |
| Roman numerals `Ⅰ Ⅱ Ⅲ Ⅳ Ⅴ Ⅹ Ⅼ Ⅽ Ⅾ Ⅿ` | 19/19 | 19/19 |
| Currency, including `€ ₹ ₽ ₿` | 22/22 | 22/22 |
| Editorial — `† ‡ § ¶ • ‰ ′ ″ « » „ … ※ ‽ ⁂` | 22/22 | 22/22 |
| **IPA** — 96 IPA Extensions, 128 Phonetic Extensions, 64 Supplement | 50/50 | 50/50 |
| **Circled numbers** `① ② ③ ⑩ ⑴ ⒜ Ⓐ ⓐ` | **17/17** | **0/17** |
| **Dice** `⚀ ⚁ ⚂ ⚃ ⚄ ⚅` | **6/6** | **0/6** |
| Playing card suits `♠ ♡ ♢ ♣ ♥ ♦` | 8/8 | 4/8 |
| **Alchemical** `🜁 🜂 🜃 🜄 🜍 🜔 🜚` | **116 characters** | **0** |
| Geometric shapes `■ □ ▲ △ ▼ ▽ ◆ ◇ ○ ● ◐` | 19/19 | 10/19 |

The blunt version, block by block:

| Unicode block | VistaTypeLP Sans | Tahoma |
|---|---|---|
| Mathematical Operators | 256 | 18 |
| Supplemental Mathematical Operators | 256 | 0 |
| Miscellaneous Mathematical Symbols-A and B | 176 | 0 |
| **Mathematical Alphanumeric Symbols** | **996** | **0** |
| Arrows + Supplemental Arrows-A and B | 256 | 7 |
| Miscellaneous Technical | 214 | 4 |
| Miscellaneous Symbols | 161 | 12 |
| Enclosed Alphanumerics + Supplement | 317 | 0 |
| Dingbats | 43 | 0 |
| Alchemical Symbols | 116 | 0 |
| Geometric Shapes + Extended | 185 | 16 |
| Devanagari | 129 | 0 |
| Cyrillic Extended-A, B, C | 137 | 0 |
| Hebrew | 0 | 88 |
| Arabic (all blocks) | 63 | 696 |
| Thai | 0 | 87 |
| Armenian | 0 | 91 |
| Box Drawing | 10 | 40 |
| **Total mapped characters** | **6,450** | **3,772** |

---

## What VistaTypeLP Sans does *not* have

Measured, not assumed. In every case Word will substitute another typeface at
another size, silently — which for a large print book is the one thing that must
not happen. **Check before setting a book that needs any of these.**

| Missing | Detail |
|---|---|
| **Chess pieces** `♔ ♕ ♖ ♗ ♘ ♙ ♚ ♛ ♜ ♝ ♞ ♟` | U+2654–265F. All twelve absent. The Chess Symbols block at U+1FA00 is absent too. **Tahoma has not got them either** — no chess book sets in either typeface without a substitution |
| **Braille patterns** `⠁ ⠂ ⠃ ⡿ ⣿` | The whole of U+2800–28FF, all 256, absent. Tahoma likewise. Nothing here sets braille cells as print characters |
| **Musical staff notation** `𝄞 𝄢 𝄡` | The Musical Symbols block at U+1D100 is absent. The four accidentals and note heads that live in Miscellaneous Symbols — `♩ ♪ ♫ ♬ ♭ ♮ ♯` — *are* there |
| **Dingbat checks and crosses** `✔ ✗ ✘ ✕ ✖ ☑ ☒ ☐ ✂ ✈` | Only `✓` U+2713 is present. Tahoma has none of them at all |
| **Weather and warning** `☀ ☁ ☂ ☃ ⚡ ⚠ ☎ ☏` | Absent, along with most of the pictorial half of Miscellaneous Symbols |
| **Coptic** `Ⲁ ⲁ Ϣ ϣ Ϥ ϥ` | Absent. Tahoma carries 14 of the Coptic letters. Modern and polytonic **Greek are complete in both** — this affects Coptic only |
| **Box drawing** `│ ├ ┤ ┬ ┴ ┼ ║ ╔ ╗ ╚ ╝` | 10 of 128 present. Tahoma has 40. Neither will draw a full box |
| **Emoji** | Effectively none. Five characters from Miscellaneous Symbols and Pictographs, one Emoticon |
| **Hebrew, Yiddish, Arabic, Persian, Urdu, Thai, Armenian** | See the languages table above |
| **Chinese, Japanese, Korean** | Absent, as in Tahoma |

---

## VistaTypeLP Sans against Tahoma, side by side

| | VistaTypeLP Sans | Tahoma |
|---|---|---|
| Faces drawn | **4** — Regular, Bold, Italic, Bold Italic | 2 — Regular, Bold. Word fakes the italic |
| Characters | **6,450** | 3,772 |
| Glyphs | 7,871 | 4,498 |
| Zero | **Slashed by default** — no setting to switch on | Unslashed. Nothing will slash it |
| x-height | 0.5583 em | 0.5454 em — **2.3% smaller** |
| Cap height | 0.7438 em | 0.7271 em — 2.2% smaller |
| Line height at "Single" | 1.20729 em | 1.20703 em — the same to within 0.0003 |
| Set width, running text | **9.4% wider** | narrower |
| Nominal point size | Matches the VistaType ruler (the 25/24 rescale) | Runs small against it |
| Layout scripts | Latin, Greek, Cyrillic, **Devanagari** | Latin, Greek, Cyrillic, **Hebrew, Arabic, Thai** |
| Mathematics | **Complete** — 1,684 math characters | 18 operators and nothing else |
| Typographic features | fractions, small caps, old-style and tabular figures, superiors and inferiors, four stylistic sets | kerning and ligatures only |
| Embedding | **fsType 0** — installable, no restriction | fsType 8 — editable embedding |
| License | SIL Open Font License | Microsoft proprietary, licensed with Windows |
| On the machine already? | No — VistaType LP installs it | Yes, on every Windows since 1999 |

### The width difference is the one that costs something

Measured at 18 point on two sample sentences:

| Text | VistaTypeLP Sans | Tahoma | Difference |
|---|---|---|---|
| "The quick brown fox jumps over the lazy dog. Now is the time for all good men to come to the aid of their party." | 13.637 in | 12.459 in | **+9.5%** |
| "In 1994 the patient presented with acute pharyngitis; temperature 38.5 C, pulse 96." | 10.161 in | 9.295 in | **+9.3%** |

At the same point size on the same page, VistaTypeLP Sans fits about **nine
percent less text on a line**. Across a book that is more line breaks, more
pages, and different pagination. It is not a defect — the face is a little
larger on the body, which is the point — but a book reset from Tahoma into
VistaTypeLP Sans at the same point size **will repaginate**, and a book already
in a reader's hands must not be reset. That is exactly why
`Lp_Attach_The_Template` protects a book already set in the old face rather than
carrying it forward.

### Where Tahoma still wins

Three places, and they are real:

1. **Hebrew, Arabic and Thai.** Tahoma sets all three properly, VistaTypeLP Sans
   sets none of them. A transcriber working on a text with Hebrew or Arabic in it
   needs Tahoma or a third typeface for those runs.
2. **It is already there.** Tahoma is on every Windows machine. VistaTypeLP Sans
   has to be installed, and if a book set in it travels to a machine without it,
   the reader gets a substitution — which is why `Lp_Attach_The_Template` embeds
   the face, unsubsetted, in every book set in it.
3. **The underline.** In Tahoma the underline and the underscore are drawn as the
   same bar, so a row of underscores makes one unbroken rule. In VistaTypeLP Sans
   the underline is thinner and sits inside the underscore, leaving a 0.19-point
   sliver unpainted at 18 point — the row of underscores comes out with holes in
   it. That is why fill-in lines are typed in Tahoma whatever face the book is
   set in, and why `Lp_Tahoma_The_Fill_Ins` puts them back after an attach.

---

## Known limitation — Word's Insert Symbol does not show it all

Open Insert → Symbol → More Symbols in Word, pick VistaTypeLP Sans, and the
**Subset** list is missing whole ranges that the font genuinely has. Arrows,
mathematical operators, the math italic and bold letters, miscellaneous
technical, miscellaneous symbols, dingbats, circled numbers, box drawing and
block elements are all absent from that list.

**What was measured:** the font's OS/2 table carries a set of flags naming which
Unicode ranges it covers, and VistaTypeLP Sans has **16 of those flags unset**
even though it has characters in every one of the ranges they name. Those 16
ranges are exactly the ones missing from Word's Subset list — Word builds that
list from the flags rather than from the font's actual character map.

**2,807 of its 6,450 characters — 43.5% — cannot be reached through the Subset
list.** Tahoma has no unset flags, which is why Tahoma browses correctly.

The characters themselves are all there and all reachable. Two ways to get at
one that the Subset list will not show:

- Type the hexadecimal code point and press **Alt+X** — `21D2` then Alt+X gives `⇒`.
- In Insert → Symbol, set **from:** to Unicode (hex) and type the code into the
  **Character code** box; the character is found even when its subset is not listed.

**This is a fixable defect in the font, not in Word.** Setting those 16 flags
correctly in `tools/lib/build_vistatypelp_sans.py` — fontTools will compute them
from the character map in one call — would make the whole set browsable. It has
not been done yet.

---

## Rebuilding it

```
python3 tools/lib/build_vistatypelp_sans.py --out assets/fonts/vistatypelp-sans
```

It fetches the current Noto Sans, Noto Sans Math and Noto Sans Symbols from
Google, applies the seven steps in the table above, and checks that every face
can set the required character list — Latin, Greek, the IPA and the math
operators — before writing anything. A face that cannot set them fails the build
rather than reaching a transcriber.

`--offline SRC` uses previously downloaded sources in `dist/fonts/` instead of
fetching.

**Test character coverage before changing donors.** The face that shipped from
3.0.101 to 3.0.196 was dropped because it had no Greek, no IPA and almost no
mathematics, and nothing had checked. The required-character list in the build
script is what stops that happening twice.

---

## Files

| Path | What |
|---|---|
| `assets/fonts/vistatypelp-sans/VistaTypeLPSans-*.ttf` | The four faces, tracked |
| `assets/fonts/vistatypelp-sans/OFL.txt` | Noto Sans license |
| `assets/fonts/vistatypelp-sans/OFL-NotoSansMath.txt` | Noto Sans Math license |
| `assets/fonts/vistatypelp-sans/OFL-NotoSansSymbols.txt` | Noto Sans Symbols license |
| `tools/lib/build_vistatypelp_sans.py` | The build |
| `tools/lib/merge_symbol_glyphs.py` | Folds the math and symbol donors into a face |

All three license files are verbatim copies of someone else's text and must
never be edited.

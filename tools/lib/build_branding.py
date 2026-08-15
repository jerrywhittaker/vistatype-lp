#!/usr/bin/env python3
"""Generate the installer's icon and wizard images from the artwork in assets/branding/.

Two reasons this exists rather than a folder of hand-made files:

  1. Windows needs the icon as a multi-size .ico and the wizard needs several pixel sizes of
     each image, so that a high-DPI screen gets a sharp one instead of a stretched one. That is
     eleven files from two originals, and keeping eleven in step by hand is how they go stale.

  2. Until 8/15/2026 the installer set NO icon and NO version numbers, so the built Setup.exe
     was a generic Inno Setup stub reporting FileVersion 0.0.0.0 -- which is one of the two
     high-weight factors (the other being unsigned) that Microsoft's own people identified when
     an identical machine-learning detection hit one of their tools. See docs/Code-Signing.md.
     A file that looks like a product is less likely to be judged as though it were not one.

INPUT  assets/branding/*.png   hand-supplied artwork, the source of truth. Never edited here.
OUTPUT installer/branding/*    build outputs, tracked so the build box gets them. Regenerate
                               them, never hand-edit them.

installer/ is WIPED on the build box before each copy, so these travel with the rest of the
installer folder and a deleted one really does leave the next Setup.exe.

Usage:  python3 tools/lib/build_branding.py [--check-only]
"""
import sys
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print(
        "ERROR: this needs Pillow (the Python imaging library).\n"
        "       pip install --user Pillow",
        file=sys.stderr,
    )
    sys.exit(2)

SRC = Path("assets/branding")
OUT = Path("installer/branding")
ICON_SRC = SRC / "vistatype-icon.png"
WORDMARK_SRC = SRC / "vistatype-wordmark.png"

# Windows picks whichever of these fits the place it is drawing: 16 in a title bar, 32 on the
# desktop, 256 in a large Explorer view. An .ico carrying only one size gets scaled by Windows
# and looks it.
ICO_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)

# The wordmark, laid across the top of the welcome and finished pages.
#
# Jerry, 8/15/2026: the tall left-hand panel used to carry the mark AND the wordmark stacked, and
# having both on one screen looked wrong. The panel is gone from both of those pages; the
# wordmark sits at the top and the text runs the full width beneath it.
#
# WIDTHS in pixels at each display scaling from 100% to 250%. The installer draws whichever it
# picks at its NATURAL size -- no stretching at all, which is why there is one per step rather
# than one image scaled to fit: Windows' stretch is not smooth, and it shows on lettering.
#
# BMP, not PNG, because the wizard code loads this one itself and Inno Setup's scripting can
# only read a bitmap. Flattened onto white to match the page behind it.
WORDMARK_WIDTHS = (260, 325, 390, 455, 520, 650)
WORDMARK_BACKGROUND = (255, 255, 255)

# The small image sits top-right on every page after the welcome one. Square.
SMALL_SIZES = (58, 77, 97, 116, 124, 143, 159)

# The small image sits top-right on every page BETWEEN the welcome and finished ones, which is
# where the mark still appears.


def fit(img, width=None, height=None):
    """Scale to the given width or height, keeping the shape."""
    w, h = img.size
    if width:
        height = max(1, round(h * width / w))
    else:
        width = max(1, round(w * height / h))
    return img.resize((width, height), Image.LANCZOS)


def square(img, size):
    """Centre the image on a transparent square canvas of the given size."""
    scaled = fit(img, width=size) if img.width >= img.height else fit(img, height=size)
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    canvas.paste(scaled, ((size - scaled.width) // 2, (size - scaled.height) // 2), scaled)
    return canvas


def wordmark_strip(wordmark, width):
    """The wordmark at one exact width, flattened onto the page's own white."""
    scaled = fit(wordmark, width=width)
    canvas = Image.new("RGB", scaled.size, WORDMARK_BACKGROUND)
    canvas.paste(scaled, (0, 0), scaled)
    return canvas


def expected_outputs():
    names = ["vistatype.ico"]
    names += [f"welcome-wordmark-{w}.bmp" for w in WORDMARK_WIDTHS]
    names += [f"wizard-small-{s}.png" for s in SMALL_SIZES]
    return names


def main():
    check_only = "--check-only" in sys.argv

    for src in (ICON_SRC, WORDMARK_SRC):
        if not src.is_file():
            print(f"ERROR: missing artwork {src}", file=sys.stderr)
            return 2

    if check_only:
        missing = [n for n in expected_outputs() if not (OUT / n).is_file()]
        if missing:
            print(
                "ERROR: installer branding is not built. Missing:\n  "
                + "\n  ".join(missing)
                + "\nRun: make branding",
                file=sys.stderr,
            )
            return 1
        print(f"branding OK ({len(expected_outputs())} files in {OUT})")
        return 0

    OUT.mkdir(parents=True, exist_ok=True)
    icon = Image.open(ICON_SRC).convert("RGBA")
    wordmark = Image.open(WORDMARK_SRC).convert("RGBA")

    # Pillow writes every requested size into one .ico when they are passed as `sizes`, but it
    # takes them from ONE source image, so hand it the largest square and let it downscale.
    largest = max(ICO_SIZES)
    square(icon, largest).save(OUT / "vistatype.ico", format="ICO", sizes=[(s, s) for s in ICO_SIZES])

    for w in WORDMARK_WIDTHS:
        wordmark_strip(wordmark, w).save(OUT / f"welcome-wordmark-{w}.bmp")

    for s in SMALL_SIZES:
        square(icon, s).save(OUT / f"wizard-small-{s}.png")

    stale = sorted(OUT.glob("wizard-2*.png")) + sorted(OUT.glob("wizard-3*.png")) + \
        sorted(OUT.glob("wizard-4*.png"))
    for old in stale:
        old.unlink()

    made = expected_outputs()
    print(f"Wrote {len(made)} files to {OUT}/ from {SRC}/:")
    print(f"  vistatype.ico            {len(ICO_SIZES)} sizes, {'/'.join(str(s) for s in ICO_SIZES)}")
    print(f"  welcome-wordmark-*.bmp   {len(WORDMARK_WIDTHS)} widths for the welcome/finished pages")
    print(f"  wizard-small-*.png       {len(SMALL_SIZES)} sizes for the top-right corner")
    if stale:
        print(f"  removed {len(stale)} superseded welcome-panel image(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

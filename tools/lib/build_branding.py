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

# The welcome page's image area, at the DPI settings Inno Setup documents. Setup picks the
# closest and does not have to stretch. All four are the same shape, so nothing is distorted.
WIZARD_SIZES = ((202, 386), (269, 515), (336, 643), (404, 772))

# The small image sits top-right on every page after the welcome one. Square.
SMALL_SIZES = (58, 77, 97, 116, 124, 143, 159)

# Fractions of the welcome panel. The mark and the name are stacked as ONE group and that group
# is centred, rather than each being pinned to its own fixed height -- pinning them left the
# bottom third of the panel empty and the two drifting apart. Slightly above true centre,
# because an optically centred block sits a little high.
ICON_WIDTH = 0.58
WORDMARK_WIDTH = 0.86
GAP = 0.045
GROUP_CENTRE = 0.44


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


def welcome_panel(icon, wordmark, size):
    """The tall image on the welcome page: the mark above, the name below.

    Transparent, so it sits on the wizard's own background whatever Windows theme is in use.
    """
    width, height = size
    canvas = Image.new("RGBA", size, (0, 0, 0, 0))

    mark = fit(icon, width=round(width * ICON_WIDTH))
    name = fit(wordmark, width=round(width * WORDMARK_WIDTH))
    gap = round(height * GAP)

    stack = mark.height + gap + name.height
    top = round(height * GROUP_CENTRE) - stack // 2

    canvas.paste(mark, ((width - mark.width) // 2, top), mark)
    canvas.paste(name, ((width - name.width) // 2, top + mark.height + gap), name)
    return canvas


def expected_outputs():
    names = ["vistatype.ico"]
    names += [f"wizard-{w}x{h}.png" for w, h in WIZARD_SIZES]
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

    for w, h in WIZARD_SIZES:
        welcome_panel(icon, wordmark, (w, h)).save(OUT / f"wizard-{w}x{h}.png")

    for s in SMALL_SIZES:
        square(icon, s).save(OUT / f"wizard-small-{s}.png")

    made = expected_outputs()
    print(f"Wrote {len(made)} files to {OUT}/ from {SRC}/:")
    print(f"  vistatype.ico          {len(ICO_SIZES)} sizes, {'/'.join(str(s) for s in ICO_SIZES)}")
    print(f"  wizard-*.png           {len(WIZARD_SIZES)} sizes for the welcome page")
    print(f"  wizard-small-*.png     {len(SMALL_SIZES)} sizes for the top-right corner")
    return 0


if __name__ == "__main__":
    sys.exit(main())

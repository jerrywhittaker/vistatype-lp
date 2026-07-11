#!/usr/bin/env python3
"""Inject an embedded customUI14 ribbon part into a Word .dotm package (Linux-only).

A .dotm is a zip. This adds:
  * customUI/customUI14.xml            (the ribbon definition)
  * a root relationship pointing at it (type .../2007/relationships/ui/extensibility)

The [Content_Types].xml already declares a Default for the "xml" extension in Word
packages, so no content-type edit is needed. Re-running is idempotent: an existing
customUI part and its relationship are replaced, not duplicated.

This is the last step of `make build`, after Word has produced the .dotm (so the
RibbonCallbacks dispatcher module is already compiled in). No Word required.

Usage:  python3 tools/lib/inject_customui.py <target.dotm> [src/ribbon/customUI14.xml]
"""
import re
import shutil
import sys
import tempfile
import zipfile

REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL_ID = "rIdVistaTypeCustomUI"
PART = "customUI/customUI14.xml"
RELS = "_rels/.rels"


def add_relationship(rels_xml: str) -> str:
    # Drop any prior VistaType customUI relationship, then insert ours.
    rels_xml = re.sub(
        r'<Relationship\b[^>]*Target="customUI/customUI14\.xml"[^>]*/>', "", rels_xml
    )
    rel = (f'<Relationship Id="{REL_ID}" Type="{REL_TYPE}" '
           f'Target="{PART}"/>')
    return rels_xml.replace("</Relationships>", rel + "</Relationships>")


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    dotm = sys.argv[1]
    ribbon_path = sys.argv[2] if len(sys.argv) > 2 else "src/ribbon/customUI14.xml"
    with open(ribbon_path, "r", encoding="utf-8") as fh:
        ribbon = fh.read()

    with zipfile.ZipFile(dotm) as z:
        names = z.namelist()
        if RELS not in names:
            sys.exit(f"ERROR: {RELS} not found in {dotm}")
        items = {n: z.read(n) for n in names}

    items[RELS] = add_relationship(items[RELS].decode("utf-8")).encode("utf-8")
    items[PART] = ribbon.encode("utf-8")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".dotm")
    tmp.close()
    with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml conventionally comes first.
        order = sorted(items, key=lambda n: (n != "[Content_Types].xml", n))
        for n in order:
            z.writestr(n, items[n])
    shutil.move(tmp.name, dotm)
    print(f"Injected {PART} into {dotm} ({len(ribbon)} bytes, rel {REL_ID}).")


if __name__ == "__main__":
    main()

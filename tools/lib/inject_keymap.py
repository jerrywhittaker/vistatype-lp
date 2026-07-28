#!/usr/bin/env python3
"""Inject a keyboard-shortcut (keymap) part into a Word package (Linux-only).

Word stores keyboard assignments in `word/customizations.xml` inside the .dotx/.dotm zip.
That part is *source* -- it holds bindings Jerry assigned by hand -- but it is trapped
inside a binary, so this mirrors tools/lib/inject_customui.py and lets the text under
src/keymap/ be the thing that is edited and tracked.

Adds/replaces:
  * word/customizations.xml
  * word/_rels/document.xml.rels  -> a keyMapCustomizations relationship
  * [Content_Types].xml           -> the keyMapCustomizations override

Re-running is idempotent: an existing keymap part, relationship and override are replaced,
never duplicated. No Word required.

Usage:  python3 tools/lib/inject_keymap.py <target.dotx|.dotm> <keymap.xml>
"""
import re
import shutil
import sys
import tempfile
import zipfile

PART = "word/customizations.xml"
RELS = "word/_rels/document.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations"
REL_ID = "rIdVistaTypeKeymap"
CT = "application/vnd.ms-word.keyMapCustomizations+xml"


def add_relationship(rels_xml: str) -> str:
    """Point a keyMapCustomizations relationship at the keymap part.

    Word writes the Target relative to word/, i.e. "customizations.xml". Any prior
    relationship of this type is dropped first so re-running cannot duplicate it.
    """
    if REL_TYPE in rels_xml:
        return rels_xml  # already wired (normal case: Word authored it)
    rel = f'<Relationship Id="{REL_ID}" Type="{REL_TYPE}" Target="customizations.xml"/>'
    return rels_xml.replace("</Relationships>", rel + "</Relationships>")


def add_content_type(ct_xml: str) -> str:
    if f'PartName="/{PART}"' in ct_xml:
        return ct_xml
    override = f'<Override PartName="/{PART}" ContentType="{CT}"/>'
    return ct_xml.replace("</Types>", override + "</Types>")


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    package, keymap_path = sys.argv[1], sys.argv[2]

    with open(keymap_path, "r", encoding="utf-8") as fh:
        keymap = fh.read()

    with zipfile.ZipFile(package) as z:
        names = z.namelist()
        if RELS not in names:
            sys.exit(f"ERROR: {RELS} not found in {package}")
        if CONTENT_TYPES not in names:
            sys.exit(f"ERROR: {CONTENT_TYPES} not found in {package}")
        items = {n: z.read(n) for n in names}

    items[RELS] = add_relationship(items[RELS].decode("utf-8")).encode("utf-8")
    items[CONTENT_TYPES] = add_content_type(items[CONTENT_TYPES].decode("utf-8")).encode("utf-8")
    items[PART] = keymap.encode("utf-8")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    tmp.close()
    with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as z:
        order = sorted(items, key=lambda n: (n != CONTENT_TYPES, n))
        for n in order:
            z.writestr(n, items[n])
    shutil.move(tmp.name, package)

    macros = len(re.findall(r"wne:macroName=", keymap))
    keys = len(re.findall(r"<wne:keymap\b", keymap))
    print(f"Injected {PART} into {package} "
          f"({len(keymap)} bytes, {keys} key assignments, {macros} bound to macros).")


if __name__ == "__main__":
    main()

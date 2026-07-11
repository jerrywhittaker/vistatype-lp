#!/usr/bin/env python3
"""Extract VistaType's QAT macro buttons from Word.officeUI into a data fragment.

The installer merges these into each user's own Word.officeUI (non-destructively),
so a fresh install arrives with the VistaType quick-access icons pre-stocked while
the user's ribbon and their own QAT items are left untouched.

Emits installer/qat-controls.xml -- a simple, namespace-free list the install-time
PowerShell reads:

    <qatControls>
      <button macro="Sh_Doc_Info" label="Document Settings" imageMso="Info"/>
      ...
    </qatControls>

Usage:  python3 tools/lib/extract_qat.py [Word.officeUI] [installer/qat-controls.xml]
"""
import sys
import xml.etree.ElementTree as ET

MSO = "http://schemas.microsoft.com/office/2009/07/customui"


def esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "Word.officeUI"
    dst = sys.argv[2] if len(sys.argv) > 2 else "installer/qat-controls.xml"

    root = ET.parse(src).getroot()
    shared = root.find(f".//{{{MSO}}}qat/{{{MSO}}}sharedControls")
    if shared is None:
        sys.exit("No <qat>/<sharedControls> found in " + src)

    seen, rows = set(), []
    for e in list(shared):
        if e.tag.endswith("}button") and e.attrib.get("onAction"):
            macro = e.attrib["onAction"]
            if macro in seen:
                continue
            seen.add(macro)
            rows.append((macro, e.attrib.get("label", ""), e.attrib.get("imageMso", "")))

    out = ["<qatControls>"]
    for macro, label, image in rows:
        out.append(f'  <button macro="{esc(macro)}" label="{esc(label)}" '
                   f'imageMso="{esc(image)}"/>')
    out.append("</qatControls>")
    with open(dst, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out) + "\n")
    print(f"Wrote {dst} ({len(rows)} QAT buttons)")


if __name__ == "__main__":
    main()

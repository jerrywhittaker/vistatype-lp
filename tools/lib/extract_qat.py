#!/usr/bin/env python3
"""OBSOLETE / REFERENCE ONLY -- no longer part of the build or install pipeline.

The QAT is now a hand-maintained full toolbar in installer/qat-template.officeUI, which the
installer imposes via installer/scripts/Merge-Qat.ps1 (referencing the add-in's ribbon
controls, writing both the Roaming and Local Word.officeUI). This extractor used to emit the
old installer/qat-controls.xml (macro-button list); that file has been removed. Kept only as a
reference for how the legacy Word.officeUI QAT was parsed. Edit qat-template.officeUI by hand.

This extractor only ever emitted the custom macro buttons:

    <qatControls>
      <button macro="Sh_Doc_Info" label="Document Settings" imageMso="Info"/>
      ...
    </qatControls>

Usage (reference only):  python3 tools/lib/extract_qat.py [Word.officeUI] [out.xml]
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

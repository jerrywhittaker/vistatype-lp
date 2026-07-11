#!/usr/bin/env python3
"""Convert the legacy global Word.officeUI ribbon into an embedded customUI14.xml.

The .officeUI file is a whole-ribbon customization that OVERWRITES each user's
ribbon/QAT. An embedded customUI part inside the template MERGES instead, and
travels with the add-in. This extracts the three custom VistaType tabs and rewrites
every macro button to a single VBA dispatcher (RibbonAction) that runs the macro
named in the control's `tag` -- so the existing parameterless Subs need no changes.

Built-in Word controls (idQ="mso:...") become <control idMso="..."/>. QAT
customizations are intentionally NOT carried over (a template customUI can't merge
them reliably; the "LP and BRL QAT Icons" tab lets users self-populate their QAT).

Usage:  python3 tools/lib/officeui_to_customui.py [Word.officeUI] [out.xml]
        defaults: Word.officeUI -> src/ribbon/customUI14.xml
"""
import re
import sys
import xml.etree.ElementTree as ET

MSO = "http://schemas.microsoft.com/office/2009/07/customui"
NS = {"mso": MSO}


def sanitize(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]", "_", s)


def esc(s: str) -> str:
    """XML-escape an attribute value (labels may contain literal '&&' mnemonics)."""
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def attr(e, name):
    return e.attrib.get(name) or e.attrib.get(f"{{{MSO}}}{name}")


def convert_control(e, out, ids):
    tag = e.tag.split("}")[-1]
    if tag == "button":
        macro = attr(e, "onAction")
        if not macro:
            return  # non-macro custom button we can't map; skip
        bid = f"btn_{sanitize(macro)}"
        while bid in ids:
            bid += "_x"
        ids.add(bid)
        parts = [f'id="{bid}"']
        if attr(e, "label"):
            parts.append(f'label="{esc(attr(e, "label"))}"')
        if attr(e, "imageMso"):
            parts.append(f'imageMso="{esc(attr(e, "imageMso"))}"')
        parts.append('onAction="RibbonAction"')
        parts.append(f'tag="{macro}"')
        out.append("        <button " + " ".join(parts) + "/>")
    elif tag == "control":
        idq = attr(e, "idQ") or ""
        if idq.startswith("mso:"):
            out.append(f'        <control idMso="{idq[4:]}"/>')  # built-in command
    elif tag == "separator":
        sid = f"sep_{len(ids)}"
        ids.add(sid)
        out.append(f'        <separator id="{sid}"/>')


def convert_group(g, out, ids):
    gid = attr(g, "id") or f"grp_{len(ids)}"
    gid = "grp_" + sanitize(gid)[-16:]
    while gid in ids:
        gid += "_x"
    ids.add(gid)
    parts = [f'id="{gid}"']
    if attr(g, "label"):
        parts.append(f'label="{esc(attr(g, "label"))}"')
    if attr(g, "imageMso"):
        parts.append(f'imageMso="{esc(attr(g, "imageMso"))}"')
    if attr(g, "autoScale"):
        parts.append(f'autoScale="{attr(g, "autoScale")}"')
    out.append("      <group " + " ".join(parts) + ">")
    for child in g:
        convert_control(child, out, ids)
    out.append("      </group>")


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "Word.officeUI"
    dst = sys.argv[2] if len(sys.argv) > 2 else "src/ribbon/customUI14.xml"

    tree = ET.parse(src)
    root = tree.getroot()

    # Custom tabs are <mso:tab> with an `id` (built-in tab toggles use idQ instead).
    tabs = [t for t in root.iter(f"{{{MSO}}}tab") if attr(t, "id") and attr(t, "label")]

    ids = set()
    out = []
    out.append(f'<customUI xmlns="{MSO}">')
    out.append("  <ribbon>")
    out.append("    <tabs>")
    for t in tabs:
        label = attr(t, "label")
        tid = "tab_" + sanitize(label)
        parts = [f'id="{tid}"', f'label="{esc(label)}"']
        # All three custom tabs sit just left of the built-in Insert tab.
        parts.append('insertBeforeMso="TabInsert"')
        out.append("    <tab " + " ".join(parts) + ">")
        for g in t:
            if g.tag.endswith("}group"):
                convert_group(g, out, ids)
        out.append("    </tab>")
    out.append("    </tabs>")
    out.append("  </ribbon>")
    out.append("</customUI>")

    with open(dst, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out) + "\n")
    print(f"Wrote {dst}  ({len(tabs)} tabs, {sum(1 for l in out if '<button' in l)} macro buttons)")


if __name__ == "__main__":
    main()

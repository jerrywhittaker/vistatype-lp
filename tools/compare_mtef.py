#!/usr/bin/env python3
"""Compare VistaType LP's MathType equations with MathType's own - read only.

The DAISY/NIMAS converter writes each MathML equation as MathType's MTEF itself
(Sh_MTEF_Walk and the Sh_MTEF_* helpers in src/vba/LPandBrlMacros.bas) and falls back to a
Word equation when the walker refuses one. This is a faithful Python copy of that walker, run
over a whole book, so a change to the VBA can be measured before it is built:

  - how many of the book's equations the walker accepts, and why it refuses the rest;
  - for each accepted equation, its records against MathType's own for the same equation, taken
    from a copy of the converted book after MathType's Convert Equations (every equation then a
    MathType object, in book order). Size and color records are set aside - MathType's
    conversion adds COLOR records, and the sizes follow its own preferences.

KEEP IT IN STEP WITH THE VBA. It is a measuring tool only as long as it walks exactly as
Sh_MTEF_Walk does; change both together. It reads files and prints; it writes nothing.

Usage:
    python3 -I tools/compare_mtef.py BOOK.xml CONVERTED.docx [--mixed MIXED.docx]
                                     [--show KIND] [--limit N] [--index I]

  BOOK.xml        the DAISY/NIMAS book's .xml
  CONVERTED.docx  the converter's output after MathType's Convert Equations
  --mixed         the converter's output BEFORE that conversion: its Word equations mark the
                  equations the walker refused when it was made, and the report is then about
                  those (for example the 503 of Basic Math, 11/5/2026)
  --show KIND     print the equations whose first difference is of this kind (a substring)
  --index I       print one equation (0-based, book order): MathML, VistaType LP's, MathType's

Needs olefile (pip install olefile). Written 11/5/2026 for the Basic Math fallbacks (Jerry).
"""
import argparse
import collections
import difflib
import io
import re
import sys
import zipfile
from xml.dom import minidom

try:
    import olefile
except ImportError:
    sys.exit("needs the 'olefile' package (pip install olefile)")

NS = "http://www.w3.org/1998/Math/MathML"


# --- MathML out of the book: Sh_MathML_Uncomment, Sh_MathML_Fragments, Sh_MathML_With_Namespace

def math_pattern(p):
    return ("<" + p + "math(?=[\\s>/])(?:\"[^\"]*\"|'[^']*'|[^'\">/]|/(?!>))*"
            "(?:/>|>[\\s\\S]*?</" + p + "math\\s*>)")


def attr(tag, name):
    m = re.search(r'\s' + re.escape(name) + r'\s*=\s*(?:"([^"]*)"|\'([^\']*)\')', tag)
    if not m:
        return ''
    return m.group(1) if m.group(1) is not None else m.group(2)


def is_eq_picture(img):
    src = attr(img, 'src').strip()
    if src == '' or src[:5].lower() == 'data:':
        return False
    src = src.split('?')[0].split('#')[0]
    cut = max(src.rfind('/'), src.rfind('\\'))
    return 'eq' in src[cut + 1:].lower()


def uncomment(xml, p):
    rx = re.compile(r"(?:(<img(?=[\s>/])(?:\"[^\"]*\"|'[^']*'|[^'\">])*>)(\s*))?<!--([\s\S]*?)-->")
    one = re.compile(r"^\s*(" + math_pattern(p) + r")\s*$")
    out, pos, tag = [], 0, '<' + p + 'math'
    for m in rx.finditer(xml):
        out.append(xml[pos:m.start()])
        img, gap = m.group(1) or '', m.group(2) or ''
        mathml = ''
        mm = one.match(m.group(3) or '')
        if mm:
            mathml = mm.group(1)
            if mathml.find(tag, 1) >= 0:
                mathml = ''
        if mathml and img and not is_eq_picture(img):
            mathml = ''
        out.append(img + gap + mathml)
        pos = m.end()
    out.append(xml[pos:])
    return ''.join(out)


def fragments(xml, p):
    return [m.group(0) for m in re.finditer(math_pattern(p), xml)]


def with_ns(frag, p):
    tag = '<' + p + 'math'
    if not frag.startswith(tag):
        return frag
    decl = 'xmlns:' + p[:-1] if p else 'xmlns'
    rx = r"<[^\s>/]+(?:\"[^\"]*\"|'[^']*'|[^'\">])*>"
    m = re.match(rx, frag)
    gt = m.end() if m else len(frag)
    head = re.sub(r"\s+xmlns[0-9]+:[A-Za-z_][A-Za-z0-9_]*\s*=\s*(?:\"[^\"]*\"|'[^']*')", "", frag[:gt])
    frag = head + frag[gt:]
    m = re.match(rx, frag)
    gt = m.end() if m else len(frag)
    if attr(frag[:gt], decl) != '':
        return frag
    return tag + ' ' + decl + '="' + NS + '"' + frag[len(tag):]


def book_fragments(path):
    text = open(path, 'rb').read().decode('utf-8')
    return [with_ns(f, 'm:') for f in fragments(uncomment(text, 'm:'), 'm:')]


# --- the walker: Sh_MTEF_Walk and its helpers --------------------------------------------------

def u16(s):
    """VBA strings are UTF-16 code units."""
    b = s.encode('utf-16-le')
    return [b[i] | (b[i + 1] << 8) for i in range(0, len(b), 2)]


def bname(n):
    return n.localName if n.nodeType == 1 else ''


def text(n):
    if n.nodeType in (3, 4):
        return n.data
    return ''.join(text(c) for c in n.childNodes)


def collapse(s):
    s = s.replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')
    while '  ' in s:
        s = s.replace('  ', ' ')
    return s.strip(' ')


def getattr_(n, a):
    return n.getAttribute(a) if n.hasAttribute(a) else ''


def kids_of(n):
    c = []
    for k in n.childNodes:
        if k.nodeType == 1:
            c.append(k)
        elif k.nodeType in (3, 4):
            if collapse(k.data) != '':
                return None
    return c


def mo_text(n):
    if bname(n) != 'mo' or n.namespaceURI != NS:
        return ''
    for k in n.childNodes:
        if k.nodeType == 1:
            return ''
    return collapse(text(n))


def match_fence(kids, i, last):
    stack = ''
    for j in range(i, last + 1):
        t = mo_text(kids[j])
        if t in ('(', '[', '{'):
            stack += t
        elif t in (')', ']', '}'):
            if stack == '':
                return 0
            if stack[-1] != '([{'[')]}'.index(t)]:
                return -1
            stack = stack[:-1]
            if stack == '':
                return j
        elif t == '|':
            if stack[-1:] == '|':
                stack = stack[:-1]
                if stack == '':
                    return j
            else:
                stack += t
    return 0


class St:
    def __init__(self):
        self.out = bytearray()
        self.size = 10
        self.styles = []      # Sh_MTEF_Walk's "styles": the FONT_STYLE_DEFs written, in order
        self.why = ''


def fail(st, why):
    if st.why == '':
        st.why = why


def add(st, b):
    st.out.append(b & 0xFF)


def set_size(st, size):
    if st.size != size:
        add(st, size)
        st.size = size


# Sh_MTEF_Style_Index: bold italic Times New Roman, and the two Euclid fonts, each through a
# FONT_STYLE_DEF written once at its first use (with, for a Euclid font, its ENCODING_DEF and
# FONT_DEF). A character names one as typeface -1 for the first written, -2 the second, -3 the
# third. MEASURED 11/5/2026 through MathType's own MathML import on VistaBuild.
EUCLID = {'E1': 'Euclid Math One', 'E2': 'Euclid Math Two'}


def style_index(st, key):
    if key not in st.styles:
        if key == 'B':
            for b in (8, 1, 3):
                add(st, b)
        else:
            enc = 'EuclidMath' + key[1]
            n = sum(1 for k in st.styles if k != 'B')
            add(st, 0x13)
            for ch in enc.encode('ascii'):
                add(st, ch)
            add(st, 0)
            add(st, 0x11)
            add(st, 6 + n)
            for ch in EUCLID[key].encode('ascii'):
                add(st, ch)
            add(st, 0)
            add(st, 8)
            add(st, 5 + n)
            add(st, 0)
        st.styles.append(key)
    return st.styles.index(key) + 1


def put_char(st, tf, code, f8, size, emb):
    if tf == -1:
        tf = -style_index(st, 'B')
    elif tf == -11:
        tf = -style_index(st, 'E1')
    elif tf == -12:
        tf = -style_index(st, 'E2')
    set_size(st, size)
    add(st, 2)
    add(st, (4 if f8 >= 0 else 0) | (1 if emb else 0))
    add(st, tf + 128)
    add(st, code & 0xFF)
    add(st, (code >> 8) & 0xFF)
    if f8 >= 0:
        add(st, f8)
    if emb:
        add(st, 6)
        add(st, 0)
        add(st, emb)
        add(st, 0)


def put_template(st, sel, var, size):
    set_size(st, size)
    for b in (3, 0, sel, var, 0):
        add(st, b)


def put_line(st, kids, first, last, size):
    set_size(st, size)
    add(st, 1)
    if kids is None:
        add(st, 1)
        return
    add(st, 0)
    put_row(st, kids, first, last, size)
    add(st, 0)


def put_fence(st, kids, i, j, size):
    t = mo_text(kids[i])
    sel, o, c = {'(': (1, 0x28, 0x29), '{': (2, 0x7B, 0x7D), '[': (3, 0x5B, 0x5D),
                 '|': (4, 0xEC07, 0xEC08)}[t]
    put_template(st, sel, 3, size)
    put_line(st, kids, i + 1, j - 1, size)
    put_char(st, 22, o, -1, size, 0)
    put_char(st, 22, c, -1, size, 0)
    add(st, 0)


def put_row(st, kids, first, last, size):
    i = first
    while i <= last:
        if st.why:
            return
        t = mo_text(kids[i])
        j = 0
        if t in ('(', '[', '{', '|'):
            j = match_fence(kids, i, last)
        if j < 0:
            fail(st, 'brackets that do not match')
            return
        elif j > 0:
            put_fence(st, kids, i, j, size)
            i = j + 1
        else:
            put_node(st, kids[i], size)
            i += 1


ACCEPT = ("mrow", "mstyle", "semantics", "msup", "msub", "msubsup", "mfrac", "msqrt", "mroot",
          "menclose", "mfenced", "mover", "munder", "mtable", "mspace", "mmultiscripts")

# Sh_MTEF_Space_Codes: an mspace's width as MathType's own MathML import writes it, one or two
# of its SPACE characters. MEASURED 11/5/2026 on VistaBuild; anything else is refused.
SPACES = {'': (0xEF01,), 'veryverythinmathspace': (0xEF01,), 'verythinmathspace': (0xEF07,),
          'thinmathspace': (0xEF07,), 'mediummathspace': (0xEF03,), 'thickmathspace': (0xEF03,),
          'verythickmathspace': (0xEF04,), 'veryverythickmathspace': (0xEF04, 0xEF07),
          '0.1667em': (0xEF07,), '0.2222em': (0xEF03,), '0.25em': (0xEF03,),
          '0.2778em': (0xEF03,), '0.3333em': (0xEF04,), '0.5em': (0xEF04, 0xEF03),
          '1em': (0xEF05,), '2em': (0xEF06,)}


def put_node(st, node, size):
    if st.why:
        return
    if node.namespaceURI != NS:
        fail(st, node.nodeName + ' is not MathML')
        return
    tag = bname(node)
    if tag in ('mi', 'mn', 'mo', 'mtext'):
        put_token(st, node, size, 0)
        return
    if tag not in ACCEPT:
        fail(st, tag + ' is not handled')
        return
    kids = kids_of(node)
    if kids is None:
        fail(st, 'text outside a token element, in ' + tag)
        return
    n = len(kids)
    if tag in ('mrow', 'mstyle'):
        if tag == 'mrow' and n >= 2:
            t = mo_text(kids[0])
            if t in ('(', '[', '{', '|'):
                if match_fence(kids, 0, n - 1) == n - 1:
                    put_fence(st, kids, 0, n - 1, size)
                    return
        put_row(st, kids, 0, n - 1, size)
        return
    if tag == 'semantics':
        if n == 0:
            fail(st, 'an empty semantics is not handled')
            return
        for k in kids[1:]:
            if bname(k) not in ('annotation', 'annotation-xml'):
                fail(st, k.nodeName + ' in semantics is not handled')
                return
        put_node(st, kids[0], size)
        return
    if tag == 'mspace':
        put_space(st, node, kids, size)
        return
    {'msup': put_scripts, 'msub': put_scripts, 'msubsup': put_scripts, 'mfrac': put_fraction,
     'msqrt': put_root, 'mroot': put_root, 'menclose': put_enclose, 'mfenced': put_fenced,
     'mover': put_accent, 'munder': put_accent, 'mtable': put_table,
     'mmultiscripts': put_prescripts}[tag](st, node, kids, size)


# True only while comparing: MathType's own route (Word equation, then Convert Equations) drops
# every mspace, so the comparison is made once as the converter writes and once without them.
OMML_ROUTE = False


def put_space(st, node, kids, size):
    w = collapse(getattr_(node, 'width')).lower()
    if len(kids) > 0 or w not in SPACES or getattr_(node, 'linebreak'):
        fail(st, 'mspace width="%s" is not handled' % w)
        return
    if OMML_ROUTE:
        return
    for code in SPACES[w]:
        put_char(st, 24, code, -1, size, 0)


def put_prescripts(st, node, kids, size):
    # mmultiscripts with scripts BEFORE the base only (borrowing marks): template 29 variation 1,
    # the subscript then the superscript slot, an absent or <none/> one an empty slot; the base
    # after the template. As MathType's own conversion wrote them in Basic Math (7 of 7).
    ok = len(kids) in (3, 4) and bname(kids[1]) == 'mprescripts' and bname(kids[0]) not in ('mprescripts', 'none')
    if ok:
        for k in kids[2:]:
            if bname(k) == 'mprescripts':
                ok = False
    if not ok:
        fail(st, 'mmultiscripts other than base, mprescripts and one or two scripts is not handled')
        return
    ss = script_size(size)
    put_template(st, 29, 1, size)
    for k in (kids[2:] + [None])[:2]:
        if k is None or bname(k) == 'none':
            put_line(st, [], 0, -1, ss)
        else:
            put_line(st, [k], 0, 0, ss)
    add(st, 0)
    put_node(st, kids[0], size)


def script_size(s):
    return 12 if s >= 11 else 11


def one_char(node):
    while bname(node) == 'mrow':
        k = kids_of(node)
        if k is None or len(k) != 1:
            return None
        node = k[0]
    if node.namespaceURI != NS:
        return None
    if bname(node) not in ('mi', 'mn', 'mo', 'mtext'):
        return None
    k = kids_of(node)
    if k is not None and len(k) > 0:
        return None
    tt = u16(collapse(text(node)))
    if len(tt) != 1:
        return None
    if 0x2061 <= tt[0] <= 0x2063:
        return None
    return node


def put_scripts(st, node, kids, size):
    tag = bname(node)
    if tag == 'msubsup':
        if len(kids) != 3:
            fail(st, 'msubsup needs three arguments')
            return
        sel = 29
    else:
        if len(kids) != 2:
            fail(st, tag + ' needs two arguments')
            return
        sel = 28 if tag == 'msup' else 27
    ss = script_size(size)
    emb = 0
    if sel == 28:
        sc = kids[1]
        while bname(sc) == 'mrow':
            sk = kids_of(sc)
            if sk is None or len(sk) != 1:
                break
            sc = sk[0]
        t = mo_text(sc)
        if t in ("'", '′'):
            emb = 5
        elif t == '″':
            emb = 6
        elif t == '‴':
            emb = 18
        elif bname(sc) == 'mrow':
            sk = kids_of(sc)
            if sk is not None and len(sk) == 2 and mo_text(sk[0]) == '′' and mo_text(sk[1]) == '′':
                emb = 6
        if emb:
            b = one_char(kids[0])
            if b is not None:
                put_token(st, b, size, emb)
                return
    put_node(st, kids[0], size)
    put_template(st, sel, 0, size)
    if sel == 28:
        put_line(st, None, 0, 0, ss)
        put_line(st, [kids[1]], 0, 0, ss)
    elif sel == 27:
        put_line(st, [kids[1]], 0, 0, ss)
        put_line(st, None, 0, 0, ss)
    else:
        put_line(st, [kids[1]], 0, 0, ss)
        put_line(st, [kids[2]], 0, 0, ss)
    add(st, 0)


def put_fraction(st, node, kids, size):
    if len(kids) != 2:
        fail(st, 'mfrac needs two arguments')
        return
    lt = collapse(getattr_(node, 'linethickness')).lower()
    bev = collapse(getattr_(node, 'bevelled')).lower()
    if bev not in ('', 'true', 'false'):
        fail(st, 'mfrac bevelled="%s" is not handled' % bev)
        return
    if re.search(r'^(0+\.?0*|\.0+)(px|pt|em|ex|mm|cm|in)?$', lt):
        if bev == 'true' or size != 10:
            fail(st, 'a fraction with no line, bevelled or in a script, is not handled')
            return
        set_size(st, size)
        for b in (5, 0, 1, 1, 1, 2, 1, 0, 0):
            add(st, b)
    else:
        put_template(st, 11, 2 if bev == 'true' else 0, size)
    put_line(st, [kids[0]], 0, 0, size)
    put_line(st, [kids[1]], 0, 0, size)
    add(st, 0)


def put_root(st, node, kids, size):
    if bname(node) == 'msqrt':
        put_template(st, 10, 0, size)
        put_line(st, kids, 0, len(kids) - 1, size)
        put_line(st, None, 0, 0, script_size(size))
    else:
        if len(kids) != 2:
            fail(st, 'mroot needs two arguments')
            return
        put_template(st, 10, 1, size)
        put_line(st, [kids[0]], 0, 0, size)
        put_line(st, [kids[1]], 0, 0, script_size(size))
    add(st, 0)


def put_enclose(st, node, kids, size):
    nt = collapse(getattr_(node, 'notation'))
    if nt == 'radical':
        put_template(st, 10, 0, size)
        put_line(st, kids, 0, len(kids) - 1, size)
        put_line(st, None, 0, 0, script_size(size))
        add(st, 0)
        return
    if nt == 'longdiv':
        # MEASURED 11/5/2026, MathType's own MathML import: template 26 variation 0, the
        # dividend, then an absent slot (no quotient). The divisor is whatever comes before.
        put_template(st, 26, 0, size)
        put_line(st, kids, 0, len(kids) - 1, size)
        put_line(st, None, 0, 0, size)
        add(st, 0)
        return
    tbl = {'box': (37, 30), 'top': (13, 0), 'bottom': (12, 0), 'left': (4, 1), 'right': (4, 2),
           'updiagonalstrike': (36, 2), 'downdiagonalstrike': (36, 4), 'horizontalstrike': (36, 1)}
    if nt not in tbl:
        fail(st, 'menclose notation="%s" is not handled' % nt)
        return
    put_template(st, tbl[nt][0], tbl[nt][1], size)
    put_line(st, kids, 0, len(kids) - 1, size)
    if nt == 'left':
        put_char(st, 22, 0xEC07, -1, size, 0)
    if nt == 'right':
        put_char(st, 22, 0xEC08, -1, size, 0)
    add(st, 0)


def put_fenced(st, node, kids, size):
    if len(kids) > 1:
        fail(st, 'mfenced with more than one argument is not handled')
        return
    o = collapse(node.getAttribute('open')) if node.hasAttribute('open') else '('
    c = collapse(node.getAttribute('close')) if node.hasAttribute('close') else ')'
    tbl = {'()': (1, 0x28, 0x29), '{}': (2, 0x7B, 0x7D), '[]': (3, 0x5B, 0x5D), '||': (4, 0xEC07, 0xEC08)}
    if o + c not in tbl:
        fail(st, 'mfenced open="%s" close="%s" is not handled' % (o, c))
        return
    sel, oc, cc = tbl[o + c]
    put_template(st, sel, 3, size)
    put_line(st, kids, 0, len(kids) - 1, size)
    put_char(st, 22, oc, -1, size, 0)
    put_char(st, 22, cc, -1, size, 0)
    add(st, 0)


def put_accent(st, node, kids, size):
    if len(kids) != 2:
        fail(st, bname(node) + ' needs two arguments')
        return
    au = u16(mo_text(kids[1]))
    if len(au) != 1:
        fail(st, bname(node) + ' that is not an accent is not handled')
        return
    code = au[0]
    over = bname(node) == 'mover'
    isacc = collapse(getattr_(node, 'accent')).lower() == 'true'
    emb, sel, var, exp, need = 0, -1, 0, 0, False
    if over:
        if code in (0xAF, 0x305):
            emb, sel = 17, 13
        elif code == 0x5E:
            emb, sel, exp, need = 9, 33, 0x302, True
        elif code == 0x7E:
            emb, sel, exp, need = 8, 32, 0x303, True
        elif code == 0x2DC:
            emb = 8
        elif code == 0x2D9:
            emb = 2
        elif code == 0x2192:
            emb, sel, var, exp, need = 11, 31, 2, 0x20D7, True
        elif code == 0x21C0:
            sel, var, exp, need = 31, 10, 0x20D1, True
        elif code == 0x2194:
            sel, var, exp, need = 31, 3, 0x20E1, True
    else:
        if code == 0x5F:
            emb, sel = 29, 12
        elif code in (0x332, 0xAF):
            emb = 29
    oc = one_char(kids[0])
    if oc is not None and emb:
        put_token(st, oc, size, emb)
    elif oc is None and sel >= 0 and (isacc or not need):
        put_template(st, sel, var, size)
        put_line(st, [kids[0]], 0, 0, size)
        if exp:
            put_char(st, 22, exp, -1, size, 0)
        add(st, 0)
    else:
        fail(st, '%s with U+%04X is not handled' % (bname(node), code))


def put_table(st, node, rows, size):
    if size != 10:
        fail(st, 'a table inside a script is not handled')
        return
    if len(rows) == 0 or len(rows) > 250:
        fail(st, 'a table with %d rows is not handled' % len(rows))
        return
    cols = 0
    for r in rows:
        if bname(r) != 'mtr' or r.namespaceURI != NS:
            fail(st, r.nodeName + ' in a table is not handled')
            return
        cells = kids_of(r)
        if cells is None:
            fail(st, 'text outside a token element, in mtr')
            return
        if len(cells) == 0 or len(cells) > 250:
            fail(st, 'a table row with %d cells is not handled' % len(cells))
            return
        cols = max(cols, len(cells))
        for c in cells:
            if bname(c) != 'mtd' or c.namespaceURI != NS:
                fail(st, c.nodeName + ' in a table row is not handled')
                return
            if getattr_(c, 'rowspan') != '' or getattr_(c, 'columnspan') != '':
                fail(st, 'a table cell with rowspan or columnspan is not handled')
                return
    set_size(st, size)
    for b in (5, 0, 1, 1, 1, len(rows), cols):
        add(st, b)
    for _ in range((len(rows) + 4) // 4):
        add(st, 0)
    for _ in range((cols + 4) // 4):
        add(st, 0)
    for r in rows:
        cells = kids_of(r)
        for c in cells:
            ct = kids_of(c)
            if ct is None:
                fail(st, 'text outside a token element, in mtd')
                return
            put_line(st, ct, 0, len(ct) - 1, size)
            if st.why:
                return
        for _ in range(len(cells), cols):
            put_line(st, [], 0, -1, size)
    add(st, 0)


SYM = {0x2B: 0x2B, 0x3D: 0x3D, 0x3C: 0x3C, 0x3E: 0x3E, 0xB0: 0xB0, 0xB1: 0xB1, 0x2212: 0x2D,
       0x22C5: 0xD7, 0x2260: 0xB9, 0x2248: 0xBB, 0xD7: 0xB4, 0xF7: 0xB8, 0x2264: 0xA3,
       0x2265: 0xB3, 0x221E: 0xA5, 0x2220: 0xD0, 0x22A5: 0x5E, 0x2245: 0x40, 0x223C: 0x7E,
       0x2208: 0xCE, 0x2209: 0xCF, 0x2282: 0xCC, 0x2286: 0xCD, 0x222A: 0xC8, 0x2229: 0xC7,
       0x2190: 0xAC, 0x2191: 0xAD, 0x2192: 0xAE, 0x2193: 0xAF, 0x2194: 0xAB, 0x21D2: 0xDE,
       0x21D4: 0xDB, 0x2234: 0x5C, 0x221D: 0xB5, 0x2261: 0xBA, 0x2026: 0xBC, 0x2202: 0xB6,
       0x2207: 0xD1, 0x2211: 0xE5, 0x220F: 0xD5, 0x222B: 0xF2, 0x221A: 0xD6, 0xAC: 0xD8,
       0x2200: 0x22, 0x2203: 0x24, 0x2033: 0xB2, 0x2205: 0xC6, 0x2044: 0xA4,
       0x2135: 0xC0}
EXTRA = {0x25B3: 0x56, 0x2225: 0x50, 0x226A: 0x3D, 0x226B: 0x3F, 0x22EF: 0x4C, 0x2218: 0x6F,
         0x2213: 0x6D, 0x2243: 0x3B}
UP = "ABGDEZHQIKLMNXOPR?STUFCYW"
LO = "abgdezhqiklmnxoprVstufcyw"
# Sh_MTEF_Euclid_Place: (font, place) - typeface -11 Euclid Math One, -12 Euclid Math Two.
EUCLID_PLACE = {0x224D: (-11, 0xA9), 0x2115: (-12, 0x4E), 0x2124: (-12, 0x5A),
                0x211A: (-12, 0x51), 0x211D: (-12, 0x52), 0x2102: (-12, 0x43)}


def greek(code):
    if 0x391 <= code <= 0x3A9 and code != 0x3A2:
        return ord(UP[code - 0x391])
    if 0x3B1 <= code <= 0x3C9:
        return ord(LO[code - 0x3B1])
    return {0x3D1: 0x4A, 0x3D5: 0x66, 0x3D6: 0x76}.get(code, -1)


MOPUNCT = ".,()[]{}|_/:~!%'*#&$@;?\"\\·•"


def is_letter(c):
    return (0x41 <= c <= 0x5A) or (0x61 <= c <= 0x7A)


def put_token(st, node, size, emb):
    tag = bname(node)
    for k in node.childNodes:
        if k.nodeType == 1:
            fail(st, '%s holding an element (%s) is not handled' % (tag, k.nodeName))
            return
    txt = collapse(text(node))
    if txt == '':
        return
    v = collapse(getattr_(node, 'mathvariant'))
    if tag == 'mi' and v not in ('', 'normal', 'bold', 'italic', 'bold-italic'):
        fail(st, 'mi mathvariant="%s" is not handled' % v)
        return
    cu = u16(txt)
    allL = all(is_letter(c) for c in cu)
    res = []
    for code0 in cu:
        code = code0
        ch = chr(code0) if not 0xD800 <= code0 <= 0xDFFF else ''
        L = is_letter(code)
        tf, f8 = -2, -1
        if tag == 'mi':
            if 0x30 <= code <= 0x39:
                if v in ('', 'bold-italic', 'normal'):
                    tf = 8
            elif L:
                tf = {'bold-italic': -1, 'italic': 3, 'bold': 7, 'normal': 1}.get(v, 3 if len(cu) == 1 else 1)
            elif code == 0x20:
                tf, code = 1, 0xA0
            elif code == 0xA0:
                tf = 1
            elif code == 0x640:
                tf = 3
            elif greek(code) >= 0:
                if v == 'normal':
                    tf = 1
                elif v != 'italic':
                    f8 = greek(code)
                    tf = 5 if 0x391 <= code <= 0x3A9 else 4
            elif v == '' and ch and ch in '%$':
                tf = 2
            elif v == 'normal' and ch == '$':
                tf = 1
            elif v == '' and code in (0x221E, 0x2135):
                tf, f8 = 6, SYM[code]
            elif v == '' and code in EUCLID_PLACE:
                tf, f8 = EUCLID_PLACE[code]
            elif v == 'normal':
                if ch and ch in ',.:;?!&{}':
                    tf = 2
                elif code in (0x3D, 0x2B):
                    tf, f8 = 6, code
                elif code in (0x2D, 0x2013):
                    tf, code, f8 = 6, 0x2212, 0x2D
                elif code == 0x2003:
                    tf, code = 24, 0xEF05
                elif code == 0x200A:
                    tf, code = 24, 0xEF0A
        elif tag == 'mn':
            if 0x30 <= code <= 0x39:
                tf = 8
            elif ch in ('.', ','):
                tf = 2
            elif code == 0x2009:
                tf, code = 24, 0xEF02
            elif code == 0x2212:
                tf, f8 = 6, 0x2D
            elif code == 0x20:
                tf, code = 1, 0xA0
            elif code == 0x2026:
                tf, f8 = 6, 0xBC
            elif L:
                tf = 1
        elif tag == 'mo':
            if allL:
                tf = 3
            elif 0x2061 <= code <= 0x2063:
                tf = -3
            elif code in SYM and code != 0x2135:
                tf, f8 = 6, SYM[code]
            elif code in EXTRA:
                tf, f8 = 11, EXTRA[code]
            elif code in (0x2D, 0x2013, 0x2014):
                tf, code, f8 = 6, 0x2212, 0x2D
            elif code == 0x2032:
                tf, code = 2, 0x27
            elif code == 0x2217:
                tf, code = 2, 0x2A
            elif code in (0xA0, 0x2019):
                tf = 2
            elif code == 0x224D:
                tf, f8 = EUCLID_PLACE[code]
            elif ch and ch in MOPUNCT:
                tf = 2
        elif tag == 'mtext':
            if code == 0x2009:
                tf, code = 24, 0xEF02
            elif code == 0x2003:
                tf, code = 24, 0xEF05
            elif code == 0x200A:
                tf, code = 24, 0xEF0A
            elif code in (0x20, 0xA0):
                tf, code = 1, 0xA0
            elif code == 0x2212:
                tf, code = 1, 0x2D
            elif ((0x21 <= code <= 0x7E) or (0xA1 <= code <= 0xFF)
                  or code in (0x2018, 0x2019, 0x201C, 0x201D, 0x2026, 0x2022, 0x2013, 0x2014, 0x20AC)
                  or (0x391 <= code <= 0x3A1) or (0x3A3 <= code <= 0x3A9) or (0x3B1 <= code <= 0x3C9)):
                tf = 1
        if tf == -2:
            fail(st, 'the character U+%04X in %s%s is not handled' % (code0, tag, ' (%s)' % v if v else ''))
            return
        if tf != -3:
            res.append((tf, code, f8))
    if emb and len(res) != 1:
        fail(st, 'an accent on more than one character is not handled')
        return
    for tf, code, f8 in res:
        put_char(st, tf, code, f8, size, emb)


def walk(mathml):
    st = St()
    try:
        doc = minidom.parseString(mathml.encode('utf-8'))
    except Exception:
        st.why = 'the MathML could not be read'
        return st
    try:
        root = doc.documentElement
        if bname(root) != 'math' or root.namespaceURI != NS:
            st.why = 'it is not a MathML math element'
            return st
        k = kids_of(root)
        if k is None:
            st.why = 'text outside a token element'
            return st
        put_row(st, k, 0, len(k) - 1, 10)
    except Exception as e:      # a fault in this copy, not a refusal the VBA would give
        st.why = 'PYTHON ERROR %r' % e
    return st


# --- MTEF records: read MathType's own, and VistaType LP's, into comparable lists -----------------------

TF = {1: 'TEXT', 2: 'FUNC', 3: 'VAR', 4: 'LCGRK', 5: 'UCGRK', 6: 'SYM', 7: 'VEC', 8: 'NUM',
      11: 'MTEXTRA', 22: 'EXPAND', 24: 'SPACE'}
SET_ASIDE = {9, 10, 11, 12, 13, 14, 15, 16}     # sizes and colors


class Records:
    """The records of an MTEF body as short labels, one per record, nesting as END labels."""

    def __init__(self, b):
        self.b, self.p, self.out = b, 0, []
        self.colors = []          # (start, end) of every COLOR and COLOR_DEF record

    def u8(self):
        v = self.b[self.p]
        self.p += 1
        return v

    def u16(self):
        v = self.b[self.p] | (self.b[self.p + 1] << 8)
        self.p += 2
        return v

    def cstr(self):
        e = self.b.index(0, self.p)
        v = self.b[self.p:e].decode('latin1')
        self.p = e + 1
        return v

    def nudge(self, opt):
        if opt & 8:
            dx, dy = self.u8(), self.u8()
            if dx == 128 and dy == 128:
                self.u16()
                self.u16()
            return ' nudge'
        return ''

    def objs(self):
        while self.p < len(self.b):
            t = self.u8()
            if t == 0:
                self.out.append((0, 'END'))
                return
            self.rec(t)

    def ruler(self):
        n = self.u8()
        for _ in range(n):
            self.u8()
            self.u16()

    def rec(self, t):
        if t == 1:
            o = self.u8()
            nd = self.nudge(o)
            if o & 4:
                self.u16()
            if o & 2:
                self.ruler()
            if o & 1:
                self.out.append((1, 'LINE null' + nd))
                return
            self.out.append((1, 'LINE' + nd))
            self.objs()
        elif t == 2:
            o = self.u8()
            nd = self.nudge(o)
            tf = self.u8() - 128
            code = None if o & 0x20 else self.u16()
            f8 = self.u8() if o & 4 else None
            if o & 0x10:
                self.u16()
            s = 'CHAR %s U+%04X' % (TF.get(tf, str(tf)), code or 0)
            if f8 is not None:
                s += ' f8=%02X' % f8
            self.out.append((2, s + nd + (' [emb]' if o & 1 else '')))
            if o & 1:
                self.objs()
        elif t == 3:
            o = self.u8()
            nd = self.nudge(o)
            sel, var = self.u8(), self.u8()
            if var & 0x80:
                var = (var & 0x7F) | (self.u8() << 7)
            self.u8()
            self.out.append((3, 'TMPL %d/%d' % (sel, var) + nd))
            self.objs()
        elif t == 4:
            o = self.u8()
            nd = self.nudge(o)
            ha, va = self.u8(), self.u8()
            if o & 2:
                self.ruler()
            self.out.append((4, 'PILE %d %d' % (ha, va) + nd))
            self.objs()
        elif t == 5:
            o = self.u8()
            nd = self.nudge(o)
            va, hj, vj, r, c = (self.u8() for _ in range(5))
            for _ in range((2 * (r + 1) + 7) // 8 + (2 * (c + 1) + 7) // 8):
                self.u8()
            self.out.append((5, 'MATRIX %dx%d v%d h%d j%d' % (r, c, va, hj, vj) + nd))
            self.objs()
        elif t == 6:
            o = self.u8()
            nd = self.nudge(o)
            self.out.append((6, 'EMBELL %d' % self.u8() + nd))
        elif t == 7:
            self.ruler()
            self.out.append((7, 'RULER'))
        elif t == 8:
            fi = self.u8()
            if fi == 255:
                fi = self.u16()
            self.out.append((8, 'FONT_STYLE_DEF %d/%d' % (fi, self.u8())))
        elif t == 9:
            n = self.u8()
            if n == 101:
                self.u16()
            elif n == 100:
                self.u8()
                self.u16()
            else:
                self.u8()
            self.out.append((9, 'SIZE'))
        elif t in (10, 11, 12, 13, 14):
            self.out.append((t, {10: 'FULL', 11: 'SUB', 12: 'SUB2', 13: 'SYM', 14: 'SUBSYM'}[t]))
        elif t == 15:
            start = self.p - 1
            i = self.u8()
            if i == 255:
                self.u16()
            self.out.append((15, 'COLOR'))
            self.colors.append((start, self.p))
        elif t == 16:
            start = self.p - 1
            o = self.u8()
            for _ in range(4 if o & 1 else 3):
                self.u16()
            if o & 4:
                self.cstr()
            self.out.append((16, 'COLOR_DEF'))
            self.colors.append((start, self.p))
        elif t == 17:
            e = self.u8()
            if e == 255:
                e = self.u16()
            self.out.append((17, 'FONT_DEF %d %s' % (e, self.cstr())))
        elif t == 18:
            raise ValueError('EQN_PREFS inside the body')
        elif t == 19:
            self.out.append((19, 'ENCODING_DEF ' + self.cstr()))
        elif t >= 100:
            self.p += self.u16()
            self.out.append((t, 'FUTURE %d' % t))
        else:
            raise ValueError('unknown record %d at %d' % (t, self.p - 1))


PREAMBLE = bytes.fromhex(
    "050100070C44534D543700011357696E416C6C4261736963436F6465506167657300110554696D6573204E657720526F"
    "6D616E00110353796D626F6C001105436F7572696572204E65770011044D5420457874726100120008212F27F25F218F"
    "212F475F4150F21F1E4150F4150F4100F445F425F48F425F4100F4100F435F4100F21F20A5F20A25F48F21F4100F4100"
    "F40F48F417F48F4100F21A5F445F45F45F45F45F410F0C01000100010202020200020001010100030001000400000A01"
    "00")


def labels(body, one_column_letters=False):
    """Labels of an MTEF body that starts inside the opening LINE (as the walker's records do).

    one_column_letters: a letter directly in a cell of a one-column MATRIX is labeled LETTER,
    whatever its typeface, and a no-break space there NBSP. MathType's own conversion makes those
    letters plain italic whatever the book says, and such a space a FUNCTION character; the
    converter writes them as the book says (Jerry, 10/20/2026) - an intended difference.
    """
    r = Records(bytes(body))
    try:
        r.objs()
        while r.p < len(r.b):
            r.objs()
    except Exception as e:
        r.out.append((-1, 'UNREADABLE %r' % e))
    out, stack = [], []            # per open container: its label
    for t, s in r.out:
        in_cell = len(stack) >= 2 and stack[-1].startswith('LINE') and stack[-2].startswith('MATRIX') \
            and 'x1 ' in stack[-2]
        if t in (1, 3, 4, 5) and not s.startswith('LINE null'):
            stack.append(s)
        elif t == 2 and s.endswith(' [emb]'):
            stack.append(s)
        if t == 0 and stack:
            stack.pop()
        if t in SET_ASIDE:
            continue
        if one_column_letters and t == 2 and in_cell:
            m = re.match(r'CHAR (\S+) U\+00([4-7][0-9A-F]|A0)( \[emb\])?$', s)
            if m and m.group(2) == 'A0' and m.group(1) in ('TEXT', 'FUNC'):
                s = 'CHAR NBSP U+00A0'
            elif m and m.group(2) != 'A0' and chr(int(m.group(2), 16)).isalpha() \
                    and m.group(1) in ('TEXT', 'VAR', '-1', '-2', '-3', 'VEC'):
                s = 'CHAR LETTER U+00' + m.group(2) + (m.group(3) or '')
        out.append(s)
    return out


def without_colors(body):
    """MathType's records with its COLOR and COLOR_DEF records taken out, as bytes."""
    r = Records(bytes(body))
    try:
        r.objs()
        while r.p < len(r.b):
            r.objs()
    except Exception:
        return bytes(body)
    out, last = bytearray(), 0
    for a, b in r.colors:
        out += body[last:a]
        last = b
    return bytes(out + body[last:])


def mathtype_body(native):
    """MathType's Equation Native stream -> the records after the opening LINE of the equation.

    The header, encodings, fonts and preferences before it are skipped whatever they hold:
    Format Equations rewrites the preferences of every equation it touches.
    """
    m = native[native[0] | (native[1] << 8):]
    try:
        p = 5
        p = m.index(0, p) + 2                   # application key, then the options byte
        while True:
            t = m[p]
            p += 1
            if t == 19:                         # ENCODING_DEF
                p = m.index(0, p) + 1
            elif t == 17:                       # FONT_DEF
                p += 3 if m[p] == 255 else 1
                p = m.index(0, p) + 1
            elif t == 18:                       # EQN_PREFS: options, two dimension lists, styles
                p += 1
                for _ in range(2):
                    n, p, got = m[p], p + 1, 0
                    while got < n:
                        for nib in (m[p] >> 4, m[p] & 15):
                            if got < n and nib == 15:
                                got += 1
                        p += 1
                n, p = m[p], p + 1
                for _ in range(n):
                    p += 2 if m[p] else 1
            elif t in (10, 11, 12, 13, 14):
                pass
            elif t == 9:
                p += {100: 3, 101: 2}.get(m[p], 1) + 1
            elif t == 1 and m[p] == 0:          # the opening LINE
                return m[p + 1:]
            else:
                return None
    except (IndexError, ValueError):
        return None


def doc_equations(path):
    """Every equation of a .docx in order: ('O', None) for a Word equation, ('M', native) for a
    MathType object."""
    z = zipfile.ZipFile(path)
    doc = z.read('word/document.xml').decode('utf8')
    rels = z.read('word/_rels/document.xml.rels').decode('utf8')
    rmap = {attr(r, 'Id'): attr(r, 'Target') for r in re.findall(r'<Relationship\b[^>]*>', rels)}
    out = []
    for m in re.finditer(r'<m:oMath[ >]|<m:oMath/>|<o:OLEObject\b[^>]*>', doc):
        s = m.group(0)
        if s.startswith('<m:'):
            out.append(('O', None))
        elif attr(s, 'ProgID').startswith('Equation.DSMT'):
            o = olefile.OleFileIO(io.BytesIO(z.read('word/' + rmap[attr(s, 'r:id')])))
            out.append(('M', o.openstream('Equation Native').read()))
    return out


def first_difference(ours, theirs):
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ours, theirs, autojunk=False).get_opcodes():
        if tag != 'equal':
            a = ' | '.join(ours[i1:i2][:3]) or '-'
            b = ' | '.join(theirs[j1:j2][:3]) or '-'
            return 'VistaType LP: %s  //  MathType: %s' % (a, b)
    return ''


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('book')
    ap.add_argument('converted')
    ap.add_argument('--mixed')
    ap.add_argument('--show')
    ap.add_argument('--limit', type=int, default=5)
    ap.add_argument('--index', type=int)
    a = ap.parse_args(argv)

    frags = book_fragments(a.book)
    conv = doc_equations(a.converted)
    print('MathML equations in the book: %d; equations in the converted copy: %d' % (len(frags), len(conv)))
    if len(conv) != len(frags):
        print('  (the counts differ, so equations cannot be paired by position - comparison skipped)')
    focus = range(len(frags))
    if a.mixed:
        mixed = doc_equations(a.mixed)
        focus = [i for i, (k, _) in enumerate(mixed) if k == 'O']
        print('Word equations in --mixed (the set reported on): %d' % len(focus))

    if a.index is not None:
        i = a.index
        st = walk(frags[i])
        print(frags[i])
        print('refused:', st.why or '(accepted)')
        print('VistaType LP:', ' | '.join(labels(st.out + b'\x00\x00')))
        print('VistaType LP, hex:', (bytes(st.out) + b'\x00\x00').hex().upper())
        if len(conv) == len(frags) and conv[i][0] == 'M':
            print('MathType: ', ' | '.join(labels(mathtype_body(conv[i][1]) or b'')))
            print('MathType, hex, colors out:', without_colors(mathtype_body(conv[i][1]) or b'').hex().upper())
        return 0

    why = collections.Counter()
    same, exact, intended, kinds, examples = 0, 0, 0, collections.Counter(), collections.defaultdict(list)
    for i in focus:
        st = walk(frags[i])
        if st.why:
            why[st.why] += 1
            continue
        if len(conv) != len(frags) or conv[i][0] != 'M':
            continue
        body = mathtype_body(conv[i][1])
        if body is None:
            kinds['MathType equation could not be read'] += 1
            continue
        ours = labels(st.out + b'\x00\x00')    # the walker's records, the opening LINE's END, the last END
        theirs = labels(body)
        if ours == theirs:
            same += 1
            if bytes(st.out) + b'\x00\x00' == without_colors(body):
                exact += 1
            continue
        global OMML_ROUTE
        OMML_ROUTE = True
        ours = labels(walk(frags[i]).out + b'\x00\x00', True)
        OMML_ROUTE = False
        theirs = labels(body, True)
        if ours == theirs:
            intended += 1
        else:
            k = first_difference(ours, theirs)
            kinds[k] += 1
            examples[k].append(i)

    n = len(focus)
    print('\nAccepted: %d of %d; refused: %d' % (n - sum(why.values()), n, sum(why.values())))
    for k, v in why.most_common():
        print('  %5d  %s' % (v, k))
    if len(conv) == len(frags):
        print('\nOf the accepted, the same records as MathType (sizes and colors set aside): %d' % same)
        print('  of which byte for byte the same, once MathType\'s COLOR records are taken out: %d' % exact)
        print('The same once two intended differences are set aside - spaces from mspace, which')
        print('MathType\'s route drops, and letters and no-break spaces directly in a cell of a one-column')
        print('table, which it makes italic and FUNCTION: %d' % intended)
        print('Different: %d, by the first difference (with those two set aside):' % sum(kinds.values()))
        for k, v in kinds.most_common():
            print('  %5d  %s   e.g. #%s' % (v, k, ', #'.join(str(x) for x in examples[k][:4])))
    if a.show:
        for k in kinds:
            if a.show in k:
                for i in examples[k][:a.limit]:
                    print('\n#%d  %s' % (i, k))
                    print(frags[i])
                    print('VistaType LP:', ' | '.join(labels(walk(frags[i]).out + b'\x00\x00', True)))
                    print('MathType: ', ' | '.join(labels(mathtype_body(conv[i][1]), True)))
    return 0


if __name__ == '__main__':
    sys.exit(main())

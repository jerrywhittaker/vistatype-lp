#!/usr/bin/env python3
"""Linux-only VBA source reader for VistaType LP.

Reads LPandBRL.dotm (a zip), pulls word/vbaProject.bin (an OLE compound file),
decompresses the MS-OVBA-compressed module streams, and writes plain-text source
to reference/vba-src/ for reading, diffing, and code review.

This is a READ aid only. It does NOT produce import-ready files (UserForm .frx
binaries are not reconstructed). The authoritative source tree under src/ is
produced by tools/windows/Export-Vba.ps1 running in Word. See CLAUDE.md.

Usage:  python3 tools/lib/decompress_vba.py [path-to.dotm] [out-dir]
        defaults: LPandBRL.dotm  ->  reference/vba-src/

Requires: olefile  (pure-Python; no Office needed).
"""
import io
import os
import struct
import sys
import zipfile

try:
    import olefile
except ImportError:
    sys.exit("ERROR: needs the 'olefile' package (pip install olefile).")


def decompress(data: bytes) -> bytes:
    """MS-OVBA 2.4.1.3.1 DecompressStream."""
    if not data or data[0] != 0x01:
        raise ValueError("bad compressed-container signature")
    out = bytearray()
    i, n = 1, len(data)
    while i < n:
        header = struct.unpack("<H", data[i:i + 2])[0]
        i += 2
        size = (header & 0x0FFF) + 3
        compressed = (header >> 15) & 1
        chunk_end = i + size - 2
        dec_start = len(out)
        if not compressed:
            out.extend(data[i:i + 4096])
            i += 4096
            continue
        cur = i
        while cur < chunk_end and cur < n:
            flags = data[cur]
            cur += 1
            for bit in range(8):
                if cur >= chunk_end or cur >= n:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[cur])
                    cur += 1
                else:
                    token = struct.unpack("<H", data[cur:cur + 2])[0]
                    cur += 2
                    rel = len(out) - dec_start
                    x = rel - 1
                    b = 0
                    while x > 0:
                        x >>= 1
                        b += 1
                    bits = max(b, 4)
                    length = (token & (0xFFFF >> bits)) + 3
                    offset = (token >> (16 - bits)) + 1
                    start = len(out) - offset
                    for k in range(length):
                        out.append(out[start + k])
        i = chunk_end
    return bytes(out)


def module_offsets(dir_stream: bytes):
    """Yield (stream_name, text_offset) by scanning the decompressed dir stream
    for MODULESTREAMNAME (0x1A) and MODULEOFFSET (0x31) records."""
    d = decompress(dir_stream)
    i, n = 0, len(d)
    cur = None
    while i < n - 6:
        rid = struct.unpack("<H", d[i:i + 2])[0]
        size = struct.unpack("<I", d[i + 2:i + 6])[0]
        if rid == 0x1A and 0 < size < 64 and i + 6 + size <= n:
            name = d[i + 6:i + 6 + size]
            if all(32 <= c < 127 for c in name):
                cur = name.decode("latin-1")
                i += 6 + size
                continue
        if rid == 0x31 and size == 4:
            off = struct.unpack("<I", d[i + 6:i + 10])[0]
            if cur is not None:
                yield cur, off
                cur = None
            i += 10
            continue
        i += 1


def main():
    dotm = sys.argv[1] if len(sys.argv) > 1 else "LPandBRL.dotm"
    outdir = sys.argv[2] if len(sys.argv) > 2 else os.path.join("reference", "vba-src")
    os.makedirs(outdir, exist_ok=True)

    with zipfile.ZipFile(dotm) as z:
        bin_bytes = z.read("word/vbaProject.bin")
    ole = olefile.OleFileIO(io.BytesIO(bin_bytes))

    dir_stream = ole.openstream("VBA/dir").read()
    count = 0
    for stream, off in module_offsets(dir_stream):
        raw = ole.openstream("VBA/" + stream).read()
        # VBA stores module source in the system ANSI codepage (Windows-1252 for Western
        # installs), NOT ISO-8859-1: e.g. byte 0x91/0x92 are the smart quotes U+2018/U+2019,
        # 0x96 an en-dash, 0x85 an ellipsis. Decoding as latin-1 turns those into C1 control
        # codes. cp1252 recovers the intended punctuation.
        src = decompress(raw[off:]).decode("cp1252", errors="replace")
        with open(os.path.join(outdir, stream + ".vba"), "w", encoding="utf-8") as fh:
            fh.write(src)
        count += 1
    print(f"Wrote {count} modules to {outdir}/")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Generate icon.ico for COM Monitor — the app's port-list motif: stacked rows
with coloured status dots (green / amber / cyan) on a dark rounded tile. Pure
standard library: shapes are drawn with signed-distance fields (cheap per-pixel
anti-aliasing at every size) and the PNG/ICO containers are written by hand, so
no Pillow / external tools.

Run:  python make_icon.py            → writes icon.ico (+ prints an ASCII preview)
"""

import math
import struct
import zlib
import sys

# ── palette (matches the app) ───────────────────────────────────────────────
BG   = (0x16, 0x1a, 0x1d)        # dark tile
EDGE = (0x2b, 0x34, 0x3a)        # subtle tile border
BAR  = (0x39, 0x44, 0x4c)        # list-row bars
DOTS = ((0x7e, 0xc8, 0x7e),      # status dots, top→bottom: green / amber / cyan
        (0xff, 0xcc, 0x44),      #   (C_FREE / C_NEW / C_PORT)
        (0x61, 0xda, 0xfb))

SIZES = (256, 128, 64, 48, 32, 24, 16)


def clamp(v, a=0.0, b=1.0):
    return a if v < a else b if v > b else v


# ── signed-distance fields (centre-relative coords, +y down) ─────────────────
def sd_round_box(px, py, hw, hh, r):
    qx = abs(px) - hw + r
    qy = abs(py) - hh + r
    return math.hypot(max(qx, 0.0), max(qy, 0.0)) + min(max(qx, qy), 0.0) - r


def sd_circle(px, py, r):
    return math.hypot(px, py) - r


def _over(dst, color, a):
    """src-over: paint `color` at coverage `a` onto premultiplied-ish dst tuple
    (r,g,b,alpha) with straight alpha."""
    r, g, b, da = dst
    return (color[0] * a + r * (1 - a),
            color[1] * a + g * (1 - a),
            color[2] * a + b * (1 - a),
            a + da * (1 - a))


def render(S):
    cx = cy = S / 2.0
    hw = hh = 0.455 * S            # tile half-size
    rr = 0.235 * S                 # tile corner radius
    bw = 0.022 * S                 # tile border width

    gap    = 0.225 * S             # vertical spacing between the three rows
    dot_r  = 0.062 * S             # status-dot radius
    dot_x  = -0.250 * S            # dot centre (centre-relative)
    bar_x0, bar_x1 = -0.130 * S, 0.310 * S
    bar_cx = (bar_x0 + bar_x1) / 2.0
    bar_hw = (bar_x1 - bar_x0) / 2.0
    bar_hh = 0.052 * S

    img = bytearray(S * S * 4)
    for y in range(S):
        for x in range(S):
            px = x + 0.5 - cx
            py = y + 0.5 - cy
            d = (0.0, 0.0, 0.0, 0.0)

            # tile (rounded square) + border
            cov = clamp(0.5 - sd_round_box(px, py, hw, hh, rr))
            if cov > 0.0:
                d = _over(d, BG, cov)
                ring = clamp(0.5 - sd_round_box(px, py, hw - bw, hh - bw,
                                                max(rr - bw, 0.0)))
                d = _over(d, EDGE, cov * (1.0 - ring))

            # three rows: a status dot on the left, a list bar to its right
            for row in range(3):
                ry = py - (row - 1) * gap
                covb = clamp(0.5 - sd_round_box(px - bar_cx, ry,
                                                bar_hw, bar_hh, bar_hh))
                if covb > 0.0:
                    d = _over(d, BAR, covb)
                covd = clamp(0.5 - sd_circle(px - dot_x, ry, dot_r))
                if covd > 0.0:
                    d = _over(d, DOTS[row], covd)

            i = (y * S + x) * 4
            img[i + 0] = int(clamp(d[0], 0, 255))
            img[i + 1] = int(clamp(d[1], 0, 255))
            img[i + 2] = int(clamp(d[2], 0, 255))
            img[i + 3] = int(clamp(d[3] * 255.0, 0, 255))
    return bytes(img)


# ── PNG / ICO encoders ───────────────────────────────────────────────────────
def png_bytes(S, rgba):
    def chunk(typ, data):
        body = typ + data
        return (struct.pack(">I", len(data)) + body
                + struct.pack(">I", zlib.crc32(body) & 0xffffffff))

    ihdr = struct.pack(">IIBBBBB", S, S, 8, 6, 0, 0, 0)   # 8-bit RGBA
    raw = bytearray()
    stride = S * 4
    for y in range(S):
        raw.append(0)                                      # filter: none
        raw += rgba[y * stride:(y + 1) * stride]
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
            + chunk(b"IEND", b""))


def bmp_bytes(S, rgba):
    """Classic BMP/DIB icon image: BITMAPINFOHEADER + bottom-up 32-bit BGRA XOR
    bitmap + 1-bpp AND mask (1 = transparent, for legacy renderers)."""
    hdr = struct.pack("<IiiHHIIiiII", 40, S, S * 2, 1, 32, 0, 0, 0, 0, 0, 0)
    xor = bytearray()
    for y in range(S - 1, -1, -1):                          # bottom-up
        for x in range(S):
            i = (y * S + x) * 4
            xor += bytes((rgba[i + 2], rgba[i + 1], rgba[i + 0], rgba[i + 3]))
    stride = ((S + 31) // 32) * 4
    andm = bytearray()
    for y in range(S - 1, -1, -1):
        bits = bytearray(stride)
        for x in range(S):
            if rgba[(y * S + x) * 4 + 3] < 128:
                bits[x >> 3] |= 0x80 >> (x & 7)
        andm += bits
    return hdr + bytes(xor) + bytes(andm)


def ico_bytes(entries):
    """entries: list of (size, blob). Each blob is a PNG (256) or BMP/DIB."""
    out = struct.pack("<HHH", 0, 1, len(entries))
    offset = 6 + 16 * len(entries)
    blob = b""
    for size, png in entries:
        wh = 0 if size >= 256 else size
        out += struct.pack("<BBBBHHII", wh, wh, 0, 0, 1, 32, len(png), offset)
        blob += png
        offset += len(png)
    return out + blob


def ascii_preview(S=44):
    rgba = render(S)
    lines = []
    for y in range(S):
        row = []
        for x in range(S):
            i = (y * S + x) * 4
            a = rgba[i + 3]
            tot = rgba[i] + rgba[i + 1] + rgba[i + 2]
            row.append(" " if a < 100 else
                       "#" if tot > 300 else      # status dot
                       "+" if tot > 150 else      # list bar
                       ".")                        # tile
        lines.append("".join(row))
    return "\n".join(lines)


if __name__ == "__main__":
    print(ascii_preview())
    entries = []
    for s in SIZES:
        rgba = render(s)
        entries.append((s, png_bytes(s, rgba) if s >= 128 else bmp_bytes(s, rgba)))
    with open("icon.ico", "wb") as f:
        f.write(ico_bytes(entries))
    print(f"\nwrote icon.ico ({sum(len(p) for _, p in entries)} bytes, "
          f"sizes {', '.join(str(s) for s in SIZES)})")

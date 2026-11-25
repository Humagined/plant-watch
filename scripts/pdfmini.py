"""A minimal PDF writer (text, lines, fills, links; the report's own fonts; US Letter). Standard library only.

Text is real text in a WinAnsi-encoded font, so a reader can extract the figures.
"""

import json
import zlib
from pathlib import Path

FONTS = Path(__file__).resolve().parent.parent / "assets" / "fonts"
METRICS = json.loads((FONTS / "metrics.json").read_text(encoding="utf-8"))
REGULAR, BOLD, DISPLAY = "text-400", "text-600", "display-800"
ORDER = (REGULAR, BOLD, DISPLAY)


def winansi(text):
    """What the fonts can show: keeps › – “ ” €, and an unknown character becomes ?."""
    return str(text).encode("cp1252", "replace")


def font_of(bold):
    return bold if bold in ORDER else BOLD if bold else REGULAR


def width(text, size, bold=False):
    w = METRICS[font_of(bold)]["widths"]
    return sum(w[b - 32] if b >= 32 else 0 for b in winansi(text)) * size / 1000


def esc(text):
    t = winansi(text).decode("cp1252", "ignore")
    return t.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def wrap(text, size, bold, limit):
    """Lines that fit. A word longer than the line (a path, a formula) is broken."""
    lines, cur = [], ""
    for word in str(text).split():
        trial = f"{cur} {word}" if cur else word
        if width(trial, size, bold) <= limit:
            cur = trial
            continue
        if cur:
            lines.append(cur)
        cur = word
        while width(cur, size, bold) > limit and len(cur) > 1:
            n = len(cur)
            while n > 1 and width(cur[:n], size, bold) > limit:
                n -= 1
            lines.append(cur[:n])
            cur = cur[n:]
    return lines + [cur] if cur else lines


def _rgb(color):
    return " ".join(f"{int(color[i:i + 2], 16) / 255:.3f}" for i in (1, 3, 5))


class Page:
    def __init__(self):
        self.ops, self.links = [], []

    def text(self, x, y, s, size=10, bold=False, right=False, color="#0E1A2B"):
        if right:
            x -= width(s, size, bold)
        f = ORDER.index(font_of(bold)) + 1
        self.ops.append(f"BT {_rgb(color)} rg /F{f} {size} Tf {x:.1f} {y:.1f} Td ({esc(s)}) Tj ET")

    def line(self, x1, y1, x2, y2, w=0.5, color="#DBDBDB"):
        self.ops.append(f"{_rgb(color)} RG {w} w {x1:.1f} {y1:.1f} m {x2:.1f} {y2:.1f} l S")

    def rect(self, x, y, w, h, color="#FFFFFF", border=None):
        self.ops.append(f"{_rgb(color)} rg {x:.1f} {y:.1f} {w:.1f} {h:.1f} re f")
        if border:
            self.ops.append(f"{_rgb(border)} RG 0.6 w {x:.1f} {y:.1f} {w:.1f} {h:.1f} re S")

    def link(self, x, y, w, h, url):
        self.links.append((x, y, w, h, url))


def pdf(pages):
    objs = [None, None]

    def add(body):
        objs.append(body if isinstance(body, bytes) else body.encode("cp1252"))
        return len(objs)

    def stream(data, extra=""):
        z = zlib.compress(data, 9)
        return f"<< /Length {len(z)} /Filter /FlateDecode {extra}>>\nstream\n".encode() + z + b"\nendstream"

    fonts = []
    for name in ORDER:
        m, data = METRICS[name], (FONTS / f"{name}.ttf").read_bytes()
        file = add(stream(data, f"/Length1 {len(data)} "))
        desc = add(f"<< /Type /FontDescriptor /FontName /{m['name']} /Flags 32 /FontBBox [{' '.join(map(str, m['bbox']))}] "
                   f"/ItalicAngle 0 /Ascent {m['ascent']} /Descent {m['descent']} /CapHeight {m['cap']} /StemV 90 "
                   f"/FontFile2 {file} 0 R >>")
        fonts.append(add(f"<< /Type /Font /Subtype /TrueType /BaseFont /{m['name']} /Encoding /WinAnsiEncoding "
                         f"/FirstChar 32 /LastChar 255 /Widths [{' '.join(map(str, m['widths']))}] /FontDescriptor {desc} 0 R >>"))
    res = " ".join(f"/F{i} {n} 0 R" for i, n in enumerate(fonts, 1))
    kids = []
    for pg in pages:
        content = add(stream("\n".join(pg.ops).encode("cp1252")))
        annots = [add(f"<< /Type /Annot /Subtype /Link /Rect [{x:.1f} {y:.1f} {x + w:.1f} {y + h:.1f}] /Border [0 0 0] "
                      f"/A << /S /URI /URI ({esc(u)}) >> >>") for x, y, w, h, u in pg.links]
        extra = f" /Annots [{' '.join(f'{a} 0 R' for a in annots)}]" if annots else ""
        kids.append(add(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents {content} 0 R "
                        f"/Resources << /Font << {res} >> >>{extra} >>"))
    objs[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objs[1] = f"<< /Type /Pages /Kids [{' '.join(f'{k} 0 R' for k in kids)}] /Count {len(kids)} >>".encode()
    buf, offsets = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"), []
    for i, o in enumerate(objs, 1):
        offsets.append(len(buf))
        buf += f"{i} 0 obj\n".encode() + o + b"\nendobj\n"
    xref = len(buf)
    buf += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    buf += "".join(f"{o:010d} 00000 n \n" for o in offsets).encode()
    buf += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(buf)

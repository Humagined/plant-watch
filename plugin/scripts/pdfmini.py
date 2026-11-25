"""A minimal PDF writer (text, lines, fills, links; Helvetica; US Letter). Standard library only.

Text is real text in a WinAnsi-encoded font, so a reader can extract the figures.
"""

import zlib

_REG = ("278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 278 278 584 584 "
        "584 556 1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 667 778 722 667 611 722 667 944 667 667 611 278 "
        "278 278 469 556 333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 556 556 333 500 278 556 500 722 500 500 "
        "500 334 260 334 584")
_BOLD = ("278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 333 333 584 584 "
         "584 611 975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 667 778 722 667 611 722 667 944 667 667 611 333 "
         "278 333 584 556 333 556 611 556 611 556 333 611 611 278 278 556 278 889 611 611 611 611 389 556 333 611 556 778 556 556 "
         "500 389 280 389 584")
# Punctuation outside ASCII the reports use, the same in both weights to within a few units
_MORE = {"›": 333, "‹": 333, "–": 556, "—": 1000, "‘": 250, "’": 250, "“": 400, "”": 400, "•": 350, "…": 1000, "×": 584,
         "÷": 584, "°": 400, "€": 556}


def _widths(table):
    w = {32 + i: int(v) for i, v in enumerate(table.split())}
    w.update({ch.encode("cp1252")[0]: v for ch, v in _MORE.items()})
    return w


REGULAR, BOLD, DISPLAY = "Helvetica", "Helvetica-Bold", "Helvetica-Bold"
ORDER = ("Helvetica", "Helvetica-Bold")
WIDTHS = {"Helvetica": _widths(_REG), "Helvetica-Bold": _widths(_BOLD)}


def winansi(text):
    """What the fonts can show: keeps › – “ ” €, and an unknown character becomes ?."""
    return str(text).encode("cp1252", "replace")


def font_of(bold):
    return bold if bold in ORDER else BOLD if bold else REGULAR


def width(text, size, bold=False):
    w = WIDTHS[font_of(bold)]
    return sum(w.get(b, 556) for b in winansi(text)) * size / 1000


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

    def stream(data):
        z = zlib.compress(data, 9)
        return f"<< /Length {len(z)} /Filter /FlateDecode >>\nstream\n".encode() + z + b"\nendstream"

    fonts = [add(f"<< /Type /Font /Subtype /Type1 /BaseFont /{name} /Encoding /WinAnsiEncoding >>") for name in ORDER]
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

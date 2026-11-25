"""A minimal PDF writer (text, lines, fills; Helvetica; US Letter). Standard library only.

Used for the printable report. Text is real text, so a reader can extract every figure.
"""

_HELV = ("278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 "
         "556 556 278 278 584 584 584 556 1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 "
         "667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 333 556 556 500 556 556 278 556 "
         "556 222 222 500 222 833 556 556 556 556 333 500 278 556 500 722 500 500 500 334 260 334 584")
W = {chr(32 + i): int(v) for i, v in enumerate(_HELV.split())}
# WinAnsi punctuation the reports use (Helvetica AFM widths)
W.update({"›": 333, "‹": 333, "–": 556, "—": 1000, "‘": 222, "’": 222, "“": 333, "”": 333, "•": 350,
          "…": 1000, "×": 584, "÷": 584, "°": 400, "€": 556})


def width(text, size, bold=False):
    return sum(W.get(ch, 556) for ch in text) * size / 1000 * (1.06 if bold else 1)


def esc(text):
    t = str(text).encode("cp1252", "replace").decode("cp1252")  # WinAnsi, the fonts' encoding: keeps › – “ ” €
    return t.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


class Page:
    def __init__(self):
        self.ops = []

    def text(self, x, y, s, size=10, bold=False, right=False, gray=0.0):
        if right:
            x -= width(s, size, bold)
        self.ops.append(f"BT {gray:.2f} g /{'F2' if bold else 'F1'} {size} Tf {x:.1f} {y:.1f} Td ({esc(s)}) Tj ET")

    def line(self, x1, y1, x2, y2, w=0.5, gray=0.75):
        self.ops.append(f"{gray:.2f} G {w} w {x1:.1f} {y1:.1f} m {x2:.1f} {y2:.1f} l S")

    def rect(self, x, y, w, h, gray=0.93):
        self.ops.append(f"{gray:.2f} g {x:.1f} {y:.1f} {w:.1f} {h:.1f} re f")


def pdf(pages):
    objs = [None, None,
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"]
    kids = []
    for pg in pages:
        stream = "\n".join(pg.ops).encode("cp1252")
        objs.append(b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream")
        objs.append(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents {len(objs)} 0 R "
                    f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> >>".encode())
        kids.append(len(objs))
    objs[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objs[1] = f"<< /Type /Pages /Kids [{' '.join(f'{k} 0 R' for k in kids)}] /Count {len(kids)} >>".encode()
    buf, offsets = bytearray(b"%PDF-1.4\n"), []
    for i, o in enumerate(objs, 1):
        offsets.append(len(buf))
        buf += f"{i} 0 obj\n".encode() + o + b"\nendobj\n"
    xref = len(buf)
    buf += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    buf += "".join(f"{o:010d} 00000 n \n" for o in offsets).encode()
    buf += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(buf)

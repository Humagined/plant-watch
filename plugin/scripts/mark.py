"""The Humagined mark, drawn from lines and shapes: an SVG for the web report, drawing steps for the PDF."""

SIZE = (1988, 1875)
BOXES = [((0, 0, 1988, 1875), "#000000"), ((71, 71, 1917, 1804), "#FFFFFF"), ((134, 134, 1853, 1740), "#000000"),
         ((665, 538, 827, 1189), "#FFFFFF"), ((1075, 538, 1238, 1189), "#FFFFFF"), ((827, 807, 1075, 927), "#FFFFFF")]
WORD = ("humagined.", 356, 1539, 240, 1318)            # text, left, baseline, size, width
STROKE = 10
_P = {"a": (1464, 690), "b": (1575, 700), "c": (1646, 735), "d": (1689, 813), "e": (1684, 902), "f": (1663, 946),
      "g": (1610, 976), "h": (1549, 978), "i": (1510, 1006), "j": (1483, 991), "k": (1480, 970), "l": (1468, 911),
      "m": (1407, 876), "n": (1371, 850), "o": (1330, 809), "p": (1336, 761), "q": (1358, 726), "r": (1407, 699),
      "s": (1416, 741), "t": (1384, 777), "u": (1433, 812), "v": (1441, 848), "w": (1493, 866), "x": (1499, 729),
      "y": (1503, 778), "z": (1540, 785), "A": (1549, 850), "B": (1621, 786), "C": (1611, 875), "D": (1565, 923),
      "E": (1492, 923), "F": (1632, 926), "G": (1539, 962)}
_EDGES = ("ab bc cd de ef fg gh hi ij jk kl lm mn no op pq qr ra rs sa sx ax xb xy yz zb zB bB cB Bd qt pt ot st tu ou uy uv nv "
          "vm vw mw wy wA zA AB BC dC Ce AC AD CD CF DF Fe Ff Fg DE Ew El Ek DG Gg Gh").split()
LINES = [(*_P[a], *_P[b], STROKE) for a, b in _EDGES] + [(1322, 1047, 1416, 1181, 24), (1416, 1047, 1322, 1181, 24)]


def svg(width):
    w, h = SIZE
    out = [f"<svg role='img' aria-label='Humagined' width='{width}' height='{round(width * h / w)}' viewBox='0 0 {w} {h}'>"]
    out += [f"<rect x='{x1}' y='{y1}' width='{x2 - x1}' height='{y2 - y1}' fill='{c}'/>" for (x1, y1, x2, y2), c in BOXES]
    out += [f"<path d='M{x1} {y1}L{x2} {y2}' stroke='#FFFFFF' stroke-width='{n}'/>" for x1, y1, x2, y2, n in LINES]
    t, x, y, size, tw = WORD
    out.append(f"<text x='{x}' y='{y}' font-family='Helvetica,Arial,sans-serif' font-weight='700' font-size='{size}' "
               f"textLength='{tw}' fill='#FFFFFF'>{t}</text></svg>")
    return "".join(out)


def draw(page, x, y, width):
    """On a PDF page, with the mark's lower left corner at (x, y)."""
    k = width / SIZE[0]
    top = y + SIZE[1] * k
    for (x1, y1, x2, y2), color in BOXES:
        page.rect(x + x1 * k, top - y2 * k, (x2 - x1) * k, (y2 - y1) * k, color)
    for x1, y1, x2, y2, n in LINES:
        page.line(x + x1 * k, top - y1 * k, x + x2 * k, top - y2 * k, max(0.25, n * k), "#FFFFFF")
    t, tx, ty, size, _ = WORD
    page.text(x + tx * k, top - ty * k, t, round(size * k, 1), bold=True, color="#FFFFFF")

"""Best-effort PDF text, one list of lines per page. Standard library only.

Used to verify a model's extraction (is the quoted line really on the cited page?), not to
extract figures. Handles uncompressed and Flate-compressed content streams with simple
fonts; returns None for a page it can't read, so callers say "couldn't verify" instead of guessing.
"""

import re
import zlib
from pathlib import Path

import safeio

OBJ_RE = re.compile(rb"(\d+)\s+0\s+obj(.*?)endobj", re.S)


def _string(raw):
    out, i = bytearray(), 0
    while i < len(raw):
        ch = raw[i]
        if ch == 0x5C and i + 1 < len(raw):  # backslash
            nx = raw[i + 1]
            esc = {ord("n"): 10, ord("r"): 13, ord("t"): 9, ord("b"): 8, ord("f"): 12}
            if nx in esc:
                out.append(esc[nx]); i += 2; continue
            if 0x30 <= nx <= 0x37:
                m = re.match(rb"[0-7]{1,3}", raw[i + 1:i + 4])
                out.append(int(m.group(0), 8) & 255); i += 1 + len(m.group(0)); continue
            out.append(nx); i += 2; continue
        out.append(ch); i += 1
    # Simple fonts are almost always WinAnsi (cp1252): 0x80-0x9F are – — ‘ ’ “ ” › and friends, not controls.
    return "".join(bytes([b]).decode("cp1252", "ignore") or chr(b) for b in out)


def _text_runs(content):
    """[(x, y, text)] from BT..ET blocks using Td/Tm positions and Tj/TJ strings."""
    runs = []
    for block in re.findall(rb"BT(.*?)ET", content, re.S):
        x = y = 0.0
        tm = re.search(rb"([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm", block)
        if tm:
            x, y = float(tm.group(5)), float(tm.group(6))
        for td in re.finditer(rb"([-\d.]+)\s+([-\d.]+)\s+T[dD]", block):
            x += float(td.group(1)); y += float(td.group(2))
        parts = [_string(m.group(1)) for m in re.finditer(rb"\(((?:\\.|[^\\)])*)\)\s*Tj", block, re.S)]
        for arr in re.finditer(rb"\[(.*?)\]\s*TJ", block, re.S):
            parts.append("".join(_string(m.group(1)) for m in re.finditer(rb"\(((?:\\.|[^\\)])*)\)", arr.group(1), re.S)))
        if parts:
            runs.append((x, y, "".join(parts)))
    return runs


def pages(path):
    safeio.check_size(path)
    data = Path(path).read_bytes()
    objs = {int(m.group(1)): m.group(2) for m in OBJ_RE.finditer(data)}

    def stream(num):
        body = objs.get(num, b"")
        m = re.search(rb"stream\r?\n(.*?)\r?\nendstream", body, re.S)
        if not m:
            return b""
        raw = m.group(1)
        if b"/FlateDecode" in body[:m.start()]:
            try:
                raw = safeio.inflate(raw)
            except (zlib.error, safeio.Refused):
                return None
        return raw

    page_nums = [n for n, b in sorted(objs.items()) if re.search(rb"/Type\s*/Page(?!s)", b)]
    root = next((b for b in objs.values() if re.search(rb"/Type\s*/Pages", b) and b"/Kids" in b), None)
    if root:
        kids = [int(k) for k in re.findall(rb"(\d+)\s+0\s+R", re.search(rb"/Kids\s*\[(.*?)\]", root, re.S).group(1))]
        page_nums = [k for k in kids if k in objs] or page_nums
    out = []
    for pn in page_nums:
        body = objs[pn]
        cm = re.search(rb"/Contents\s*(\[(.*?)\]|(\d+)\s+0\s+R)", body, re.S)
        refs = [int(x) for x in re.findall(rb"(\d+)\s+0\s+R", cm.group(2))] if cm and cm.group(2) else \
               ([int(cm.group(3))] if cm else [])
        content = b""
        unreadable = False
        for r in refs:
            s = stream(r)
            if s is None:
                unreadable = True
            else:
                content += s + b"\n"
        runs = _text_runs(content)
        if unreadable and not runs:
            out.append(None)
            continue
        lines = {}
        for x, y, t in runs:
            key = round(y)
            lines.setdefault(key, []).append((x, t))
        out.append([" ".join(t for _, t in sorted(v)) for _, v in sorted(lines.items(), key=lambda kv: -kv[0])])
    return out


def norm(s):
    return re.sub(r"\s+", " ", s.replace(",", "")).strip().lower()


def verify(path, page, quote):
    """True / False if the quoted text is / isn't on the page; None if the page can't be read."""
    ps = pages(path)
    if page < 1 or page > len(ps) or ps[page - 1] is None:
        return None
    text = norm(" \n ".join(ps[page - 1]))
    return norm(quote) in text

"""Where Plant Watch may write, how it writes, and how much it will read. Standard library only.

The plant's own files are never written. A result may go anywhere outside the plant folder, or inside it
only under a folder named `plant-watch`. An existing file is replaced only if Plant Watch wrote it and
nobody has changed it since.
"""

import hashlib
import json
import os
import tempfile
import zipfile
import zlib
from pathlib import Path

OWN_FOLDER = "plant-watch"
MAX_FILE = 200 * 1024 * 1024            # a file as it sits on disk
MAX_PART = 300 * 1024 * 1024            # one part of a workbook, unpacked
MAX_BOOK = 800 * 1024 * 1024            # a whole workbook, unpacked
MAX_PDF_STREAM = 50 * 1024 * 1024       # one PDF stream, unpacked


class Refused(ValueError):
    """A read or write Plant Watch won't do. The message is for the person."""


def read_text(path):
    return Path(path).read_text(encoding="utf-8")


def within(path, folder):
    try:
        Path(path).resolve().relative_to(Path(folder).resolve())
        return True
    except ValueError:
        return False


def check_place(path, folder):
    """Refuse a place among the plant's own files."""
    if not folder or not within(path, folder):
        return
    rel = Path(path).resolve().relative_to(Path(folder).resolve())
    # The plant folder may itself be the plant-watch working folder's parent, never the other way round.
    if OWN_FOLDER not in rel.parts:
        raise Refused(f"{path} is among the plant's own files. Save results outside {folder}, "
                      f"or in a folder named {OWN_FOLDER} inside it.")


def _digest(data):
    return hashlib.sha256(data).hexdigest()


def _atomic(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, prefix=".plant-watch-")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        try:
            os.replace(tmp, path)
        except PermissionError:
            raise Refused(f"{path.name} couldn't be replaced. If it's open in another program, close it and try again.")
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise


def save_state(path, text):
    """A file in the state folder that belongs to Plant Watch by definition."""
    _atomic(path, text.encode("utf-8"))


def settings(state):
    p = Path(state) / "settings.json"
    return json.loads(read_text(p)) if p.exists() else {}


def remember(state, cfg):
    """Keep the plant folder and name beside the ledger, so every later command knows what not to write."""
    s = settings(state)
    s.update({"plant": cfg.get("plant", ""), "folder": str(Path(cfg["folder"]).resolve()) if cfg.get("folder") else ""})
    save_state(Path(state) / "settings.json", json.dumps(s, indent=1))


def save_result(state, path, data, sources=()):
    """Write a report or export. `data` is text or bytes. `sources` are files the person pointed us at."""
    if isinstance(data, str):
        data = data.encode("utf-8")
    path = Path(path)
    check_place(path, settings(state).get("folder"))
    real = path.resolve()
    if any(real == Path(s).resolve() for s in sources):
        raise Refused(f"{path} is one of the plant's own files. Choose another name.")
    index = Path(state) / "outputs.json"
    ours = json.loads(read_text(index)) if index.exists() else {}
    if path.exists():
        if not path.is_file() or ours.get(str(real)) != _digest(path.read_bytes()):
            raise Refused(f"{path} already exists and isn't an unchanged Plant Watch file. Choose another name.")
    _atomic(path, data)
    ours[str(real)] = _digest(data)
    save_state(index, json.dumps(ours, indent=1))


def check_size(path):
    size = os.path.getsize(path)
    if size > MAX_FILE:
        raise Refused(f"{Path(path).name} is {size // 1048576:,} MB. Plant Watch reads files up to {MAX_FILE // 1048576} MB.")


def open_book(path):
    check_size(path)
    z = zipfile.ZipFile(path)
    if sum(i.file_size for i in z.infolist()) > MAX_BOOK:
        z.close()
        raise Refused(f"{Path(path).name} unpacks to more than Plant Watch will read.")
    return z


def book_part(z, name):
    """One part of a workbook. XML that declares a DTD is refused: Excel never writes one."""
    with z.open(name) as f:
        data = f.read(MAX_PART + 1)
    if len(data) > MAX_PART:
        raise Refused(f"{name} unpacks to more than Plant Watch will read.")
    head = data[:4096].replace(b"\x00", b"").upper()
    if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
        raise Refused(f"{name} declares a document type, which a workbook never does.")
    return data


def inflate(raw):
    d = zlib.decompressobj()
    out = d.decompress(raw, MAX_PDF_STREAM)
    if d.unconsumed_tail:
        raise Refused("A PDF stream unpacks to more than Plant Watch will read.")
    return out


def csv_safe(v):
    """Text a spreadsheet would run as a formula is shown as text instead."""
    if isinstance(v, str) and v.lstrip(" \t\r\n")[:1] in ("=", "+", "-", "@") or isinstance(v, str) and v[:1] in ("\t", "\r"):
        return "'" + v
    return v

"""What the README promises, held to by tests."""
import ast
import csv
import json
import shutil
import subprocess
import sys
import zipfile
import zlib
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
import northfork as nf  # noqa: E402
import definitions  # noqa: E402
import pdftext  # noqa: E402
import report  # noqa: E402
import safeio  # noqa: E402
import sources as S  # noqa: E402
import watch  # noqa: E402
import xlsx_audit as xa  # noqa: E402

DEMO = ROOT / "demo" / "northfork-precision"
NETWORK = {"socket", "ssl", "http", "urllib", "ftplib", "smtplib", "poplib", "imaplib", "telnetlib", "xmlrpc", "asyncio",
           "requests", "httpx", "aiohttp", "webbrowser", "subprocess", "ctypes", "multiprocessing", "importlib", "pty"}


def digests(folder):
    import hashlib
    return {p.relative_to(folder).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in sorted(folder.rglob("*")) if p.is_file()}


@pytest.fixture
def demo(tmp_path):
    base = tmp_path / "shop"
    shutil.copytree(DEMO, base)
    cfg = S.load_config(nf.config(base, base / "plant"))
    state = tmp_path / "work" / "plant-watch" / "state"
    state.mkdir(parents=True)
    shutil.copy(ROOT / "tests" / "fixtures" / "northfork_dictionary.json", state / "dictionary.json")
    nf.read_reports(base / "plant", state)
    return base, cfg, state


def test_the_shipped_code_imports_nothing_that_can_reach_a_network_or_start_a_program():
    for f in (ROOT / "scripts").glob("*.py"):
        for node in ast.walk(ast.parse(f.read_text(encoding="utf-8"))):
            names = [a.name for a in node.names] if isinstance(node, ast.Import) else \
                [node.module or ""] if isinstance(node, ast.ImportFrom) else []
            assert not {n.split(".")[0] for n in names} & NETWORK, f
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
                assert node.func.id not in ("eval", "exec", "__import__", "compile"), f


def test_a_full_run_on_the_demo_changes_no_source_and_ends_with_the_booking_line(demo, tmp_path):
    base, cfg, state = demo
    before = digests(base)
    watch.run(cfg, state, date(2026, 11, 25))
    out = state.parent
    report.render(cfg, state, out / "r.html", str(out / "r.pdf"))
    r = watch.export(state, out / "findings.csv")
    assert {k: v for k, v in digests(base).items() if k in before} == before
    assert r["count"] == 39 and r["definitions"] == 3 and r["problems"] == []
    assert "tesserex" not in json.dumps(r).lower()
    line = ("Found mismatches? Book a free 30-minute working session to walk through them: "
            "https://www.humagined.ai/book/bryce (through January 31, 2027).")
    import re
    assert line in re.sub(r"<[^>]+>", "", (out / "r.html").read_text(encoding="utf-8"))
    pages = pdftext.pages(out / "r.pdf")
    assert line in " ".join(" ".join(p) for p in pages[-2:] if p).replace("  ", " ")
    doc = json.loads((out / "plant-watch-definitions.json").read_text(encoding="utf-8"))
    assert doc["format"] == definitions.FORMAT and {d["key"] for d in doc["definitions"]} == {"downtime", "scrap", "good"}


def test_the_stale_report_is_found_whatever_the_file_time(demo):
    import os
    base, cfg, state = demo
    os.utime(base / "linked" / "daily_production_summary.xlsx")
    watch.run(cfg, state, date(2026, 11, 25))
    import ledger
    assert any(e["kind"] == "stale_link" for e in ledger.load(state).values())


def test_results_never_land_on_the_plants_files(demo, tmp_path):
    base, cfg, state = demo
    watch.run(cfg, state, date(2026, 11, 25))
    items = base / "plant" / "erp" / "items.csv"
    for target in (items, base / "plant" / "new.csv", base / "plant" / "erp" / ".." / "x.csv"):
        with pytest.raises(safeio.Refused):
            watch.export(state, target)
    with pytest.raises(safeio.Refused):
        report.render(cfg, state, items)
    with pytest.raises(safeio.Refused):
        report.render(cfg, state, base / "linked" / "daily_production_summary.xlsx")
    with pytest.raises(safeio.Refused):
        watch.run(cfg, base / "plant" / "exports" / "state")
    try:
        (tmp_path / "link.csv").symlink_to(items)
    except OSError:
        return
    with pytest.raises(safeio.Refused):
        watch.export(state, tmp_path / "link.csv")


def test_someone_elses_file_and_an_edited_export_are_not_replaced(demo):
    _, cfg, state = demo
    watch.run(cfg, state, date(2026, 11, 25))
    out = state.parent
    (out / "mine.csv").write_text("notes", encoding="utf-8")
    with pytest.raises(safeio.Refused):
        watch.export(state, out / "mine.csv")
    watch.export(state, out / "f.csv")
    watch.export(state, out / "f.csv")                        # unchanged: replaced
    (out / "f.csv").write_text("my edits", encoding="utf-8")
    with pytest.raises(safeio.Refused):
        watch.export(state, out / "f.csv")


def test_a_plant_watch_folder_inside_the_plant_folder_is_allowed(demo):
    base, cfg, _ = demo
    state = base / "plant" / "plant-watch" / "state"
    watch.run(cfg, state, date(2026, 11, 25))
    assert watch.export(state, base / "plant" / "plant-watch" / "f.csv")["ok"]
    assert not any("plant-watch" in g["pattern"] for g in watch.scan(base / "plant")["groups"])


def test_the_csv_holds_nothing_a_spreadsheet_would_run(demo):
    _, cfg, state = demo
    watch.run(cfg, state, date(2026, 11, 25))
    out = state.parent / "f.csv"
    watch.export(state, out)
    cells = [v for r in csv.DictReader(out.open(encoding="utf-8-sig", newline="")) for v in r.values() if v]
    assert any(v.startswith("'=") for v in cells)
    assert not [v for v in cells if v[0] in "=+@\t\r" or (v[0] == "-" and not v[1:].replace(".", "").isdigit())]


@pytest.mark.parametrize("pattern", ["../outside/*.csv", "/etc/hosts", "a/../../x.csv", "C:\\x.csv", "..\\x.csv"])
def test_a_source_pattern_stays_in_the_plant_folder(tmp_path, pattern):
    with pytest.raises(safeio.Refused):
        S.files_for({"folder": str(tmp_path)}, {"id": "s", "glob": pattern})


def test_hostile_files_are_refused_not_unpacked(tmp_path, monkeypatch):
    monkeypatch.setattr(safeio, "MAX_PART", 1000)
    book = tmp_path / "b.xlsx"
    with zipfile.ZipFile(book, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/_rels/workbook.xml.rels", "x" * 5000)
    with pytest.raises(safeio.Refused):
        xa.Workbook(book)
    with zipfile.ZipFile(book, "w") as z:
        z.writestr("xl/_rels/workbook.xml.rels", "<!DOCTYPE l [<!ENTITY a 'b'>]><l>&a;</l>")
    with pytest.raises(safeio.Refused):
        xa.Workbook(book)
    monkeypatch.setattr(safeio, "MAX_PDF_STREAM", 1000)
    pdf = tmp_path / "b.pdf"
    pdf.write_bytes(b"1 0 obj<</Type /Page /Contents 2 0 R>>endobj 2 0 obj<</Filter /FlateDecode>>stream\n"
                    + zlib.compress(b"0" * 50000) + b"\nendstream endobj")
    assert pdftext.pages(pdf) == [None]


def test_the_package(tmp_path):
    sys.path.insert(0, str(ROOT / "build"))
    import package
    z = package.build(tmp_path / "plant-watch.zip")
    names = zipfile.ZipFile(z).namelist()
    assert ".claude-plugin/plugin.json" in names and "LICENSE" in names and "NOTICE" in names
    assert "demo/northfork-precision/plant/erp/items.csv" in names
    assert not any(n.startswith(("tests/", "build/", ".github/")) or "__pycache__" in n for n in names)
    assert package.build(tmp_path / "again.zip").read_bytes() == z.read_bytes()


def test_user_facing_words():
    files = [p for d in ("skills", "data", "scripts", "demo") for p in (ROOT / d).rglob("*")
             if p.suffix in (".md", ".json", ".py", ".csv")] + [ROOT / "README.md", ROOT / "SECURITY.md"]
    for p in files:
        t = p.read_text(encoding="utf-8").lower()
        assert "training" not in t and "transformation" not in t, p
        if p.name not in ("pdfmini.py", "pdftext.py"):
            assert "\u2014" not in t, p


def test_command_line_refusal_is_plain_json(tmp_path):
    (tmp_path / "s.json").write_text(json.dumps({"folder": str(tmp_path), "sources": [{"id": "x", "kind": "rows", "glob": "../*"}]}))
    r = subprocess.run([sys.executable, str(ROOT / "scripts" / "watch.py"), "inventory", "--config", str(tmp_path / "s.json")],
                       capture_output=True, text=True)
    assert r.returncode == 1 and json.loads(r.stdout)["ok"] is False

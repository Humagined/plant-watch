"""Helpers to run plant-watch against the rolling Northfork kit."""
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "scripts"))
import pdftext  # noqa: E402

KIT = Path(__file__).resolve().parent / "kit"
KIT_PY = Path(sys.executable)
DAYS = ["2026-11-09", "2026-11-10", "2026-11-11", "2026-11-12", "2026-11-13", "2026-11-16", "2026-11-17",
        "2026-11-18", "2026-11-19", "2026-11-20", "2026-11-23", "2026-11-24", "2026-11-25"]


def build(base, through):
    plant = base / "plant"
    subprocess.run([str(KIT_PY), str(KIT / "rolling.py"), "--through", through, "--out", str(plant)],
                   check=True, capture_output=True)
    return plant, json.loads((base / f"answer_key_{through}.json").read_text())


def config(base, plant):
    text = (ROOT / "fixtures" / "northfork_sources.json").read_text()
    text = text.replace("{PLANT}", str(plant)).replace("{LINKED}", str(base / "linked"))
    p = base / "sources.json"
    p.write_text(text)
    return p


def model_reading(pdf_path):
    """Stands in for the model reading a weekly pack against its confirmed template."""
    lines = pdftext.pages(pdf_path)[0]
    m = re.search(r"(\d+) (\w+) to (\d+) (\w+) (\d{4})", " ".join(lines))
    start = f"{m.group(5)}-11-{int(m.group(1)):02d}"
    end = f"{m.group(5)}-11-{int(m.group(3)):02d}"
    figs = []
    for ln in lines:
        mm = re.match(r"([A-Z]+-[A-Z0-9]+-\d+|Plant total) ([\d,]+) ([\d,]+) ([\d,]+) ([\d.]+)%", ln)
        if mm:
            ent = "total" if mm.group(1) == "Plant total" else mm.group(1)
            for metric, idx in (("good", 2), ("scrap", 3), ("downtime", 4)):
                figs.append({"metric": metric, "entity": ent, "value": float(mm.group(idx).replace(",", "")),
                             "unit": "min" if metric == "downtime" else "units", "page": 1, "quote": ln})
    note = lines[lines.index("Notes") + 1:]
    end_d = end
    pub = {"2026-11-13": "2026-11-16", "2026-11-20": "2026-11-23"}[end_d]
    return {"file": Path(pdf_path).name, "series": "Northfork weekly OEE pack", "period": {"start": start, "end": end},
            "published": pub, "footnote": " ".join(note), "figures": figs}


def linked_reading(xlsx_path):
    """Stands in for the model reading the linked daily summary: one row per day, each figure cited by cell."""
    import sources as S
    import xlsx_audit as xa
    wb = xa.Workbook(xlsx_path)
    figs = []
    for r in range(2, 200):
        d = wb.value("Daily", (1, r))
        if not d or d["value"] in (None, ""):
            break
        day = S.to_date(d["value"]).isoformat()
        for metric, col, unit in (("good", 2, "units"), ("scrap", 3, "units"), ("downtime", 4, "min")):
            v = wb.value("Daily", (col, r))
            figs.append({"metric": metric, "entity": "total", "date": day, "value": float(v["value"]), "unit": unit,
                         "sheet": "Daily", "cell": f"{'ABCD'[col - 1]}{r}"})
    return {"file": Path(xlsx_path).name, "source": "daily_summary", "series": "Daily production summary",
            "footnote": None, "figures": figs}


def read_reports(plant, state):
    ex = Path(state) / "extractions"
    ex.mkdir(parents=True, exist_ok=True)
    for pdf in sorted((plant / "reports").glob("*.pdf")):
        (ex / (pdf.stem + ".json")).write_text(json.dumps(model_reading(pdf)))
    linked = plant.parent / "linked" / "daily_production_summary.xlsx"
    if linked.exists():
        (ex / "daily_production_summary.json").write_text(json.dumps(linked_reading(linked)))

"""The five problems the first Cowork run on the day-4 Northfork plant turned up (0.1.3)."""
import json
import shutil
import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import northfork as nf  # noqa: E402
import ledger  # noqa: E402
import pdfmini  # noqa: E402
import pdftext  # noqa: E402
import sources as S  # noqa: E402
import watch  # noqa: E402

DICT = Path(__file__).resolve().parent / "fixtures" / "northfork_dictionary.json"
DAY4 = "2026-11-12"


def run(base, exclude=None, edit_reading=None, cfg_path=None):
    plant, _ = nf.build(base, DAY4)
    cfg = S.load_config(cfg_path(base, plant) if cfg_path else nf.config(base, plant))
    state = base / "plant-watch" / "state"
    state.mkdir(parents=True)
    d = json.loads(DICT.read_text())
    if exclude is not None:
        d["metrics"]["downtime"]["exclude_categories"] = exclude
    (state / "dictionary.json").write_text(json.dumps(d))
    nf.read_reports(plant, state)
    if edit_reading:
        p = state / "extractions" / "daily_production_summary.json"
        p.write_text(json.dumps(edit_reading(json.loads(p.read_text()))))
    res = watch.run(cfg, state, date.fromisoformat(DAY4))
    return [e for e in ledger.load(state).values() if e["active"]], res, cfg, state


def test_a_source_that_agrees_with_the_record_is_still_judged_against_the_definition(tmp_path):
    found, *_ = run(tmp_path, exclude=["CHG", "PM"])
    pm = [e for e in found if e["category"] == "definition" and e.get("code") == "PM"]
    assert len(pm) == 1, "the shift log still counts the 60-minute PM stop on MILL-VX2-02"
    assert pm[0]["kind"] == "definition_mismatch" and pm[0]["gap"] == -60 and pm[0]["onset"] == DAY4
    assert pm[0]["evidence"]["examples"][0]["entity"] == "MILL-VX2-02"
    chg = next(e for e in found if e["kind"] == "definition_change")
    assert "matches your definition" not in chg["detail"] and "still counts PM" in chg["detail"]
    # Before 12 Nov the log counted CHG; that's the change finding's story, not a second finding.
    assert not [e for e in found if e["category"] == "definition" and e.get("code") == "CHG"
                and e["kind"] == "definition_mismatch"]


def test_no_question_asks_to_undo_what_the_definition_already_says(tmp_path):
    found, *_ = run(tmp_path, exclude=["CHG", "PM"])
    chg = next(e for e in found if e["kind"] == "definition_change")
    assert "go back to counting" not in chg["question"]
    # With a definition that counts changeovers, the same finding should still ask it.
    found, *_ = run(tmp_path / "counts", exclude=[])
    chg = next(e for e in found if e["kind"] == "definition_change")
    assert "go back to counting CHG" in chg["question"]


def test_the_linked_summary_is_compared_figure_by_figure(tmp_path):
    found, *_ = run(tmp_path)
    assert not [e for e in found if e["category"] == "report"], "the summary agrees with the MES on day 4"

    def bump(ex):
        next(f for f in ex["figures"] if f["metric"] == "good" and f["date"] == DAY4)["value"] -= 100
        return ex
    found, *_ = run(tmp_path / "wrong", edit_reading=bump)
    rep = [e for e in found if e["category"] == "report"]
    assert [e["kind"] for e in rep] == ["extraction_unverified"], "a reading that isn't in the cited cell isn't used"
    assert "cited cell" in rep[0]["title"]


def test_the_linked_summary_is_judged_against_the_definition(tmp_path):
    found, *_ = run(tmp_path, exclude=["CHG", "PM"])
    rep = next(e for e in found if e["kind"] == "report_vs_source")
    assert rep["code"] == "+CHG+PM" and "counts CHG and PM" in rep["title"]
    day4 = next(f for f in rep["evidence"]["figures"] if str(f["start"]) == DAY4)
    import csv
    with (tmp_path / "plant" / "exports" / f"mes_downtime_events_{DAY4}.csv").open(encoding="utf-8") as f:
        left_out = sum(float(r["duration_min"]) for r in csv.DictReader(f) if r["reason_code"] in ("CHG", "PM"))
    assert left_out > 60 and day4["report"] - day4["data"] == left_out


def test_a_reading_from_the_connector_alone_is_compared_without_a_file(tmp_path):
    def connector_only(ex):
        ex["file"] = "Daily production summary (SharePoint)"
        return ex
    found, *_ = run(tmp_path, exclude=["CHG", "PM"], edit_reading=connector_only)
    assert any(e["kind"] == "report_vs_source" for e in found)


def test_pdf_keeps_the_stream_separator():
    raw = pdfmini.esc("Production › Downtime – “planned”")
    assert "?" not in raw
    assert "Production › Downtime – “planned”" in pdftext._string(raw.encode("cp1252"))
    assert pdfmini.width("›", 10) < pdfmini.width("m", 10)


def test_a_copy_of_a_workbook_elsewhere_in_the_folder_is_audited_once(tmp_path):
    def parent_folder(base, plant):
        c = json.loads(nf.config(base, plant).read_text())
        c["folder"] = str(base)          # the person pointed Plant Watch at the whole folder
        for s in c["sources"]:
            if "glob" in s:
                s["glob"] = "plant/" + s["glob"]
        p = base / "sources_parent.json"
        p.write_text(json.dumps(c))
        return p
    _, res, cfg, state = run(tmp_path, cfg_path=parent_folder)
    names = [a["workbook"] for a in res["audits"]]
    assert names.count("supervisor_shift_log_Nov.xlsx") == 1
    log = next(a for a in res["audits"] if a["workbook"] == "supervisor_shift_log_Nov.xlsx")
    assert log["path"] == "plant/spreadsheets/supervisor_shift_log_Nov.xlsx"
    assert log["copies"] == ["excel_saved/supervisor_shift_log_Nov.xlsx"]
    import report
    html, pdf = tmp_path / "plant-watch" / "r.html", tmp_path / "plant-watch" / "r.pdf"
    report.render(cfg, state, str(html), str(pdf))
    text = html.read_text()
    assert text.count("<h3>supervisor_shift_log_Nov.xlsx</h3>") == 1
    assert "excel_saved/supervisor_shift_log_Nov.xlsx" in text and "checked twice" in text
    pdf_text = " ".join(" ".join(p or []) for p in pdftext.pages(pdf))
    assert "›" in pdf_text and "excel_saved/supervisor_shift_log_Nov.xlsx" in pdf_text

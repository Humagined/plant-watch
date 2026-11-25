import json
import re
import sys
import zipfile
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
import northfork as nf  # noqa: E402
import watch  # noqa: E402
import sources as S  # noqa: E402


def test_scan_groups_daily_files_and_export_round_trips(tmp_path):
    plant, _ = nf.build(tmp_path, "2026-11-25")
    groups = {g["pattern"]: g for g in watch.scan(plant)["groups"]}
    assert groups["exports/mes_downtime_events_*.csv"]["files"] == 13
    assert "Notes" in groups["spreadsheets/supervisor_shift_log_Nov.xlsx"]["sheets"]
    cfg = S.load_config(nf.config(tmp_path, plant))
    state = tmp_path / "state"
    state.mkdir()
    nf.read_reports(plant, state)
    watch.run(cfg, state, date(2026, 11, 25))
    csv_ = watch.export(state, tmp_path / "f.csv")
    js = watch.export(state, tmp_path / "f.json")
    assert csv_["count"] == js["count"] == 39
    data = json.loads((tmp_path / "f.json").read_text(encoding="utf-8"))
    assert data["schema"] == "plant-watch.findings/v1"
    assert (tmp_path / "f.csv").read_text(encoding="utf-8-sig").splitlines()[0].startswith("id,kind,category,stage")


def test_skill_descriptions_have_no_angle_brackets():
    for skill in (ROOT / "plugin" / "skills").glob("*/SKILL.md"):
        desc = re.search(r"^description:(.*)$", skill.read_text(encoding="utf-8"), re.M).group(1)
        assert "<" not in desc and ">" not in desc, skill

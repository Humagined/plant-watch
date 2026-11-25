"""The findings ledger: one JSON line per finding, kept in the state folder.

The tool sets first_seen, last_seen, active and cleared_on. Only a person sets status
(open -> acknowledged | dismissed | resolved), always with who, when and why.
"""

import hashlib
import json
from datetime import date
from pathlib import Path

import safeio

HUMAN_STATUSES = ("open", "acknowledged", "dismissed", "resolved")


def fid(f):
    return hashlib.sha1("|".join([f["kind"]] + [str(p) for p in f["id_parts"]]).encode()).hexdigest()[:12]


def _jsonable(v):
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, dict):
        return {k: _jsonable(x) for k, x in v.items()}
    if isinstance(v, (list, tuple, set)):
        return [_jsonable(x) for x in v]
    return v


def load(state):
    p = Path(state) / "findings.jsonl"
    if not p.exists():
        return {}
    return {e["id"]: e for e in (json.loads(l) for l in p.read_text(encoding="utf-8").splitlines() if l.strip())}


def save(state, entries):
    p = Path(state) / "findings.jsonl"
    safeio.save_state(p, "".join(json.dumps(e, default=str) + "\n" for e in
                                 sorted(entries.values(), key=lambda e: (e.get("first_seen") or "", e["id"]))))


def update(state, findings, run_date):
    """Merge this run's findings. Returns {"new": [...], "changed": [...], "cleared": [...], "active": [...]}."""
    entries = load(state)
    run = run_date.isoformat()
    seen, new, changed = set(), [], []
    for f in findings:
        i = fid(f)
        seen.add(i)
        data = _jsonable({k: v for k, v in f.items() if k != "id_parts"})
        if i in entries:
            e = entries[i]
            was_active = e.get("active")
            before = {k: e.get(k) for k in ("detail", "gap", "cells", "rows")}
            e.update(data)
            e["last_seen"], e["active"], e["cleared_on"] = run, True, None
            if not was_active:
                new.append(e)              # came back after clearing
            elif before != {k: e.get(k) for k in before}:
                changed.append(e)
        else:
            e = {"id": i, **data, "first_seen": run, "last_seen": run, "active": True, "cleared_on": None,
                 "status": "open", "status_by": None, "status_on": None, "status_reason": None}
            entries[i] = e
            new.append(e)
    cleared = []
    for i, e in entries.items():
        if i not in seen and e.get("active"):
            e["active"], e["cleared_on"] = False, run
            cleared.append(e)
    save(state, entries)
    return {"new": new, "changed": changed, "cleared": cleared,
            "active": [e for e in entries.values() if e.get("active")]}


def set_status(state, finding_id, status, by, reason, on=None):
    if status not in HUMAN_STATUSES:
        raise ValueError(f"status must be one of {', '.join(HUMAN_STATUSES)}")
    if not by or not reason:
        raise ValueError("who and why are both required")
    entries = load(state)
    matches = [i for i in entries if i.startswith(finding_id)]
    if len(matches) != 1:
        raise KeyError(f"{len(matches)} findings match '{finding_id}'")
    e = entries[matches[0]]
    e["status"], e["status_by"], e["status_reason"] = status, by, reason
    e["status_on"] = (on or date.today()).isoformat()
    save(state, entries)
    return e

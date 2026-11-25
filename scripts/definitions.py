"""The confirmed definitions, checked and written out as a file the person keeps. Standard library only.

Nothing here sends anything. The file is plain: what each figure means, in the plant's own words, who
owns it, and the questions that were asked and answered.
"""

import json
from datetime import date

FORMAT = "plant-watch.definitions/2"
STATUSES = ("confirmed", "not_settled", "retired")
FIELDS = ("label", "metric", "meaning", "definition", "unit", "grain", "measure", "dated_by", "counts", "excludes",
          "exclude_categories", "record", "owner", "confirmed_by", "confirmed_on", "variants_seen")


def entries(dictionary):
    for k, e in (dictionary.get("metrics") or {}).items():
        yield "metric", k, e
    for k, e in (dictionary.get("formulas") or {}).items():
        yield "formula", k, e


def check(dictionary):
    """Problems worth fixing before the file is used, in plain words. Empty when it's sound."""
    problems = []
    for kind, key, e in entries(dictionary):
        where = f"{kind} '{e.get('label') or key}'"
        status = e.get("status", "not_settled")
        if status not in STATUSES:
            problems.append(f"{where}: status must be confirmed, not_settled or retired, not '{status}'")
        if e.get("measure") and e["measure"] not in ("sum", "ratio", "period"):
            problems.append(f"{where}: measure must be sum, ratio or period, not '{e['measure']}'")
        if status == "confirmed" and not (e.get("owner") or {}).get("name"):
            problems.append(f"{where}: confirmed without an owner (who is accountable for this number?)")
    return problems


def document(dictionary, plant, today=None):
    out, not_settled = [], 0
    for kind, key, e in entries(dictionary):
        if e.get("status") == "confirmed":
            out.append({"kind": kind, "key": key, **{f: e[f] for f in FIELDS if e.get(f) not in (None, [], "")},
                        "questions": [{f: d[f] for f in ("question", "answer", "rationale") if d.get(f)}
                                      for d in e.get("decisions") or []]})
        elif e.get("status", "not_settled") == "not_settled":
            not_settled += 1
    return {"format": FORMAT, "plant": plant, "dictionary_version": dictionary.get("version", 1),
            "exported_on": (today or date.today()).isoformat(), "definitions": out, "not_settled": not_settled}


def text(doc):
    return json.dumps(doc, indent=1, ensure_ascii=False) + "\n"

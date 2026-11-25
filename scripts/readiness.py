"""Readiness for AI agents: a score from 0 to 100 worked out from the findings. Standard library only.

An agent answers with whatever it finds, so what matters is whether the numbers mean one thing, the records
connect, the spreadsheets hold up, each number has an owner, and the reports match the data. Those are the parts
a knowledge graph is built from; the report says it in plain words. Five checks, 20 points each. A check with
nothing to measure is left out and the rest are scaled, and the report says so. Fewer than three checks: no score.

This measures the files that were read. It isn't a promise about an agent, or about files that weren't read.
"""

import detect
import sources as S

MIN_CHECKS = 3
REPORT_PROBLEMS = {"stale_link", "link_unreachable", "extraction_unverified", "report_vs_source", "report_definition_drift"}

CHECKS = {
    "meaning": ("Do your numbers mean one thing?",
                "An agent can't tell which version of a number you meant.",
                "Settle the questions listed under the mismatches, starting with the first."),
    "records": ("Do your records connect?",
                "An agent follows a job from one system to another by its number. A number that doesn't match is a dead end.",
                "Agree on one format for job numbers, and correct the ones that don't match."),
    "spreadsheets": ("Do your spreadsheets hold up?",
                     "An agent reads a broken formula's result as if it were right.",
                     "Fix the spreadsheet problems listed first."),
    "owners": ("Does someone own each number?",
               "When an agent meets a question, someone has to answer it.",
               "Name an owner for each number, and confirm what it means."),
    "reports": ("Do your reports match your data?",
                "An agent may read an old pack and trust it.",
                "Refresh the late reports, and correct the ones that don't tie out."),
}
BANDS = [(85, "Ready for a first agent", "Start with one agent on one job, with a person checking its answers."),
         (65, "Close", "A few fixes stand between you and a first agent."),
         (40, "Getting there", "Settle the basics first. An agent would repeat today's disagreements."),
         (0, "Not ready yet", "An agent would give confident answers from numbers that disagree.")]


def check(key, good, total, unit):
    return {"key": key, "good": good, "total": total, "unit": unit, "share": (good / total) if total else None}


def assess(cfg, facts, entries, dictionary, formula_cells):
    active = [e for e in entries if e.get("active")]
    metrics = dictionary.get("metrics") or {}
    formulas = dictionary.get("formulas") or {}

    # Meaning: each metric counted in two or more places, and each label calculated in two or more workbooks.
    shared = [m for m in detect.METRICS if m != "job_key" and len(
        {f["source"] for f in facts if f.get(m) is not None}
        | {s["id"] for s in cfg["sources"] if s.get("kind") in ("report", "link") and m in s.get("measures", [])}) > 1]
    disagree = {e.get("metric") for e in active if e["kind"] in (
        "definition_change", "definition_mismatch", "unexplained_gap", "report_vs_source", "report_definition_drift")}
    labels = [g for g in detect.formula_groups(S.workbooks(cfg)) if g["workbooks"] > 1]
    conflicts = {e.get("label") for e in active if e["kind"] == "formula_definition_conflict"}
    one = sum(m not in disagree for m in shared) + sum(g["label"] not in conflicts for g in labels)
    out = [check("meaning", one, len(shared) + len(labels), "shared numbers agree where we compared them")]

    # Records: rows whose job number is in the master list exactly as typed.
    good = total = 0
    for src in cfg["sources"]:
        if src.get("key") and src.get("key_master"):
            master = S.master_keys(cfg, src["key_master"])
            rows = [f for f in facts if f["source"] == src["id"] and f.get("key")]
            total += len(rows)
            good += sum(f["key"] in master for f in rows)
    out.append(check("records", good, total, "records match a work order as typed"))

    # Spreadsheets: formula cells with no problem flagged on them.
    flagged = {(e["workbook"], e.get("sheet"), e.get("cell")) for e in active if e["category"] == "excel" and e.get("formula")}
    out.append(check("spreadsheets", max(0, formula_cells - len(flagged)), formula_cells, "formulas had no problem flagged"))

    # Owners: the shared numbers a person has confirmed and put a name to.
    named = [metrics.get(m) or {} for m in shared] + [formulas.get(g["label"].lower()) or {} for g in labels]
    out.append(check("owners", sum(e.get("status") == "confirmed" and bool((e.get("owner") or {}).get("name")) for e in named),
                     len(named), "shared numbers have a confirmed meaning and an owner"))

    # Reports: each report file or link without a problem of its own.
    reports = {}
    for s in cfg["sources"]:
        if s.get("kind") == "report":
            reports.update({p.name: True for p in S.files_for(cfg, s)})
        elif s.get("kind") == "link":
            reports[s.get("label") or s["id"]] = True
    ids = {s["id"]: s.get("label") or s["id"] for s in cfg["sources"] if s.get("kind") == "link"}
    for e in active:
        if e["kind"] in REPORT_PROBLEMS:
            for name in [e.get("report")] + [ids[x] for x in e.get("sources") or [] if x in ids]:
                if name in reports:
                    reports[name] = False
    out.append(check("reports", sum(reports.values()), len(reports), "reports had no problem flagged"))

    measured = [c for c in out if c["share"] is not None]
    if len(measured) < MIN_CHECKS:
        return {"score": None, "band": None, "advice": None, "checks": out, "measured": len(measured), "next": []}
    score = int(100 * sum(c["share"] for c in measured) / len(measured))        # rounded down
    weakest = min(c["share"] for c in measured)
    if weakest < 0.5:                       # one weak check holds the result back: an agent needs them together
        score = min(score, 64)
    band = next(b for b in BANDS if score >= b[0])
    nxt = [c["key"] for c in sorted(measured, key=lambda c: c["share"]) if c["share"] < 0.95][:2]
    return {"score": score, "band": band[1], "advice": band[2], "checks": out, "measured": len(measured),
            "held_back": weakest < 0.5, "next": nxt}

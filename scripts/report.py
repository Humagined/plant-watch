"""The Metric Mismatch Report: HTML for reading, PDF for printing. Every number comes from the
ledger, the dictionary and the snapshots; nothing here is typed by the model."""

import html
import re
import json
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import detect
import ledger
import pdfmini
import sources as S
import safeio
import xlsx_audit as xa

DATA = Path(__file__).resolve().parent.parent / "data"
STREAMS = json.loads((DATA / "value_streams.json").read_text(encoding="utf-8"))["streams"]
USE_CASES = json.loads((DATA / "use_cases.json").read_text(encoding="utf-8"))["stages"]
STAGE = {st["key"]: (s, st) for s in STREAMS for st in s["stages"]}
DEFN_KINDS = {"definition_change", "definition_mismatch", "unexplained_gap", "report_vs_source",
              "report_definition_drift", "code_flag_changed", "formula_definition_conflict"}
EXCEL_GROUPS = [("Broken", {"cached_error", "ref_error", "broken_defined_name", "circular_reference", "workbook_unreadable"}),
                ("Will break", {"div_by_blank", "lookup_out_of_range", "missing_sheet_reference", "unknown_function"}),
                ("Improper references", {"external_link", "wrong_period_sheet", "hidden_row_in_range",
                                         "hidden_sheet_reference", "empty_range"}),
                ("Range errors", {"range_short"}),
                ("Inconsistent formulas", {"pattern_break", "overwritten_formula"}),
                ("Hidden assumptions", {"hardcoded_constant", "text_in_numeric_column", "duplicate_row"})]


EXCEL_ROWS_PER_WORKBOOK = 150      # problems listed per workbook in the report; the export has every cell


def excel_groups(entries):
    """One problem per run of cells: the same kind, formula pattern and shown value in adjacent rows of
    one column (a broken formula filled down 900 rows is one fix), and adjacent hidden rows counted by
    the same formulas. The ledger keeps every cell; this is how the report shows them."""
    keyed, single = defaultdict(list), []
    for x in entries:
        m = re.fullmatch(r"([A-Z]{1,3})(\d+)", x.get("cell") or "")
        if not m or not x.get("sheet"):
            single.append([x])
            continue
        col, row = m.group(1), int(m.group(2))
        if x["kind"] == "hidden_row_in_range":
            key = (x["workbook"], x["sheet"], x["kind"], "rows", x.get("used_by_count"), tuple((x.get("used_by") or [])[:3]))
        else:
            shape = xa.r1c1(x["formula"], (xa.col_num(col), row)) if x.get("formula") else None
            key = (x["workbook"], x["sheet"], x["kind"], col, shape, str(x.get("cached")), str(x.get("will_show")))
        keyed[key].append((row, x))
    runs = single
    for key, rows in keyed.items():
        rows.sort(key=lambda t: t[0])
        cur = [rows[0]]
        for r in rows[1:]:
            if r[0] == cur[-1][0] + 1:
                cur.append(r)
            else:
                runs.append([x for _, x in cur])
                cur = [r]
        runs.append([x for _, x in cur])
    out = []
    for xs in runs:
        first, last = xs[0], xs[-1]
        n = len(xs)
        if n == 1:
            where = first["cell"]
        elif first["kind"] == "hidden_row_in_range":
            where = f"rows {first['cell'][1:]}–{last['cell'][1:]}"
        else:
            where = f"{first['cell']}:{last['cell']}"
        if n == 1:
            problem = first["problem"]
        elif first["kind"] == "hidden_row_in_range":
            users = first.get("used_by") or []
            problem = (f"Rows {first['cell'][1:]}–{last['cell'][1:]} ({n:,} rows) are hidden but still counted by "
                       f"{first.get('used_by_count') or len(users):,} formula(s)" + (f", e.g. {', '.join(users[:3])}" if users else "") + ".")
        else:
            problem = f"{n:,} adjacent cells with the same formula and the same problem. The first: {first['problem']}"
        out.append({**first, "where": where, "count": n, "problem": problem,
                    "label": f"{first['workbook']} › {first.get('sheet') or 'workbook'} › {where}"})
    return sorted(out, key=lambda g: (g.get("severity", 9), -g["count"], g["workbook"], g.get("sheet") or "", g["cell"] or ""))


BOOKING_URL = "https://www.humagined.ai/book/bryce"
BOOKING = ("Found mismatches? Book a free 30-minute working session to walk through them: "
           f"{BOOKING_URL} (through January 31, 2027).")
LINK = f"<a href='{BOOKING_URL}'>{BOOKING_URL}</a>"


def e(x):
    return html.escape("" if x is None else str(x))


def kpis_for(f):
    m = detect.METRICS.get(f.get("metric") or "", {})
    stage_kpis = STAGE.get(f.get("stage"), (None, {}))[1].get("kpis", []) if f.get("stage") else []
    return sorted(set(m.get("affects", [])) | set(stage_kpis))


def model(cfg, state):
    entries = [x for x in ledger.load(state).values() if x.get("active")]
    runs = [json.loads(l) for l in (Path(state) / "runs.jsonl").read_text(encoding="utf-8").splitlines()] \
        if (Path(state) / "runs.jsonl").exists() else []
    as_of = runs[-1]["as_of"] if runs else date.today().isoformat()
    prev = runs[-2]["as_of"] if len(runs) > 1 else None
    all_entries = list(ledger.load(state).values())
    since = {"new": [x for x in entries if x["first_seen"] == as_of] if prev else [],
             "cleared": [x for x in all_entries if x.get("cleared_on") == as_of]}
    dictionary = (json.loads((Path(state) / "dictionary.json").read_text(encoding="utf-8"))
                  if (Path(state) / "dictionary.json").exists() else {}).get("metrics", {})
    snap = sorted((Path(state) / "snapshots").glob("*.json"))
    audits = json.loads(snap[-1].read_text(encoding="utf-8")).get("audits", []) if snap else []

    # Value stream summary: streams with a source or a finding.
    source_stages = set()
    for s in cfg["sources"]:
        if s.get("stage"):
            source_stages.add(s["stage"])
        for m in (s.get("measures") or {}):
            if detect.stage_of(m):
                source_stages.add(detect.stage_of(m))
    rows = {}
    for st in STREAMS:
        keys = {x["key"] for x in st["stages"]}
        fs = [x for x in entries if x.get("stage") in keys]
        if not fs and not (keys & source_stages):
            continue
        defs = [x for x in fs if x["kind"] in DEFN_KINDS]
        excel = [x for x in fs if x["category"] == "excel"]
        rows[st["key"]] = {
            "stream": st, "findings": fs,
            "metrics": sorted({x["metric"] for x in fs if x.get("metric")} |
                              {m for m, info in detect.METRICS.items() if info["stage"] in keys
                               and any(m in (s.get("measures") or {}) for s in cfg["sources"])}),
            "contested": sorted({x["metric"] for x in defs if x.get("metric")}),
            "mismatches": len(defs), "excel": len(excel),
            "keys": len([x for x in fs if x["category"] == "keys"]),
            "anomalies": len([x for x in fs if x["category"] == "anomaly"]),
            "reports": len([x for x in fs if x["category"] == "report" and x["kind"] not in DEFN_KINDS]),
            "kpis": sorted({k for x in fs for k in kpis_for(x)}),
            "share": round(sum(x.get("share") or 0 for x in defs), 4)}
    # Ranking rule (shown in the report): contested definitions on anchored KPIs, then gap size,
    # then Excel faults feeding those KPIs, then everything else.
    ranked = sorted(rows.values(), key=lambda r: (-len([m for m in r["contested"] if detect.METRICS.get(m, {}).get("affects")]),
                                                    -r["share"], -r["excel"], -(r["keys"] + r["anomalies"] + r["reports"])))
    # Suggested first use cases: rank use cases (stages grouped by the outcome they deliver), same rule.
    groups = defaultdict(list)
    for x in entries:
        if x.get("stage") in USE_CASES:
            groups[USE_CASES[x["stage"]]["title"]].append(x)
    scored = []
    for title, fs in groups.items():
        defs = [x for x in fs if x["kind"] in DEFN_KINDS]
        contested = {x["metric"] for x in defs if x.get("metric") and detect.METRICS.get(x["metric"], {}).get("affects")}
        excel_n = len([x for x in fs if x["category"] == "excel"])
        other = len([x for x in fs if x["category"] in ("keys", "anomaly", "report") and x["kind"] not in DEFN_KINDS])
        scored.append(((-len(contested), -sum(x.get("share") or 0 for x in defs), -excel_n, -other), title, fs, defs, excel_n))
    use_cases = []
    order = {k: i for i, k in enumerate(["definition_change", "definition_mismatch", "unexplained_gap",
                                         "report_vs_source", "code_flag_changed", "report_definition_drift"])}
    for _, title, fs, defs, excel_n in sorted(scored, key=lambda t: t[0])[:3]:
        defs = sorted(defs, key=lambda x: order.get(x["kind"], 9))
        stage = Counter(x["stage"] for x in fs).most_common(1)[0][0]
        others = [x for x in fs if x["category"] in ("keys", "anomaly", "report") and x["kind"] not in DEFN_KINDS]
        use_cases.append({"stream": STAGE[stage][0]["title"], "title": title, "outcome": USE_CASES[stage]["outcome"],
                          "why": [x["title"] for x in defs[:3]] + [x["title"] for x in others[:2]]
                                 + ([f"{excel_n} spreadsheet faults feeding these numbers"] if excel_n else []),
                          "kpis": sorted({k for x in fs for k in kpis_for(x)}),
                          "questions": sorted({x["question"] for x in defs + others if x.get("question")}),
                          "sources": sorted({detect.src_label(cfg, s) for x in fs for s in x.get("sources", [])
                                             if any(c["id"] == s for c in cfg["sources"])})})
    excel = [x for x in entries if x["category"] == "excel"]
    groups = excel_groups(excel)
    questions = []
    for x in sorted([x for x in entries if x.get("question")], key=lambda x: -(abs(x.get("share") or 0) * 100 + abs(x.get("gap") or 0) / 1000)):
        if x["question"] not in [q["question"] for q in questions]:
            questions.append({"question": x["question"], "from": x["title"], "stage": x.get("stage")})
    formula_defs = sorted([x for x in entries if x["kind"] == "formula_definition_conflict"],
                          key=lambda x: (-(x.get("workbooks") or 0), -(x.get("places") or 0), x["title"]))
    return {"plant": cfg.get("plant", ""), "as_of": as_of, "previous": prev, "since": since, "entries": entries,
            "formula_defs": formula_defs, "excel_groups": groups,
            "rows": ranked, "use_cases": use_cases, "excel": excel, "audits": audits, "dictionary": dictionary,
            "questions": questions, "sources": cfg["sources"]}


def stage_name(key):
    if not key or key not in STAGE:
        return "Not tied to a value stream"
    s, st = STAGE[key]
    return f"{s['title']} › {st['name']}"


CSS = """
:root{--bg:#f6f7f9;--fg:#16212e;--muted:#5b6b7c;--card:#ffffff;--line:#d9dfe6;--accent:#1f4f7a;--accent-soft:#e7eef6;
--bad:#a3342a;--warn:#9a6a10;--ok:#2f6b48;color-scheme:light}
@media (prefers-color-scheme:dark){:root{--bg:#0f1620;--fg:#e6ecf2;--muted:#9fb0c1;--card:#16202c;--line:#2a3746;
--accent:#8cb8e0;--accent-soft:#1c2b3b;--bad:#e58b80;--warn:#e0b25c;--ok:#7fc39a;color-scheme:dark}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 -apple-system,"Segoe UI",
Roboto,Helvetica,Arial,sans-serif}main{max-width:1080px;margin:0 auto;padding:32px max(16px,env(safe-area-inset-left))}
h1{font-size:28px;margin:0 0 4px;text-wrap:balance}h2{font-size:20px;margin:36px 0 10px;padding-top:12px;border-top:1px solid var(--line)}
h3{font-size:16px;margin:22px 0 6px}.muted{color:var(--muted)}.eyebrow{font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:var(--accent);font-weight:600}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin:10px 0}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.stat b{display:block;font-size:24px;font-variant-numeric:tabular-nums}.stat span{color:var(--muted);font-size:13px}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:10px;background:var(--card);margin:8px 0}
table{border-collapse:collapse;width:100%;font-size:13.5px}th,td{text-align:left;padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:top}
th{background:var(--accent-soft);font-weight:600;white-space:nowrap}td.n{text-align:right;font-variant-numeric:tabular-nums}
tr:last-child td{border-bottom:0}code{font:12.5px ui-monospace,SFMono-Regular,Menlo,monospace;word-break:break-all}
.pill{display:inline-block;font-size:12px;padding:1px 8px;border-radius:99px;border:1px solid var(--line);white-space:nowrap}
.sev1{color:var(--bad);border-color:var(--bad)}.sev2{color:var(--bad)}.sev3,.sev4{color:var(--warn);border-color:var(--warn)}
.sev5,.sev6{color:var(--muted)}.new{color:var(--accent);border-color:var(--accent)}.uc{border-left:4px solid var(--accent)}
ol,ul{padding-left:20px}footer{margin-top:40px;color:var(--muted);font-size:13px}
@media print{body{background:#fff;color:#000}.card,.tw,.stat{break-inside:avoid}}
"""


def render_html(m):
    H = []
    a = H.append
    a(f"<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
      f"<title>Metric Mismatch Report: {e(m['plant'])}</title><style>{CSS}</style></head><body><main>")
    a(f"<div class='eyebrow'>Metric Mismatch Report</div><h1>{e(m['plant'])}</h1>")
    a(f"<p class='muted'>Data through {e(m['as_of'])}. {len(m['sources'])} sources and {len(m['audits'])} workbooks checked."
      f" Findings are computed from the files listed here, not estimated.</p>")
    if m["previous"]:
        a(f"<div class='card'><b>Since the last run ({e(m['previous'])})</b>: {len(m['since']['new'])} new, "
          f"{len(m['since']['cleared'])} cleared.<ul>")
        for x in m["since"]["new"][:12]:
            a(f"<li><span class='pill new'>new</span> {e(x['title'])}</li>")
        for x in m["since"]["cleared"][:12]:
            a(f"<li><span class='pill'>cleared</span> {e(x['title'])}</li>")
        a("</ul></div>")

    # 0. Definitions that live in formulas
    if m["formula_defs"]:
        span = sum(1 for x in m["formula_defs"] if (x.get("workbooks") or 0) > 1)
        a("<h2>Same name, different calculation</h2>"
          f"<p class='muted'>{len(m['formula_defs'])} figure(s) are calculated more than one way in your spreadsheets"
          + (f", {span} of them across workbooks" if span else "") + ". Each formula is shown by the columns it reads, so a "
          "different layout doesn't count as a different definition. These are questions about what a number means, "
          "not broken cells.</p>")
        for x in m["formula_defs"]:
            ev = (x.get("evidence") or {}).get("variants", [])
            a(f"<div class='card'><b>{e(x['title'])}</b>"
              + (f" <span class='pill'>{e(detect.METRICS[x['metric']]['label'])}</span>" if x.get("metric") in detect.METRICS else "")
              + "<div class='tw'><table><tr><th>Calculated as</th><th class='n'>Places</th><th class='n'>Workbooks</th><th>For example</th></tr>")
            for v in ev:
                yours = " <span class='pill'>your definition</span>" if v.get("is_yours") else ""
                a(f"<tr><td><code>{e(v['definition'])}</code>{yours}</td>"
                  f"<td class='n'>{v['places']}</td><td class='n'>{v['workbooks']}</td><td><code>{e(v['example'])}</code></td></tr>")
            a("</table></div>" + (f"<p><b>To settle:</b> {e(x['question'])}</p>" if x.get("question") else "") + "</div>")

    # 1. Excel errors at a glance
    G = m["excel_groups"]
    a("<h2>Excel errors at a glance</h2>")
    a(f"<p class='muted'>{len(G):,} problem(s) across {sum(g['count'] for g in G):,} cell(s). A formula filled down a column "
      "with the same fault is counted once.</p>")
    a("<div class='grid'>")
    for name, kinds in EXCEL_GROUPS:
        gs = [g for g in G if g["kind"] in kinds]
        cells = sum(g["count"] for g in gs)
        a(f"<div class='stat'><b>{len(gs):,}</b><span>{e(name)}" + (f" ({cells:,} cells)" if cells != len(gs) else "") + "</span></div>")
    a("</div>")
    worst, per_kind = [], Counter()
    for g in G:
        if per_kind[g["kind"]] < 2 and len(worst) < 5:   # a spread of problems, not five of the same
            worst.append(g)
            per_kind[g["kind"]] += 1
    if worst:
        a("<h3>Worst five</h3><div class='tw'><table><tr><th>Where</th><th>Problem</th><th>Fix</th></tr>")
        for g in worst:
            a(f"<tr><td><code>{e(g['label'])}</code></td><td>{e(g['problem'])}</td><td>{e(g['fix'])}</td></tr>")
        a("</table></div>")

    # 2. Value stream summary
    a("<h2>Value stream summary</h2><div class='tw'><table><tr><th>Value stream</th><th>Metrics checked</th>"
      "<th>Contested definitions</th><th class='n'>Mismatches</th><th class='n'>Spreadsheet faults</th>"
      "<th class='n'>Key and record problems</th><th class='n'>Anomalies</th><th>KPIs affected</th></tr>")
    for r in m["rows"]:
        a(f"<tr><td><b>{e(r['stream']['title'])}</b></td><td>{e(', '.join(detect.METRICS.get(x, {}).get('label', x) for x in r['metrics'])) or '-'}</td>"
          f"<td>{e(', '.join(detect.METRICS.get(x, {}).get('label', x) for x in r['contested'])) or '-'}</td>"
          f"<td class='n'>{r['mismatches']}</td><td class='n'>{r['excel']}</td><td class='n'>{r['keys']}</td>"
          f"<td class='n'>{r['anomalies']}</td><td>{e(', '.join(r['kpis'])) or '-'}</td></tr>")
    a("</table></div>")

    # 3. Suggested first use cases
    a("<h2>Suggested first use cases</h2><p class='muted'>Ranked by contested definitions on headline KPIs, then the size of "
      "the gaps, then spreadsheet faults feeding those numbers, then other problems.</p>")
    for i, uc in enumerate(m["use_cases"], 1):
        a(f"<div class='card uc'><div class='eyebrow'>{i}. {e(uc['stream'])}</div><h3>{e(uc['title'])}</h3>"
          f"<p>{e(uc['outcome'])}</p><p><b>Why here:</b></p><ul>" + "".join(f"<li>{e(w)}</li>" for w in uc["why"]) + "</ul>"
          + (f"<p><b>Makes trustworthy:</b> {e(', '.join(uc['kpis']))}</p>" if uc["kpis"] else "")
          + (f"<p><b>Definitions to settle:</b></p><ul>" + "".join(f"<li>{e(q)}</li>" for q in uc["questions"]) + "</ul>"
             if uc["questions"] else "")
          + (f"<p class='muted'>Sources involved: {e(', '.join(uc['sources']))}</p>" if uc["sources"] else "") + "</div>")

    # 4. By stream › stage › metric
    a("<h2>By value stream, stage and metric</h2>")
    for r in m["rows"]:
        fs = [x for x in r["findings"] if x["category"] != "excel" and x["kind"] != "formula_definition_conflict"]
        if not fs:
            continue
        a(f"<h3>{e(r['stream']['title'])}</h3>")
        for stage in sorted({x.get('stage') for x in fs}, key=lambda s: s or ""):
            a(f"<p class='eyebrow'>{e(stage_name(stage))}</p>")
            for x in [x for x in fs if x.get("stage") == stage]:
                d = m["dictionary"].get(x.get("metric") or "", {})
                meaning = d.get("meaning") if d.get("status") == "confirmed" else "Not settled"
                a(f"<div class='card'><b>{e(x['title'])}</b> <span class='pill'>{e(x['kind'].replace('_', ' '))}</span>"
                  + (f" <span class='pill'>since {e(x.get('onset'))}</span>" if x.get("onset") else "")
                  + f"<p>{e(x['detail'])}</p>"
                  + (f"<p class='muted'>Your definition: {e(meaning)}</p>" if x.get("metric") in detect.METRICS and x["category"] == "definition" else ""))
                ev = x.get("evidence") or {}
                if ev.get("notes"):
                    a("<p class='muted'>Notes: " + "; ".join(f"“{e(n['line'])}”" for n in ev["notes"]) + "</p>")
                if ev.get("figures"):
                    a("<div class='tw'><table><tr><th>Line</th><th class='n'>Report</th><th class='n'>Data</th><th class='n'>Gap</th><th>Explained by</th></tr>")
                    for g in ev["figures"]:
                        a(f"<tr><td>{e(g.get('entity'))}</td><td class='n'>{g['report']:,.0f}</td><td class='n'>{g['data']:,.0f}</td>"
                          f"<td class='n'>{g['gap']:,.0f}</td><td>{e(g.get('explained_by') or '-')}</td></tr>")
                    a("</table></div>")
                if ev.get("examples"):
                    keys = [k for k in ev["examples"][0].keys()]
                    a("<div class='tw'><table><tr>" + "".join(f"<th>{e(k)}</th>" for k in keys) + "</tr>")
                    for g in ev["examples"]:
                        a("<tr>" + "".join(f"<td>{e(g.get(k))}</td>" for k in keys) + "</tr>")
                    a("</table></div>")
                if x.get("question"):
                    a(f"<p><b>To settle:</b> {e(x['question'])}</p>")
                a("</div>")

    # 5. Excel audit, complete
    a("<h2>Excel audit, by workbook</h2>")
    for au in m["audits"]:
        gs = [g for g in m["excel_groups"] if g["workbook"] == au["workbook"]]
        cells = sum(g["count"] for g in gs)
        a(f"<h3>{e(au['workbook'])}</h3><p class='muted'>{len(gs):,} problem(s) across {cells:,} cell(s). "
          f"Checks run: {e(', '.join(au['checks_run']))}.</p>")
        if au.get("copies"):
            a(f"<p class='muted'>Checked at {e(au.get('path') or au['workbook'])}. The same workbook is also at "
              f"{e('; '.join(au['copies']))} (same sheets and formulas), so it wasn't checked twice.</p>")
        if au.get("not_checked"):
            nc = au["not_checked"]
            a("<p class='muted'>Not checked: " + "; ".join(f"{e(n['sheet'])}!{e(n['cell'])} ({e(n['reason'])})" for n in nc[:20])
              + (f"; and {len(nc) - 20:,} more" if len(nc) > 20 else "") + "</p>")
        if gs:
            a("<div class='tw'><table><tr><th>Sheet</th><th>Cells</th><th>Formula</th><th>Shows</th><th>Problem</th><th>Fix</th><th>Feeds</th></tr>")
            for g in gs[:EXCEL_ROWS_PER_WORKBOOK]:
                shows = g.get("cached") or (f"will show {g['will_show']}" if g.get("will_show") else "")
                a(f"<tr><td>{e(g.get('sheet') or 'workbook')}</td><td><code>{e(g['where'])}</code>"
                  + (f"<br><span class='muted'>{g['count']:,} cells</span>" if g["count"] > 1 else "") + "</td>"
                  f"<td><code>{e(g.get('formula') or '')}</code></td><td>{e(shows)}</td>"
                  f"<td><span class='pill sev{g.get('severity', 6)}'>{e(g.get('severity_label', ''))}</span> {e(g['problem'])}</td>"
                  f"<td>{e(g['fix'])}</td><td>{e(stage_name(g.get('stage')))}</td></tr>")
            a("</table></div>")
            if len(gs) > EXCEL_ROWS_PER_WORKBOOK:
                rest = gs[EXCEL_ROWS_PER_WORKBOOK:]
                a(f"<p class='muted'>And {len(rest):,} more problem(s) across {sum(g['count'] for g in rest):,} cell(s), less "
                  "severe than those above. Export the findings (CSV or JSON) for the rest.</p>")

    # 6. Questions to settle
    if m["questions"]:
        a("<h2>Questions to settle</h2><ol>")
        for q in m["questions"]:
            a(f"<li>{e(q['question'])} <span class='muted'>({e(stage_name(q['stage']))})</span></li>")
        a("</ol>")
    a("<footer>Generated by Plant Watch from the plant's own files. Plant Watch doesn't send your files or results."
      f"<p>{e(BOOKING).replace(BOOKING_URL, LINK)}</p></footer></main></body></html>")
    return "".join(H)


def render_pdf(m):
    pages, p, y = [], None, 0

    def newpage():
        nonlocal p, y
        p = pdfmini.Page()
        pages.append(p)
        p.text(54, 756, f"Metric Mismatch Report: {m['plant']} (data through {m['as_of']})", 8, gray=0.45)
        y = 730

    def line(text, size=9.5, bold=False, gray=0.0, indent=0):
        nonlocal y
        words, cur = str(text).split(), ""
        for w in words:
            if pdfmini.width(cur + " " + w, size, bold) > 500 - indent:
                if y < 60:
                    newpage()
                p.text(54 + indent, y, cur.strip(), size, bold=bold, gray=gray)
                y -= size + 4
                cur = ""
            cur += " " + w
        if cur.strip():
            if y < 60:
                newpage()
            p.text(54 + indent, y, cur.strip(), size, bold=bold, gray=gray)
            y -= size + 4

    def gap(n=6):
        nonlocal y
        y -= n

    newpage()
    line("Metric Mismatch Report", 18, bold=True)
    line(m["plant"], 12)
    line(f"Data through {m['as_of']}. {len(m['sources'])} sources and {len(m['audits'])} workbooks checked.", 9, gray=0.35)
    gap(10)
    if m["formula_defs"]:
        line("Same name, different calculation", 13, bold=True)
        for x in m["formula_defs"]:
            line(x["title"], 10, bold=True)
            for v in (x.get("evidence") or {}).get("variants", []):
                line(f"= {v['definition']}   ({v['places']} places, {v['workbooks']} workbooks; e.g. {v['example']})"
                     + ("  [your definition]" if v.get("is_yours") else ""), 8.5, indent=10)
            if x.get("question"):
                line(f"To settle: {x['question']}", 8.5, gray=0.3, indent=10)
            gap(3)
        gap(8)
    line("Excel errors at a glance", 13, bold=True)
    G = m["excel_groups"]
    line(f"{len(G):,} problem(s) across {sum(g['count'] for g in G):,} cell(s).", 9.5)
    line("   ".join(f"{name}: {sum(1 for g in G if g['kind'] in kinds)}" for name, kinds in EXCEL_GROUPS), 9.5)
    gap(10)
    line("Value stream summary", 13, bold=True)
    for r in m["rows"]:
        line(f"{r['stream']['title']}: {r['mismatches']} mismatches, {r['excel']} spreadsheet faults, {r['keys']} key/record "
             f"problems, {r['anomalies']} anomalies. Contested: {', '.join(r['contested']) or 'none'}.", 9.5)
    gap(10)
    line("Suggested first use cases", 13, bold=True)
    for i, uc in enumerate(m["use_cases"], 1):
        line(f"{i}. {uc['title']} ({uc['stream']})", 11, bold=True)
        line(uc["outcome"], 9.5)
        for w in uc["why"]:
            line(f"- {w}", 9.5, indent=10)
        gap(4)
    gap(8)
    line("Findings by value stream and stage", 13, bold=True)
    for r in m["rows"]:
        fs = [x for x in r["findings"] if x["category"] != "excel" and x["kind"] != "formula_definition_conflict"]
        if fs:
            line(r["stream"]["title"], 11, bold=True)
            for x in fs:
                line(f"{stage_name(x.get('stage'))}: {x['title']}", 9.5, bold=True, indent=6)
                line(x["detail"], 9, indent=6)
                if x.get("question"):
                    line(f"To settle: {x['question']}", 9, gray=0.3, indent=6)
                gap(3)
    gap(8)
    line("Excel audit, by workbook", 13, bold=True)
    for au in m["audits"]:
        gs = [g for g in G if g["workbook"] == au["workbook"]]
        line(f"{au['workbook']} ({len(gs):,} problems, {sum(g['count'] for g in gs):,} cells)", 11, bold=True)
        if au.get("copies"):
            line(f"Checked at {au.get('path') or au['workbook']}. The same workbook is also at {'; '.join(au['copies'])} "
                 f"(same sheets and formulas), so it wasn't checked twice.", 8.5, gray=0.35, indent=6)
        for g in gs[:EXCEL_ROWS_PER_WORKBOOK]:
            line(f"{g.get('sheet') or 'workbook'}!{g['where']}  {g.get('formula') or ''}", 8.5, bold=True, indent=6)
            line(f"{g['problem']} Fix: {g['fix']}", 8.5, indent=6)
        if len(gs) > EXCEL_ROWS_PER_WORKBOOK:
            line(f"And {len(gs) - EXCEL_ROWS_PER_WORKBOOK:,} more problem(s); export the findings for the rest.",
                 8.5, gray=0.35, indent=6)
    if m["questions"]:
        gap(8)
        line("Questions to settle", 13, bold=True)
        for i, q in enumerate(m["questions"], 1):
            line(f"{i}. {q['question']}", 9.5)
    gap(10)
    line(BOOKING, 9.5)
    total = len(pages)
    for i, pg in enumerate(pages, 1):
        pg.text(558, 30, f"Page {i} of {total}", 8, gray=0.45, right=True)
    return pdfmini.pdf(pages)


def render(cfg, state, out_path, pdf_path=None):
    m = model(cfg, state)
    sources = [p for s in cfg["sources"] for p in S.files_for(cfg, s)]
    safeio.remember(state, cfg)
    safeio.save_result(state, out_path, render_html(m), sources)
    if pdf_path:
        safeio.save_result(state, pdf_path, render_pdf(m), sources)
    return {"ok": True, "html": str(out_path), "pdf": pdf_path, "as_of": m["as_of"],
            "streams": [{"stream": r["stream"]["title"], "mismatches": r["mismatches"], "excel": r["excel"],
                         "keys": r["keys"], "anomalies": r["anomalies"], "contested": r["contested"]} for r in m["rows"]],
            "use_cases": [{"title": u["title"], "stream": u["stream"]} for u in m["use_cases"]],
            "excel_total": len(m["excel"]), "excel_problems": len(m["excel_groups"]),
            "questions": [q["question"] for q in m["questions"]]}

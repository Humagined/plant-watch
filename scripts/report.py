"""The Metric Mismatch Report: HTML for reading, PDF for printing. Every number comes from the
ledger, the dictionary and the snapshots; nothing here is typed by the model."""

import base64
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


BOOKING_URL = "https://www.humagined.ai/book"
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
:root{--mist:#F6F6F6;--navy:#0E1A2B;--text2:#3A4A5E;--muted:#5C6A7B;--teal:#1B7169;--amber:#F2A541;--card:#fff;--line:#DBDBDB;
--bad:#B3261E;--warn:#8A5A00;color-scheme:light}
*{box-sizing:border-box}html{-webkit-print-color-adjust:exact;print-color-adjust:exact}
body{margin:0;background:var(--mist);color:var(--navy);font:16px/1.55 "Report Text",sans-serif}
.in{max-width:1120px;margin:0 auto;padding:0 24px}h1,h2,h3{font-family:"Report Display",sans-serif;font-weight:800;letter-spacing:-.02em;line-height:1.1;margin:0}
.eyebrow{font-size:13px;letter-spacing:.2em;text-transform:uppercase;color:var(--teal);font-weight:600;margin:0}
.top{background:var(--navy);color:var(--mist);padding:28px 0 44px}.top .eyebrow{color:#3CCFBF}.top h1{font-size:clamp(34px,5vw,56px);margin:10px 0 14px}
.top p{color:#C9D2DC;margin:0;max-width:720px}.brand{display:flex;align-items:center;gap:12px;font:800 20px "Report Display",sans-serif;margin-bottom:44px}
.brand i{display:grid;place-items:center;width:36px;height:36px;border-radius:8px;background:#000;border:1px solid #22344D;color:#fff;font-style:normal}
.brand span{margin-left:auto;font:600 14px "Report Text",sans-serif;color:#9FB0C2}
h2{font-size:30px;margin:56px 0 10px}h3{font-size:21px;margin:30px 0 8px}p{margin:8px 0}.muted{color:var(--muted)}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px 24px;margin:14px 0}.card>b{font-size:17px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin:16px 0}.lift{margin-top:-28px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px 20px}
.stat b{display:block;font:800 36px/1 "Report Display",sans-serif;font-variant-numeric:tabular-nums;margin-bottom:8px}.stat span{color:var(--text2);font-size:15px}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:14px;background:var(--card);margin:12px 0}
table{border-collapse:collapse;width:100%;font-size:15px}th,td{text-align:left;padding:11px 14px;border-bottom:1px solid #E2E2E2;vertical-align:top}
th{color:var(--teal);font-weight:600;white-space:nowrap;border-bottom:2px solid var(--navy)}td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
tr:last-child td{border-bottom:0}code{font:13.5px ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;overflow-wrap:anywhere}td:first-child code{display:block;min-width:150px}
.pill{display:inline-block;font-size:13px;padding:1px 10px;border-radius:99px;border:1px solid var(--line);color:var(--text2);white-space:nowrap}
.sev1,.sev2{color:var(--bad);border-color:var(--bad)}.sev3,.sev4{color:var(--warn);border-color:var(--warn)}.new{color:var(--teal);border-color:var(--teal)}
.uc{border-left:6px solid var(--teal)}.uc h3{margin:6px 0 4px;font-size:24px}.ask{border-left:4px solid var(--amber);padding-left:12px;margin:12px 0 0}
ol,ul{padding-left:22px;margin:8px 0}li{margin:4px 0}.end{background:var(--navy);color:var(--mist);margin-top:72px;padding:48px 0}.end h2{margin:0 0 12px}
.end a{color:#3CCFBF}.end .btn{display:inline-block;background:var(--amber);color:var(--navy);font-weight:600;border-radius:8px;padding:14px 26px;text-decoration:none;margin:14px 0}
.end small{color:#9FB0C2;font-size:14px;display:block;margin-top:18px}
@media print{.card,.stat,tr{break-inside:avoid}h2,h3{break-after:avoid}}
"""


def font_css():
    """The fonts, inside the file: a report loads nothing from a network."""
    out = []
    for family, weight, name in (("Report Text", 400, pdfmini.REGULAR), ("Report Text", 600, pdfmini.BOLD),
                                 ("Report Display", 800, pdfmini.DISPLAY)):
        data = base64.b64encode((pdfmini.FONTS / f"{name}.ttf").read_bytes()).decode()
        out.append(f"@font-face{{font-family:'{family}';font-weight:{weight};src:url(data:font/ttf;base64,{data}) format('truetype')}}")
    return "".join(out)


def counts(m):
    c = Counter("formula" if x["category"] == "excel" else "mismatch" if x["kind"] in DEFN_KINDS else "other" for x in m["entries"])
    return [(c["mismatch"], "definition and report mismatches"), (len(m["excel_groups"]), "spreadsheet problems"),
            (c["other"], "record, report and anomaly findings"), (len(m["questions"]), "questions to settle")]


def long_date(iso):
    d = date.fromisoformat(str(iso)[:10])
    return f"{d:%B} {d.day}, {d.year}"


def render_html(m):
    H = []
    a = H.append
    a(f"<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
      f"<title>Metric Mismatch Report: {e(m['plant'])}</title><style>{font_css()}{CSS}</style></head><body>")
    a(f"<header class='top'><div class='in'><div class='brand'><i>H</i>Humagined<span>Plant Watch</span></div>"
      f"<p class='eyebrow'>Metric Mismatch Report</p><h1>{e(m['plant'])}</h1>"
      f"<p>Data through {e(long_date(m['as_of']))}. {len(m['sources'])} sources and {len(m['audits'])} workbooks checked. "
      "Findings are computed from the files listed here, not estimated.</p></div></header><main class='in'>")
    a("<div class='grid lift'>" + "".join(f"<div class='stat'><b>{n:,}</b><span>{e(t)}</span></div>" for n, t in counts(m)) + "</div>")
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
            a("</table></div>" + (f"<p class='ask'><b>To settle:</b> {e(x['question'])}</p>" if x.get("question") else "") + "</div>")

    # 3. Suggested first use cases
    a("<h2>Where to start</h2><p class='muted'>Ranked by contested definitions on headline KPIs, then the size of "
      "the gaps, then spreadsheet faults feeding those numbers, then other problems.</p>")
    for i, uc in enumerate(m["use_cases"], 1):
        a(f"<div class='card uc'><div class='eyebrow'>{i}. {e(uc['stream'])}</div><h3>{e(uc['title'])}</h3>"
          f"<p>{e(uc['outcome'])}</p><p><b>Why here:</b></p><ul>" + "".join(f"<li>{e(w)}</li>" for w in uc["why"]) + "</ul>"
          + (f"<p><b>Makes trustworthy:</b> {e(', '.join(uc['kpis']))}</p>" if uc["kpis"] else "")
          + (f"<p><b>Definitions to settle:</b></p><ul>" + "".join(f"<li>{e(q)}</li>" for q in uc["questions"]) + "</ul>"
             if uc["questions"] else "")
          + (f"<p class='muted'>Sources involved: {e(', '.join(uc['sources']))}</p>" if uc["sources"] else "") + "</div>")

    # 4. By stream › stage › metric
    a("<h2>Mismatches by value stream</h2>")
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
                    a(f"<p class='ask'><b>To settle:</b> {e(x['question'])}</p>")
                a("</div>")

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
        a("<h3>Fix these first</h3><div class='tw'><table><tr><th>Where</th><th>Problem</th><th>Fix</th></tr>")
        for g in worst:
            a(f"<tr><td><code>{e(g['label'])}</code></td><td>{e(g['problem'])}</td><td>{e(g['fix'])}</td></tr>")
        a("</table></div>")

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
    a(f"</main><section class='end'><div class='in'><h2>Let's walk through them.</h2><p>{e(BOOKING).replace(BOOKING_URL, LINK)}</p>"
      f"<a class='btn' href='{BOOKING_URL}'>Book a Call &rarr;</a>"
      "<small>Generated by Plant Watch from the plant's own files. Plant Watch doesn't send your files or results. "
      "Check findings before you act on them.</small></div></section></body></html>")
    return "".join(H)


NAVY, TEAL, AMBER, MIST, MUTED, TEXT2, LINE = "#0E1A2B", "#1B7169", "#F2A541", "#F6F6F6", "#5C6A7B", "#3A4A5E", "#DBDBDB"
LEFT, RIGHT, FLOOR = 48, 564, 62


class Doc:
    """A page cursor: each piece moves down and starts a new page when it won't fit."""

    def __init__(self, m):
        self.m, self.pages, self.y = m, [], 0

    def page(self, first=False):
        self.p = pdfmini.Page()
        self.pages.append(self.p)
        self.p.rect(0, 0, 612, 792, MIST)
        if not first:
            self.p.text(LEFT, 756, "Metric Mismatch Report", 8.5, bold=True, color=TEAL)
            self.p.text(RIGHT, 756, f"{self.m['plant']}, data through {long_date(self.m['as_of'])}", 8.5, right=True, color=MUTED)
            self.p.line(LEFT, 748, RIGHT, 748)
            self.y = 726

    def room(self, h):
        if self.y - h < FLOOR:
            self.page()

    def gap(self, n):
        self.y -= n

    def text(self, t, size=10, bold=False, color=TEXT2, x=LEFT, limit=None, lead=1.4):
        for ln in pdfmini.wrap(t, size, bold, (limit or RIGHT) - x):
            self.room(size)
            self.p.text(x, self.y - size, ln, size, bold=bold, color=color)
            self.y -= size * lead

    def h2(self, t):
        self.room(90)
        self.gap(16)
        self.text(t, 18, pdfmini.DISPLAY, NAVY, lead=1.25)
        self.gap(6)

    def block(self, parts, bar=None):
        """A card: parts are (text, size, bold, color). Kept whole when it fits on a page."""
        inner = RIGHT - LEFT - 32
        lines = [(ln, sz, b, c) for t, sz, b, c in parts for ln in pdfmini.wrap(t, sz, b, inner)]
        while lines:
            h = sum(sz * 1.4 for _, sz, _, _ in lines) + 24
            if self.y - h < FLOOR and h < 600:
                self.page()
            fit, used = [], 24
            for ln in lines:
                if used + ln[1] * 1.4 > self.y - FLOOR and fit:
                    break
                fit.append(ln)
                used += ln[1] * 1.4
            self.p.rect(LEFT, self.y - used, RIGHT - LEFT, used, "#FFFFFF", LINE)
            if bar:
                self.p.rect(LEFT, self.y - used, 4, used, bar)
            y = self.y - 12
            for ln, sz, b, c in fit:
                self.p.text(LEFT + 18, y - sz, ln, sz, bold=b, color=c)
                y -= sz * 1.4
            self.y -= used + 8
            lines = lines[len(fit):]
            if lines:
                self.page()

    def table(self, heads, rows, widths):
        xs, x = [], LEFT
        for w in widths:
            xs.append(x)
            x += w

        def head():
            for hx, t in zip(xs, heads):
                self.p.text(hx + 6, self.y - 11, t, 9, bold=True, color=TEAL)
            self.p.line(LEFT, self.y - 17, RIGHT, self.y - 17, 1.2, NAVY)
            self.y -= 19

        self.room(70)
        head()
        for row in rows:
            cells = [pdfmini.wrap(t, 8.5, i == 0, w - 12) for i, (t, w) in enumerate(zip(row, widths))]
            h = max(len(c) for c in cells) * 11.5 + 9
            if self.y - h < FLOOR and h < 600:
                self.page()
                head()
            self.p.rect(LEFT, self.y - h, RIGHT - LEFT, h, "#FFFFFF")
            for i, (cx, c) in enumerate(zip(xs, cells)):
                for k, ln in enumerate(c):
                    self.p.text(cx + 6, self.y - 13 - k * 11.5, ln, 8.5, bold=i == 0, color=NAVY if i == 0 else TEXT2)
            self.p.line(LEFT, self.y - h, RIGHT, self.y - h)
            self.y -= h
        self.gap(10)


def render_pdf(m):
    d = Doc(m)
    d.page(first=True)
    p = d.p
    p.rect(0, 570, 612, 222, NAVY)
    p.rect(LEFT, 733, 24, 24, "#000000", "#22344D")
    p.text(LEFT + 7, 740, "H", 13, pdfmini.DISPLAY, color="#FFFFFF")
    p.text(LEFT + 34, 740, "Humagined", 14, pdfmini.DISPLAY, color=MIST)
    p.text(RIGHT, 740, "Plant Watch", 10, bold=True, right=True, color="#9FB0C2")
    p.text(LEFT, 686, "METRIC MISMATCH REPORT", 9.5, bold=True, color="#3CCFBF")
    y = 650
    for ln in pdfmini.wrap(m["plant"], 32, pdfmini.DISPLAY, RIGHT - LEFT)[:2]:
        p.text(LEFT, y, ln, 32, pdfmini.DISPLAY, color=MIST)
        y -= 34
    p.text(LEFT, y + 4, f"Data through {long_date(m['as_of'])}. {len(m['sources'])} sources and {len(m['audits'])} workbooks checked.",
           10.5, color="#C9D2DC")
    w = (RIGHT - LEFT - 30) / 4
    for i, (n, t) in enumerate(counts(m)):
        x = LEFT + i * (w + 10)
        p.rect(x, 512, w, 84, "#FFFFFF", LINE)
        p.text(x + 12, 562, f"{n:,}", 25, pdfmini.DISPLAY)
        for k, ln in enumerate(pdfmini.wrap(t, 9, False, w - 20)):
            p.text(x + 12, 545 - k * 11.5, ln, 9, color=TEXT2)
    d.y = 500

    if m["formula_defs"]:
        d.h2("Same name, different calculation")
        for x in m["formula_defs"]:
            d.block([(x["title"], 11, True, NAVY)]
                    + [(f"= {v['definition']}   ({v['places']} places, {v['workbooks']} workbooks; e.g. {v['example']})"
                        + ("  [your definition]" if v.get("is_yours") else ""), 9, False, TEXT2)
                       for v in (x.get("evidence") or {}).get("variants", [])]
                    + ([(f"To settle: {x['question']}", 9.5, True, TEAL)] if x.get("question") else []))
    if m["use_cases"]:
        d.h2("Where to start")
        for i, uc in enumerate(m["use_cases"], 1):
            d.block([(f"{i}. {uc['stream']}", 9, True, TEAL), (uc["title"], 14, pdfmini.DISPLAY, NAVY), (uc["outcome"], 10, False, TEXT2)]
                    + [(f"•  {why}", 9.5, False, TEXT2) for why in uc["why"]], bar=TEAL)
    shown = [(r, [x for x in r["findings"] if x["category"] != "excel" and x["kind"] != "formula_definition_conflict"]) for r in m["rows"]]
    if any(fs for _, fs in shown):
        d.h2("Mismatches by value stream")
        for r, fs in shown:
            for x in fs:
                d.block([(stage_name(x.get("stage")), 9, True, TEAL), (x["title"], 11, True, NAVY), (x["detail"], 9.5, False, TEXT2)]
                        + ([(f"To settle: {x['question']}", 9.5, True, NAVY)] if x.get("question") else []),
                        bar=AMBER if x["kind"] in DEFN_KINDS else None)
    if m["rows"]:
        d.h2("Value stream summary")
        d.table(["Value stream", "Contested", "Mismatches", "Spreadsheet", "Records", "Anomalies"],
                [[r["stream"]["title"], ", ".join(detect.METRICS.get(x, {}).get("label", x) for x in r["contested"]) or "-",
                  str(r["mismatches"]), str(r["excel"]), str(r["keys"]), str(r["anomalies"])] for r in m["rows"]],
                [156, 112, 66, 66, 56, 60])
    G = m["excel_groups"]
    d.h2("Excel errors at a glance")
    d.text(f"{len(G):,} problem(s) across {sum(g['count'] for g in G):,} cell(s). " +
           "  ".join(f"{name}: {sum(1 for g in G if g['kind'] in kinds)}." for name, kinds in EXCEL_GROUPS))
    for au in m["audits"]:
        gs = [g for g in G if g["workbook"] == au["workbook"]]
        d.h2(f"Excel audit: {au['workbook']}")
        d.text(f"{len(gs):,} problems, {sum(g['count'] for g in gs):,} cells.")
        if au.get("copies"):
            d.text(f"Checked at {au.get('path') or au['workbook']}. The same workbook is also at {'; '.join(au['copies'])} "
                   "(same sheets and formulas), so it wasn't checked twice.", 9, color=MUTED)
        d.gap(6)
        if gs:
            d.table(["Cell", "Formula", "Problem", "Fix"],
                    [[f"{g.get('sheet') or 'workbook'}!{g['where']}", g.get("formula") or "", g["problem"], g["fix"]]
                     for g in gs[:EXCEL_ROWS_PER_WORKBOOK]], [92, 118, 172, 134])
        if len(gs) > EXCEL_ROWS_PER_WORKBOOK:
            d.text(f"And {len(gs) - EXCEL_ROWS_PER_WORKBOOK:,} more problem(s); export the findings for the rest.", 9, color=MUTED)
    if m["questions"]:
        d.h2("Questions to settle")
        for i, q in enumerate(m["questions"], 1):
            d.text(f"{i}. {q['question']}", 10)
            d.gap(3)
    d.gap(14)
    lines = pdfmini.wrap(BOOKING, 11, True, RIGHT - LEFT - 40)
    h = len(lines) * 16 + 32
    d.room(h)
    d.p.rect(LEFT, d.y - h, RIGHT - LEFT, h, NAVY)
    for k, ln in enumerate(lines):
        d.p.text(LEFT + 20, d.y - 27 - k * 16, ln, 11, bold=True, color=MIST)
    d.p.link(LEFT, d.y - h, RIGHT - LEFT, h, BOOKING_URL)
    for i, pg in enumerate(d.pages, 1):
        pg.text(LEFT, 32, "Humagined", 9, pdfmini.DISPLAY)
        pg.text(RIGHT, 32, f"Page {i} of {len(d.pages)}", 8.5, right=True, color=MUTED)
    return pdfmini.pdf(d.pages)


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

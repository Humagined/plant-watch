"""Hand-built .xlsx files (raw XML, standard library) for parser cases openpyxl can't write:
shared formulas, external-link parts, cached errors, hidden sheets."""

import zipfile
from xml.sax.saxutils import escape

CT = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{overrides}
</Types>"""
MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' \
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
RELNS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'


def cell_xml(ref, spec):
    """spec: number | 'text' | {'f': formula, 'shared': (si, range or None), 'v': cached, 'e': error}"""
    if isinstance(spec, (int, float)):
        return f'<c r="{ref}"><v>{spec}</v></c>'
    if isinstance(spec, str):
        return f'<c r="{ref}" t="inlineStr"><is><t>{spec}</t></is></c>'
    attrs, inner = "", ""
    if "f" in spec:
        sh = spec.get("shared")
        if sh:
            si, rng = sh
            inner += f'<f t="shared" si="{si}"' + (f' ref="{rng}">{escape(spec["f"])}</f>' if rng else "/>")
        else:
            inner += f"<f>{escape(spec['f'])}</f>"
    if "e" in spec:
        attrs = ' t="e"'
        inner += f"<v>{spec['e']}</v>"
    elif "v" in spec:
        inner += f"<v>{spec['v']}</v>"
    return f'<c r="{ref}"{attrs}>{inner}</c>'


def build(path, sheets, names=None, external=None):
    """sheets: [(name, {cell: spec}, {'hidden_rows': [..], 'state': 'hidden'})]"""
    overrides, wb_sheets, rels = [], [], []
    with zipfile.ZipFile(path, "w") as z:
        for i, (name, cells, opts) in enumerate(sheets, 1):
            rows = {}
            for ref, spec in cells.items():
                r = int("".join(ch for ch in ref if ch.isdigit()))
                rows.setdefault(r, []).append(cell_xml(ref, spec))
            hidden = set((opts or {}).get("hidden_rows", []))
            flag = {True: ' hidden="1"', False: ""}
            data = "".join(f'<row r="{r}"{flag[r in hidden]}>{"".join(cs)}</row>'
                           for r, cs in sorted(rows.items()))
            z.writestr(f"xl/worksheets/sheet{i}.xml",
                       f'<?xml version="1.0" encoding="UTF-8"?><worksheet {MAIN}><sheetData>{data}</sheetData></worksheet>')
            overrides.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            state = (opts or {}).get("state")
            wb_sheets.append(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"' + (f' state="{state}"' if state else "") + "/>")
            rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
        if external:
            z.writestr("xl/externalLinks/externalLink1.xml",
                       f'<?xml version="1.0" encoding="UTF-8"?><externalLink {MAIN}><externalBook r:id="rId1"/></externalLink>')
            z.writestr("xl/externalLinks/_rels/externalLink1.xml.rels",
                       f'<?xml version="1.0" encoding="UTF-8"?><Relationships {RELNS}><Relationship Id="rId1" '
                       f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
                       f'Target="{external}" TargetMode="External"/></Relationships>')
        dn = "".join(f'<definedName name="{k}">{v}</definedName>' for k, v in (names or {}).items())
        z.writestr("xl/workbook.xml", f'<?xml version="1.0" encoding="UTF-8"?><workbook {MAIN}><sheets>{"".join(wb_sheets)}</sheets>'
                                      + (f"<definedNames>{dn}</definedNames>" if dn else "") + "</workbook>")
        z.writestr("xl/_rels/workbook.xml.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships {RELNS}>{"".join(rels)}</Relationships>')
        z.writestr("[Content_Types].xml", CT.format(overrides="".join(overrides)))
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships {RELNS}><Relationship Id="rId1" '
                                  f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    return path

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""STH-group Financial Dashboard Generator (v3 — multi-year, raw PL rows).
Usage: python generate_dashboard.py

Files are loaded from data/ subfolder (relative to this script).
If data/ is empty or missing, falls back to FALLBACK_FILES list.
"""

import zipfile, xml.etree.ElementTree as ET, json, re, os, glob

BASEDIR  = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASEDIR, 'data')
OUTPUT   = os.path.join(BASEDIR, 'sth_dashboard.html')

# Fallback paths (used when data/ folder is empty)
FALLBACK_FILES = [
    os.path.join(BASEDIR, '02-2026_\u0440\u0435\u043d\u0442\u0430\u0431\u0435\u043b\u044c\u043d\u043e\u0441\u0442\u044c_STH  (\u043e\u0442 15.04.2026)_\u0441\u0440\u0435\u0437\u044b.xlsx'),
    r'C:\Users\user\Downloads\12-2024_\u0440\u0435\u043d\u0442\u0430\u0431\u0435\u043b\u044c\u043d\u043e\u0441\u0442\u044c_STH (\u043e\u0442 18.02.2025)_\u043f\u043e\u043b\u043d\u0430\u044f \u0432\u0435\u0440\u0441\u0438\u044f.xlsx',
]

NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PNS = 'http://schemas.openxmlformats.org/package/2006/relationships'

MONTH_LABELS = {
    '01': "\u044f\u043d\u0432", '02': "\u0444\u0435\u0432", '03': "\u043c\u0430\u0440",
    '04': "\u0430\u043f\u0440", '05': "\u043c\u0430\u0439", '06': "\u0438\u044e\u043d",
    '07': "\u0438\u044e\u043b", '08': "\u0430\u0432\u0433", '09': "\u0441\u0435\u043d",
    '10': "\u043e\u043a\u0442", '11': "\u043d\u043e\u044f", '12': "\u0434\u0435\u043a",
}

# ── Excel reader ───────────────────────────────────────────────────────────────

def parse_ref(ref):
    m = re.match(r'^([A-Z]+)(\d+)$', ref)
    return (m.group(1), int(m.group(2))) if m else (None, None)

def norm_b(val):
    if val is None: return None
    if isinstance(val, float): return str(int(val)) if val == int(val) else str(val)
    return str(val).strip()

class XlReader:
    def __init__(self, path):
        self.path = path
        self.z  = zipfile.ZipFile(path)
        self.ss = self._strings()
        self.sm = self._sheets()

    def _strings(self):
        try:
            r = ET.fromstring(self.z.read('xl/sharedStrings.xml'))
            return [''.join(t.text or '' for t in si.findall(f'.//{{{NS}}}t'))
                    for si in r.findall(f'.//{{{NS}}}si')]
        except: return []

    def _sheets(self):
        rels = {r.get('Id'): r.get('Target') for r in
                ET.fromstring(self.z.read('xl/_rels/workbook.xml.rels'))
                .findall(f'{{{PNS}}}Relationship')}
        wb = ET.fromstring(self.z.read('xl/workbook.xml'))
        return {s.get('name'): 'xl/' + rels[s.get(f'{{{RNS}}}id')]
                for s in wb.findall(f'.//{{{NS}}}sheet')
                if s.get(f'{{{RNS}}}id') in rels}

    def _v(self, c):
        t, v = c.get('t',''), c.find(f'{{{NS}}}v')
        if v is None or not v.text: return None
        if t == 's':
            i = int(v.text); return self.ss[i] if i < len(self.ss) else None
        try: return float(v.text)
        except: return v.text

    def read(self, name):
        path = self.sm.get(name,'')
        if not path: return {}
        try: ws = ET.fromstring(self.z.read(path))
        except: return {}
        data = {}
        for row in ws.findall(f'.//{{{NS}}}row'):
            rn = int(row.get('r',0)); data[rn] = {}
            for c in row.findall(f'{{{NS}}}c'):
                col, _ = parse_ref(c.get('r',''))
                if col:
                    v = self._v(c)
                    if v is not None: data[rn][col] = v
        return data

    def close(self): self.z.close()

# ── File loading ───────────────────────────────────────────────────────────────

def load_xl_files():
    paths = []
    if os.path.isdir(DATA_DIR):
        paths = sorted(glob.glob(os.path.join(DATA_DIR, '*.xlsx')))
    if not paths:
        paths = [p for p in FALLBACK_FILES if os.path.exists(p)]
    readers = []
    for path in paths:
        try:
            xl = XlReader(path)
            print(f"  Loaded: {os.path.basename(path)}")
            readers.append(xl)
        except Exception as e:
            print(f"  Warning: could not read {os.path.basename(path)}: {e}")
    return readers

# ── Sheet discovery ────────────────────────────────────────────────────────────

def collect_monthly_sheets(xl_readers, from_year=2024):
    """Find all MM-YYYY monthly sheets across all xl readers, sorted by date."""
    found = {}
    for xl in xl_readers:
        for name in xl.sm:
            m = re.match(r'^(\d{2})-(\d{4})$', name)
            if not m: continue
            mm, yyyy = m.group(1), m.group(2)
            if int(yyyy) < from_year: continue
            key = f'{mm}-{yyyy}'
            if key in found: continue
            yy = yyyy[2:]
            label = f"{MONTH_LABELS.get(mm, mm)}'{yy}"
            found[key] = (xl, key, label, int(yyyy)*100 + int(mm))
    return [(xl, key, label)
            for xl, key, label, _ in sorted(found.values(), key=lambda x: x[3])]

# ── Extraction helpers ─────────────────────────────────────────────────────────

def find_row(sd, b):
    for rn, cols in sd.items():
        if norm_b(cols.get('B')) == b: return rn
    return None

def g(sd, b, col='D', default=0.0):
    rn = find_row(sd, b)
    if rn is None: return default
    v = sd[rn].get(col)
    if v is None: return default
    try: return float(v)
    except: return default

def fix_margin(m, rev=0, fin=0):
    m = float(m) if m is not None else 0.0
    if abs(m) > 5: return fin/rev if rev > 0 else 0.0
    return m

def col_ord(c):
    """Excel column letter to ordinal: A=1, B=2, Z=26, AA=27..."""
    n = 0
    for ch in c.upper():
        n = n * 26 + (ord(ch) - 64)
    return n

def get_division_map(sd):
    """Try to read division→col mapping from rows 1 or 2 of a monthly sheet."""
    proj_cols = {col for col in sd.get(3, {}) if col not in ('A','B','C','D')}
    if not proj_cols:
        return {}
    for hrow_n in [2, 1]:
        hrow = sd.get(hrow_n, {})
        EXCEL_ERRORS = {'#REF!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!'}
        candidates = {col: str(v).strip() for col, v in hrow.items()
                     if col not in ('A','B','C','D')
                     and v and isinstance(v, str)
                     and str(v).strip() not in EXCEL_ERRORS
                     and not any(str(v).strip().startswith(e) for e in EXCEL_ERRORS)
                     and len(str(v).strip()) > 1
                     and not str(v).strip().replace(' ', '').replace('-', '').isdigit()}
        if len(candidates) < 1:
            continue
        # Sort by column ordinal and propagate to project columns
        div_sorted = sorted(candidates.items(), key=lambda x: col_ord(x[0]))
        proj_sorted = sorted(proj_cols, key=col_ord)
        col_to_div = {}
        di = 0
        for pc in proj_sorted:
            pn = col_ord(pc)
            while di + 1 < len(div_sorted) and col_ord(div_sorted[di + 1][0]) <= pn:
                di += 1
            if di < len(div_sorted):
                col_to_div[pc] = div_sorted[di][1]
        if col_to_div:
            return col_to_div
    return {}

# ── Data extraction ────────────────────────────────────────────────────────────

def extract_monthly(monthly_sheets):
    rows = []
    for xl, key, label in monthly_sheets:
        sd = xl.read(key)
        rev = g(sd,'1'); fin = g(sd,'18')
        # Если B='1' аномально мало — пересчитать из подстрок
        _rev11 = g(sd,'1.1')
        _rev12 = abs(g(sd,'1.2'))
        _rev_computed = _rev11 - _rev12
        if _rev_computed > rev * 10 and _rev_computed > 1_000_000:
            rev = _rev_computed
        zp  = g(sd,'2.1.1') or g(sd,'2.1')
        margin = fix_margin(g(sd,'19'), rev, fin)
        rows.append({
            'key': key, 'label': label,
            'revenue':        round(rev),
            'project_costs':  round(g(sd,'2')),
            'zp_massive':     round(zp),
            'zp_foremen':     round(g(sd,'2.1.2')),
            'social_tax':     round(g(sd,'2.3')),
            'housing':        round(g(sd,'2.4.1')),
            'meals':          round(g(sd,'2.4.2')),
            'transport':      round(g(sd,'2.4.3')),
            'workwear':       round(g(sd,'2.4.4')),
            'workplace_org':  round(g(sd,'2.4.5')),
            'recruit_svc':    round(g(sd,'2.4.6')),
            'subcontract':    round(g(sd,'2.4.7')),
            'hr_docs':        round(g(sd,'2.5')),
            'bonus_recruit':  round(g(sd,'2.9')),
            'ndfl':           round(g(sd,'2.10')),
            'commercial':     round(g(sd,'2.11')),
            'bank_factor':    round(g(sd,'2.13')),
            'vat':            round(g(sd,'2.19')),
            'income_tax':     round(g(sd,'2.21')),
            'fin_res_super':  round(g(sd,'4')),
            'fin_res_rrp':    round(g(sd,'8')),
            'fin_res_drp':    round(g(sd,'12')),
            'fin_res_biz':    round(fin),
            'margin_pct':     round(margin * 100, 2),
            'zp_rev_pct':     round(zp/rev*100, 2) if rev > 0 else 0,
        })
    return rows

def extract_pl_rows(monthly_sheets):
    """Extract all B/C/D rows from each monthly sheet (raw P&L structure)."""
    result = {}
    for xl, key, label in monthly_sheets:
        sd = xl.read(key)
        rows = []
        for rn in sorted(sd.keys()):
            b = norm_b(sd[rn].get('B'))
            if not b: continue
            if not re.match(r'^\d+(\.\d+)*$', b): continue
            c = sd[rn].get('C')
            d = sd[rn].get('D')
            if d is None or not isinstance(d, (int, float)): continue
            d = float(d)
            if d == 0.0: continue  # skip zeros — checkbox in UI shows hint
            label_str = str(c).strip() if c and isinstance(c, str) and str(c).strip() else b
            indent = b.count('.')
            rows.append({'b': b, 'label': label_str, 'value': round(d, 4), 'indent': indent})
        # Fix B='1' if it's anomalously small
        _r1  = next((r for r in rows if r['b']=='1'), None)
        _r11 = next((r for r in rows if r['b']=='1.1'), None)
        _r12 = next((r for r in rows if r['b']=='1.2'), None)
        if _r1 and _r11:
            _comp = _r11['value'] - (abs(_r12['value']) if _r12 else 0)
            if _comp > _r1['value'] * 10 and _comp > 1_000_000:
                _r1['value'] = round(_comp, 4)
        result[key] = rows
    return result

def extract_projects(monthly_sheets):
    result = {}
    for xl, key, label in monthly_sheets:
        sd = xl.read(key)
        rev_rn = find_row(sd,'1'); mar_rn = find_row(sd,'19'); fin_rn = find_row(sd,'18')
        if not rev_rn: continue
        rev_row = sd.get(rev_rn,{}); mar_row = sd.get(mar_rn,{}) if mar_rn else {}
        fin_row = sd.get(fin_rn,{}) if fin_rn else {}
        col_to_div = get_division_map(sd)
        projects = []
        for col, name in sd.get(3,{}).items():
            if col in ('A','B','C','D'): continue
            if not name or not isinstance(name,str) or not name.strip(): continue
            clean = name.replace('_', ' ').strip()
            rev = float(rev_row.get(col,0) or 0)
            if rev < 10000: continue
            mar = float(mar_row.get(col,0) or 0)
            fin = float(fin_row.get(col,0) or 0)
            mar = fix_margin(mar, rev, fin)
            projects.append({'name': clean, 'rev': round(rev), 'margin': round(mar*100, 2), 'rrp': col_to_div.get(col, '')})
        result[key] = sorted(projects, key=lambda x: -x['rev'])
    return result

def extract_svod(xl, sheet, top=25):
    sd = xl.read(sheet)
    if not sd: return []
    names = {col: v.strip() for col,v in sd.get(3,{}).items()
             if col not in ('A','B','C','D') and v and isinstance(v,str) and v.strip()}
    rev_rn = find_row(sd,'1'); mar_rn = find_row(sd,'19'); fin_rn = find_row(sd,'18')
    if not rev_rn: return []
    out = []
    for col, name in names.items():
        rev = float(sd[rev_rn].get(col,0) or 0)
        if rev < 100: continue
        mar = float((sd.get(mar_rn,{}) if mar_rn else {}).get(col,0) or 0)
        fin = float((sd.get(fin_rn,{}) if fin_rn else {}).get(col,0) or 0)
        mar = fix_margin(mar, rev, fin)
        out.append({'name': name, 'rev': round(rev), 'margin': round(mar*100, 2)})
    out.sort(key=lambda x: -x['rev'])
    return out[:top]

# ── HTML template ──────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="ru" data-theme="dark">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>STH-group | Финансовый дашборд</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0d0f1a;--card:#151929;--card2:#0f1120;--border:#252a45;
  --text:#e8eaf6;--muted:#7986cb;--accent:#5c6bc0;
  --green:#66bb6a;--red:#ef5350;--amber:#ffa726;--purple:#ab47bc;
  --teal:#26c6da;--grid:rgba(255,255,255,0.05);
  --tab-active:#5c6bc0;
}
[data-theme="light"]{
  --bg:#eef1fb;--card:#ffffff;--card2:#f4f6ff;--border:#c8d0ea;
  --text:#1a237e;--muted:#4a5490;--accent:#3949ab;
  --green:#2e7d32;--red:#c62828;--amber:#bf360c;--purple:#6a1b9a;
  --teal:#00695c;--grid:rgba(0,0,0,0.07);--tab-active:#3949ab;
}
[data-theme="light"] .hdr{background:#3949ab}
[data-theme="light"] .hdr .hdr-brand h1,[data-theme="light"] .hdr .sub,
[data-theme="light"] #hdr-report-name,[data-theme="light"] .hdr .btn-sm{color:#fff}
[data-theme="light"] .hdr .btn-sm{border-color:rgba(255,255,255,.4)}
[data-theme="light"] .hdr .btn-sm:hover{border-color:#fff;background:rgba(255,255,255,.15)}
[data-theme="light"] .tab-bar{background:#fff;border-color:#c8d0ea}
[data-theme="light"] .tab-btn{color:#4a5490}
[data-theme="light"] .tab-btn.active{color:#3949ab;border-bottom-color:#3949ab}
[data-theme="light"] .section-hdr{background:rgba(57,73,171,.08)}
[data-theme="light"] .section-hdr td{color:#3949ab}
[data-theme="light"] tr:hover td{background:rgba(57,73,171,.06)}
[data-theme="light"] .pl-top0-total{background:rgba(57,73,171,.06)}
[data-theme="light"] th{color:#4a5490}
[data-theme="light"] .kpi{border-left-color:var(--accent)}
[data-theme="light"] .home-card:hover{box-shadow:0 12px 32px rgba(57,73,171,.18)}
[data-theme="light"] .badge{background:#3949ab}
[data-theme="light"] select{background:#fff;border-color:#c8d0ea}
[data-theme="light"] #tt{background:#fff;color:#1a237e;border-color:#c8d0ea;box-shadow:0 4px 14px rgba(0,0,0,.15)}
[data-theme="light"] .chip-g{background:rgba(46,125,50,.12);color:#2e7d32}
[data-theme="light"] .chip-r{background:rgba(198,40,40,.1);color:#c62828}
[data-theme="light"] .chip-a{background:rgba(191,54,12,.1);color:#bf360c}
[data-theme="light"] .chip-b{background:rgba(57,73,171,.12);color:#3949ab}
[data-theme="light"] .pos{color:#2e7d32}
[data-theme="light"] .neg{color:#c62828}
[data-theme="light"] .up{color:#2e7d32}
[data-theme="light"] .down{color:#c62828}
[data-theme="light"] .alert-g{background:rgba(46,125,50,.06);border-color:rgba(46,125,50,.3)}
[data-theme="light"] .alert-r{background:rgba(198,40,40,.06);border-color:rgba(198,40,40,.3)}
[data-theme="light"] .alert-b{background:rgba(57,73,171,.06);border-color:rgba(57,73,171,.3)}
[data-theme="light"] .home-card{box-shadow:0 2px 12px rgba(0,0,0,.06)}
[data-theme="light"] .modal-box-light{background:#fff}
body{background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;font-size:14px;line-height:1.5;min-height:100vh}

/* Header */
.hdr{display:flex;align-items:center;justify-content:space-between;padding:12px 24px;
     background:var(--card);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:200}
.hdr-brand{display:flex;align-items:center;gap:10px}
.hdr-brand h1{font-size:17px;font-weight:700;letter-spacing:-.3px}
.hdr-brand .sub{font-size:12px;color:var(--muted)}
.hdr-right{display:flex;gap:8px;align-items:center}
.badge{background:var(--accent);color:#fff;font-size:11px;padding:2px 8px;border-radius:20px;font-weight:600;opacity:.9}
.btn-sm{background:transparent;border:1px solid var(--border);color:var(--muted);
        padding:5px 12px;border-radius:6px;cursor:pointer;font-size:12px;transition:.15s}
.btn-sm:hover{border-color:var(--accent);color:var(--accent)}

/* Tab bar */
.tab-bar{display:flex;align-items:center;background:var(--card);border-bottom:1px solid var(--border);
         padding:0 20px;overflow-x:auto;gap:2px;position:sticky;top:49px;z-index:190}
.tab-btn{background:transparent;border:none;border-bottom:2px solid transparent;
         color:var(--muted);padding:11px 14px;cursor:pointer;font-size:13px;white-space:nowrap;
         transition:.15s;font-weight:500}
.tab-btn:hover{color:var(--text)}
.tab-btn.active{color:var(--accent);border-bottom-color:var(--accent);font-weight:600}

/* Content */
.tab-content{display:none;padding:20px 24px;max-width:1600px;margin:0 auto}
.tab-content.active{display:block}
.gap{display:flex;flex-direction:column;gap:18px}
.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:18px}
.card-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.card-title{font-size:14px;font-weight:600;display:flex;align-items:center;gap:6px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px}
@media(max-width:1000px){.row2,.row3{grid-template-columns:1fr}}

/* KPI */
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
@media(max-width:1100px){.kpi-grid{grid-template-columns:repeat(2,1fr)}}
.kpi{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:16px 18px;
     border-left:3px solid var(--accent)}
.kpi-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px}
.kpi-value{font-size:26px;font-weight:700;line-height:1.1}
.kpi-delta{font-size:12px;margin-top:5px}
.kpi-sub{font-size:11px;color:var(--muted);margin-top:3px}
.up{color:var(--green)}.down{color:var(--red)}.neutral{color:var(--muted)}
.pos{color:var(--green)}.neg{color:var(--red)}

/* Charts */
.chart-wrap{position:relative;height:340px}
.chart-sm{position:relative;height:220px}
.chart-bar{position:relative;height:320px}
.chart-donut{position:relative;height:260px;width:260px;flex-shrink:0}

/* Controls */
.ctrl{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px}
select{background:var(--card2);border:1px solid var(--border);color:var(--text);
       padding:6px 10px;border-radius:6px;font-size:13px;cursor:pointer;outline:none}
select:focus{border-color:var(--accent)}
.lbl{font-size:12px;color:var(--muted)}
.sep{color:var(--border);font-size:20px;font-weight:300}

/* Tables */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{color:var(--muted);font-weight:500;font-size:11px;text-align:left;padding:7px 10px;
   border-bottom:1px solid var(--border);text-transform:uppercase;letter-spacing:.05em;white-space:nowrap}
td{padding:6px 10px;border-bottom:1px solid rgba(255,255,255,0.04)}
[data-theme="light"] td{border-bottom:1px solid rgba(0,0,0,0.05)}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(92,107,192,.05)}
.td-r{text-align:right;font-variant-numeric:tabular-nums}
.td-name{font-weight:500}
.section-hdr{background:rgba(92,107,192,.1);font-weight:700;font-size:12px;text-transform:uppercase;
             letter-spacing:.06em;color:var(--accent)}
.section-hdr td{padding:8px 10px}
.indent1{padding-left:24px!important}
.indent2{padding-left:38px!important}
.row-bold td{font-weight:600}
.row-total td{font-weight:700;background:rgba(92,107,192,.07)}

/* Chips */
.chip{display:inline-block;padding:1px 7px;border-radius:10px;font-size:11px;font-weight:600}
.chip-g{background:rgba(102,187,106,.15);color:var(--green)}
.chip-r{background:rgba(239,83,80,.15);color:var(--red)}
.chip-a{background:rgba(255,167,38,.15);color:var(--amber)}
.chip-b{background:rgba(92,107,192,.15);color:var(--accent)}

/* Deviation bar */
.devbar{display:flex;align-items:center;gap:5px}
.devbar-inner{height:5px;border-radius:3px;min-width:2px}

/* Alerts */
.alerts{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
@media(max-width:900px){.alerts{grid-template-columns:1fr}}
.alert-box{border-radius:8px;padding:12px 14px;border:1px solid}
.alert-g{background:rgba(102,187,106,.07);border-color:rgba(102,187,106,.3)}
.alert-r{background:rgba(239,83,80,.07);border-color:rgba(239,83,80,.3)}
.alert-b{background:rgba(92,107,192,.07);border-color:rgba(92,107,192,.3)}
.alert-title{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px}
.alert-val{font-size:18px;font-weight:700}
.alert-sub{font-size:11px;color:var(--muted);margin-top:2px}

/* P&L table */
.pl-code{font-family:monospace;font-size:11px;color:var(--muted);white-space:nowrap;width:52px;padding-right:4px!important}
.pl-label{cursor:help}
.pl-label:hover{color:var(--accent)}
.pl-val{text-align:right;font-family:monospace;font-size:13px}
.pl-share{text-align:right;font-size:12px;color:var(--muted);width:90px}
.pl-zero td{opacity:.35}
.pl-top0{font-weight:700}
.pl-top0-total{font-weight:700;background:rgba(92,107,192,.07)}

/* Custom tooltip */
#tt{position:fixed;background:#1e2235;color:#e8eaf6;border:1px solid #3a4270;
    padding:7px 11px;border-radius:7px;font-size:12px;line-height:1.5;max-width:320px;
    pointer-events:none;z-index:9999;display:none;box-shadow:0 4px 14px rgba(0,0,0,.5)}
[data-theme="light"] #tt{background:#fff;color:#1a237e;border-color:#c5cae9}
/* Toggle buttons */
.tog{background:none;border:none;color:var(--muted);cursor:pointer;padding:0 4px 0 0;
     font-size:10px;line-height:1;vertical-align:middle;transition:.1s}
.tog:hover{color:var(--text)}
/* Anomaly blocks */
.anom4{display:grid;grid-template-columns:1fr 1fr;gap:14px}
@media(max-width:900px){.anom4{grid-template-columns:1fr}}
.anom-block-title{font-size:11px;font-weight:700;text-transform:uppercase;
                  letter-spacing:.06em;margin-bottom:8px;display:flex;align-items:center;gap:5px}

/* Anomaly sections */
.anom-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:900px){.anom-grid{grid-template-columns:1fr}}
.anom-stats{font-size:12px;color:var(--muted);margin-bottom:10px}

.footer{text-align:center;padding:16px;color:var(--muted);font-size:11px;
        border-top:1px solid var(--border);margin-top:10px}
/* Margin footnote */
.alloc-note{font-size:9px;font-weight:400;color:var(--muted);vertical-align:super;
            cursor:help;margin-left:2px;border-bottom:1px dashed var(--muted)}
/* Multi-select */
select[multiple]{padding:0;min-height:100px;font-size:12px}
select[multiple] option{padding:5px 8px}
/* Proj link */
.proj-link{color:var(--text);text-decoration:none;cursor:pointer}
.proj-link:hover{color:var(--accent);text-decoration:underline}
/* Export button */
.export-btn{background:transparent;border:1px solid var(--border);color:var(--muted);
            padding:4px 10px;border-radius:5px;cursor:pointer;font-size:11px;transition:.15s}
.export-btn:hover{border-color:var(--accent);color:var(--accent)}
/* Threshold row */
.thresh-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;font-size:12px;color:var(--muted)}
.thresh-row input[type=number]{width:48px;background:var(--card2);border:1px solid var(--border);
  color:var(--text);padding:3px 6px;border-radius:5px;font-size:12px;outline:none;text-align:center}
/* Print */
@media print{
  body{background:#fff!important;color:#111!important}
  .hdr{position:static!important;background:#3949ab!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .hdr .hdr-brand h1,.hdr .sub,#hdr-report-name{color:#fff!important}
  .tab-bar,.btn-sm,.export-btn,.tog,#modal-overlay{display:none!important}
  .tab-content{display:none!important}
  .tab-content.active{display:block!important;padding:10px!important}
  .card{break-inside:avoid;border:1px solid #ccc!important;background:#fff!important}
  .kpi{border:1px solid #ccc!important;background:#f9f9f9!important;border-left:3px solid #3949ab!important}
  #page-home,#page-cf,#page-pl{display:none!important}
  #page-rent{display:block!important}
  .home-wrap,.placeholder-wrap{display:none!important}
}

/* Home page */
.home-wrap{display:flex;flex-direction:column;align-items:center;justify-content:center;
           min-height:calc(100vh - 52px);padding:40px 20px;text-align:center}
.home-logo-row{display:flex;align-items:center;gap:14px;margin-bottom:10px;justify-content:center}
.home-logo-row h1{font-size:32px;font-weight:800;letter-spacing:-.5px}
.home-tagline{font-size:15px;color:var(--muted);margin-bottom:52px;letter-spacing:.02em}
.home-cards{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;max-width:860px;width:100%}
@media(max-width:700px){.home-cards{grid-template-columns:1fr}}
.home-card{background:var(--card);border:1px solid var(--border);border-radius:16px;
           padding:36px 24px 30px;cursor:pointer;transition:.18s;position:relative;overflow:hidden}
.home-card:hover{border-color:var(--accent);transform:translateY(-3px);
                 box-shadow:0 12px 32px rgba(0,0,0,.25)}
.home-card-icon{font-size:40px;margin-bottom:18px;display:block}
.home-card-name{font-size:17px;font-weight:700;margin-bottom:8px}
.home-card-desc{font-size:12px;color:var(--muted);line-height:1.5}
.home-card-badge{position:absolute;top:14px;right:14px;font-size:10px;font-weight:600;
                 padding:2px 8px;border-radius:20px}
.badge-ready{background:rgba(102,187,106,.15);color:var(--green)}
.badge-soon{background:rgba(255,167,38,.12);color:var(--amber)}
.home-card-line{position:absolute;top:0;left:0;right:0;height:3px;border-radius:16px 16px 0 0}

/* Placeholder page */
.placeholder-wrap{display:flex;flex-direction:column;align-items:center;justify-content:center;
                  min-height:calc(100vh - 120px);gap:16px;color:var(--muted)}
.placeholder-wrap .ph-icon{font-size:56px}
.placeholder-wrap h2{font-size:20px;font-weight:700;color:var(--text)}
.placeholder-wrap p{font-size:13px;max-width:340px;text-align:center;line-height:1.6}
</style>
</head>
<body>

<!-- Header -->
<header class="hdr">
  <div id="hdr-home-state" class="hdr-brand">
    <span style="font-size:22px">&#128202;</span>
    <div>
      <h1>STH-group</h1>
      <span class="sub">&#1060;&#1080;&#1085;&#1072;&#1085;&#1089;&#1086;&#1074;&#1072;&#1103; &#1086;&#1090;&#1095;&#1105;&#1090;&#1085;&#1086;&#1089;&#1090;&#1100;</span>
    </div>
  </div>
  <div id="hdr-report-state" style="display:none;align-items:center;gap:12px">
    <button class="btn-sm" onclick="showPage('home')" style="font-size:13px">&#8592; &#1043;&#1083;&#1072;&#1074;&#1085;&#1072;&#1103;</button>
    <span id="hdr-report-name" style="font-weight:700;font-size:15px"></span>
  </div>
  <div class="hdr-right">
    <span class="badge" id="period-badge" style="display:none"></span>
    <button class="btn-sm" onclick="openChangelog()" data-tip="&#1048;&#1085;&#1092;&#1086;&#1088;&#1084;&#1072;&#1094;&#1080;&#1103; &#1086; &#1076;&#1072;&#1085;&#1085;&#1099;&#1093;">&#9432; &#1044;&#1072;&#1085;&#1085;&#1099;&#1077;</button>
    <button class="btn-sm" id="theme-btn" onclick="toggleTheme()">&#9728; &#1057;&#1074;&#1077;&#1090;&#1083;&#1072;&#1103;</button>
  </div>
</header>

<div id="tt"></div>

<!-- Tab bar (only for Рентабельность) -->
<nav class="tab-bar" id="main-tabs" style="display:none">
  <button class="tab-btn active" onclick="showTab('main',this)">&#127968;&nbsp;&#1043;&#1083;&#1072;&#1074;&#1085;&#1072;&#1103;</button>
  <button class="tab-btn" onclick="showTab('trends',this)">&#128200;&nbsp;&#1058;&#1088;&#1077;&#1085;&#1076;&#1099;</button>
  <button class="tab-btn" onclick="showTab('general',this)">&#128203;&nbsp;&#1054;&#1073;&#1097;&#1080;&#1081;</button>
  <button class="tab-btn" onclick="showTab('compare',this)">&#9878;&#65039;&nbsp;&#1057;&#1088;&#1072;&#1074;&#1085;&#1077;&#1085;&#1080;&#1077;</button>
  <button class="tab-btn" onclick="showTab('anomalies',this)">&#128269;&nbsp;&#1040;&#1085;&#1086;&#1084;&#1072;&#1083;&#1080;&#1080;</button>
  <button class="tab-btn" onclick="showTab('projects',this)">&#128193;&nbsp;&#1055;&#1088;&#1086;&#1077;&#1082;&#1090;&#1099;</button>
  <button class="tab-btn" onclick="showTab('divisions',this)">&#128100;&nbsp;&#1056;&#1056;&#1055;</button>
  <button class="tab-btn" onclick="showTab('clients',this)">&#127962;&nbsp;&#1050;&#1083;&#1080;&#1077;&#1085;&#1090;&#1099;&nbsp;&amp;&nbsp;&#1043;&#1086;&#1088;&#1086;&#1076;&#1072;</button>
</nav>

<!-- ═══ PAGE: ГЛАВНАЯ (лендинг) ════════════════════════════════════════════════ -->
<div id="page-home">
  <div class="home-wrap">
    <div class="home-logo-row">
      <span style="font-size:38px">&#128202;</span>
      <h1>STH-group</h1>
    </div>
    <p class="home-tagline">&#1060;&#1080;&#1085;&#1072;&#1085;&#1089;&#1086;&#1074;&#1072;&#1103; &#1086;&#1090;&#1095;&#1105;&#1090;&#1085;&#1086;&#1089;&#1090;&#1100; &#183; &#1042;&#1099;&#1073;&#1077;&#1088;&#1080;&#1090;&#1077; &#1088;&#1072;&#1079;&#1076;&#1077;&#1083;</p>
    <div class="home-cards">

      <div class="home-card" onclick="showPage('cf')">
        <div class="home-card-line" style="background:var(--teal)"></div>
        <span class="home-card-badge badge-soon">&#1057;&#1082;&#1086;&#1088;&#1086;</span>
        <span class="home-card-icon">&#128176;</span>
        <div class="home-card-name">CF &#1054;&#1090;&#1095;&#1105;&#1090;</div>
        <div class="home-card-desc">&#1054;&#1090;&#1095;&#1105;&#1090; &#1086; &#1076;&#1074;&#1080;&#1078;&#1077;&#1085;&#1080;&#1080; &#1076;&#1077;&#1085;&#1077;&#1078;&#1085;&#1099;&#1093; &#1089;&#1088;&#1077;&#1076;&#1089;&#1090;&#1074;. &#1040;&#1085;&#1072;&#1083;&#1080;&#1079; &#1087;&#1086;&#1089;&#1090;&#1091;&#1087;&#1083;&#1077;&#1085;&#1080;&#1081; &#1080; &#1074;&#1099;&#1087;&#1083;&#1072;&#1090; &#1087;&#1086; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;&#1072;&#1084;.</div>
      </div>

      <div class="home-card" onclick="showPage('rent')">
        <div class="home-card-line" style="background:var(--accent)"></div>
        <span class="home-card-badge badge-ready">&#1044;&#1086;&#1089;&#1090;&#1091;&#1087;&#1077;&#1085;</span>
        <span class="home-card-icon">&#128200;</span>
        <div class="home-card-name">&#1054;&#1090;&#1095;&#1105;&#1090; &#1087;&#1086; &#1088;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1080;</div>
        <div class="home-card-desc">&#1040;&#1085;&#1072;&#1083;&#1080;&#1079; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1086;&#1074; &#1080; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;&#1086;&#1074;: &#1074;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;, &#1088;&#1072;&#1089;&#1093;&#1086;&#1076;&#1099;, &#1088;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100;, &#1090;&#1088;&#1077;&#1085;&#1076;&#1099; &#1087;&#1086; 26 &#1084;&#1077;&#1089;&#1103;&#1094;&#1072;&#1084;.</div>
      </div>

      <div class="home-card" onclick="showPage('pl')">
        <div class="home-card-line" style="background:var(--purple)"></div>
        <span class="home-card-badge badge-soon">&#1057;&#1082;&#1086;&#1088;&#1086;</span>
        <span class="home-card-icon">&#128203;</span>
        <div class="home-card-name">P&amp;L &#1054;&#1090;&#1095;&#1105;&#1090;</div>
        <div class="home-card-desc">&#1055;&#1086;&#1083;&#1085;&#1099;&#1081; &#1086;&#1090;&#1095;&#1105;&#1090; &#1086; &#1087;&#1088;&#1080;&#1073;&#1099;&#1083;&#1103;&#1093; &#1080; &#1091;&#1073;&#1099;&#1090;&#1082;&#1072;&#1093;. &#1044;&#1077;&#1090;&#1072;&#1083;&#1100;&#1085;&#1072;&#1103; &#1089;&#1090;&#1088;&#1091;&#1082;&#1090;&#1091;&#1088;&#1072; &#1076;&#1086;&#1093;&#1086;&#1076;&#1086;&#1074; &#1080; &#1088;&#1072;&#1089;&#1093;&#1086;&#1076;&#1086;&#1074;.</div>
      </div>

    </div>
  </div>
</div>

<!-- ═══ PAGE: CF (заглушка) ════════════════════════════════════════════════════ -->
<div id="page-cf" style="display:none">
  <div class="placeholder-wrap">
    <span class="ph-icon">&#128176;</span>
    <h2>CF &#1054;&#1090;&#1095;&#1105;&#1090;</h2>
    <p>&#1056;&#1072;&#1079;&#1076;&#1077;&#1083; &#1074; &#1088;&#1072;&#1079;&#1088;&#1072;&#1073;&#1086;&#1090;&#1082;&#1077;. &#1047;&#1076;&#1077;&#1089;&#1100; &#1087;&#1086;&#1103;&#1074;&#1080;&#1090;&#1089;&#1103; &#1086;&#1090;&#1095;&#1105;&#1090; &#1086; &#1076;&#1074;&#1080;&#1078;&#1077;&#1085;&#1080;&#1080; &#1076;&#1077;&#1085;&#1077;&#1078;&#1085;&#1099;&#1093; &#1089;&#1088;&#1077;&#1076;&#1089;&#1090;&#1074; &#1087;&#1086;&#1089;&#1083;&#1077; &#1079;&#1072;&#1075;&#1088;&#1091;&#1079;&#1082;&#1080; &#1089;&#1086;&#1086;&#1090;&#1074;&#1077;&#1090;&#1089;&#1090;&#1074;&#1091;&#1102;&#1097;&#1080;&#1093; &#1076;&#1072;&#1085;&#1085;&#1099;&#1093;.</p>
  </div>
</div>

<!-- ═══ PAGE: РЕНТАБЕЛЬНОСТЬ ═══════════════════════════════════════════════════ -->
<div id="page-rent" style="display:none">

<!-- ═══ TAB: ГЛАВНАЯ ═══════════════════════════════════════════════════════════ -->
<div id="tab-main" class="tab-content active">
  <div class="gap">
    <div class="ctrl" style="margin-bottom:0">
      <span class="lbl">&#1055;&#1077;&#1088;&#1080;&#1086;&#1076;:</span>
      <select id="mainPeriodType"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year">&#1043;&#1086;&#1076;</option></select>
      <select id="mainPeriodVal"></select>
    </div>
    <div class="kpi-grid" id="kpi-main"></div>
    <div class="alerts" id="alerts-main"></div>
    <div class="card" style="padding:14px 18px">
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">
        <span style="font-size:13px;font-weight:600">&#9881;&#65039; &#1055;&#1086;&#1088;&#1086;&#1075;&#1080; &#1088;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1080;</span>
        <div class="thresh-row">
          <span>&#1061;&#1086;&#1088;&#1086;&#1096;&#1086; &#8805;</span>
          <input type="number" id="threshGood" value="15" min="0" max="100">
          <span>%</span>
          <span style="margin-left:4px">&#1044;&#1086;&#1087;&#1091;&#1089;&#1090;&#1080;&#1084;&#1086; &#8805;</span>
          <input type="number" id="threshWarn" value="8" min="0" max="100">
          <span>%</span>
          <button class="btn-sm" onclick="applyThresholds()">&#10003; &#1055;&#1088;&#1080;&#1084;&#1077;&#1085;&#1080;&#1090;&#1100;</button>
          <span style="font-size:11px">(&#1074;&#1083;&#1080;&#1103;&#1077;&#1090; &#1085;&#1072; &#1094;&#1074;&#1077;&#1090; &#1074; &#1074;&#1089;&#1077;&#1093; &#1074;&#1082;&#1083;&#1072;&#1076;&#1082;&#1072;&#1093;)</span>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-title" style="margin-bottom:12px">&#128201; &#1058;&#1088;&#1077;&#1085;&#1076; (&#1087;&#1086;&#1089;&#1083;&#1077;&#1076;&#1085;&#1080;&#1077; 6 &#1084;&#1077;&#1089;&#1103;&#1094;&#1077;&#1074;)</div>
      <div class="chart-sm"><canvas id="miniChart"></canvas></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: ТРЕНДЫ ════════════════════════════════════════════════════════════ -->
<div id="tab-trends" class="tab-content">
  <div class="gap">
    <!-- Date range selector -->
    <div class="ctrl">
      <span class="lbl">&#1055;&#1077;&#1088;&#1080;&#1086;&#1076;:</span>
      <span class="lbl">&#1054;&#1090;</span>
      <select id="trendsFrom"></select>
      <span class="lbl">&#1076;&#1086;</span>
      <select id="trendsTo"></select>
    </div>
    <!-- Standard charts -->
    <div class="card">
      <div class="card-title" style="margin-bottom:12px">&#128200; &#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;, &#1088;&#1072;&#1089;&#1093;&#1086;&#1076;&#1099; &#1080; &#1088;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100;<span class="alloc-note" data-tip="&#1056;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100; &#1091;&#1082;&#1072;&#1079;&#1072;&#1085;&#1072; &#1076;&#1086; &#1072;&#1083;&#1083;&#1086;&#1082;&#1072;&#1094;&#1080;&#1080; &#1086;&#1073;&#1097;&#1080;&#1093; &#1088;&#1072;&#1089;&#1093;&#1086;&#1076;&#1086;&#1074;">&#174;&#1072;</span></div>
      <div class="chart-wrap"><canvas id="trendsChart"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title" style="margin-bottom:12px">&#128188; &#1047;&#1055; / &#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072; &#1080; &#1092;&#1080;&#1085;&#1072;&#1085;&#1089;&#1086;&#1074;&#1099;&#1081; &#1088;&#1077;&#1079;&#1091;&#1083;&#1100;&#1090;&#1072;&#1090;</div>
      <div class="chart-wrap"><canvas id="zpChart"></canvas></div>
    </div>
    <!-- Custom metric -->
    <div class="card">
      <div class="card-hdr" style="margin-bottom:12px">
        <div class="card-title">&#43;&#128202; &#1055;&#1086;&#1082;&#1072;&#1079;&#1072;&#1090;&#1077;&#1083;&#1100; &#1080;&#1079; &#1086;&#1090;&#1095;&#1077;&#1090;&#1072;</div>
        <div class="ctrl" style="margin:0">
          <select id="customMetricSel" style="min-width:320px"></select>
          <select id="customMetricChart2Type">
            <option value="line">&#1051;&#1080;&#1085;&#1080;&#1103;</option>
            <option value="bar">&#1057;&#1090;&#1086;&#1083;&#1073;&#1094;&#1099;</option>
          </select>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="customMetricChart"></canvas></div>
    </div>
    <!-- Project / division trends -->
    <div class="card">
      <div class="card-hdr" style="margin-bottom:12px">
        <div class="card-title">&#128193;&#128506;&#65039; &#1058;&#1088;&#1077;&#1085;&#1076; &#1087;&#1086; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1072;&#1084; / &#1076;&#1080;&#1074;&#1080;&#1079;&#1080;&#1086;&#1085;&#1072;&#1084;</div>
        <div class="ctrl" style="margin:0;flex-wrap:wrap;gap:8px">
          <select id="trendsFilterMode">
            <option value="all">&#1042;&#1089;&#1077; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1099;</option>
            <option value="projects">&#1042;&#1099;&#1073;&#1088;&#1072;&#1085;&#1085;&#1099;&#1077; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1099;</option>
            <option value="division">&#1056;&#1056;&#1055;</option>
          </select>
          <div id="trendsFilterDetail" style="display:flex;gap:8px;flex-wrap:wrap"></div>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="trendsProjChart"></canvas></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: ОБЩИЙ ═════════════════════════════════════════════════════════════ -->
<div id="tab-general" class="tab-content">
  <div class="gap">
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">&#128203; &#1042;&#1089;&#1077; &#1087;&#1086;&#1082;&#1072;&#1079;&#1072;&#1090;&#1077;&#1083;&#1080; &#1079;&#1072; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;</div>
        <div class="ctrl" style="margin:0;gap:8px">
          <span class="lbl">&#1058;&#1080;&#1087;:</span>
          <select id="generalPeriodType"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year">&#1043;&#1086;&#1076;</option></select>
          <select id="generalPeriodVal"></select>
          <button class="export-btn" onclick="exportCurrentGeneral()">&#128229; CSV</button>
        </div>
      </div>
      <p style="font-size:11px;color:var(--muted);margin-bottom:10px">
        &#1053;&#1072;&#1074;&#1077;&#1076;&#1080;&#1090;&#1077; &#1085;&#1072; &#1085;&#1072;&#1079;&#1074;&#1072;&#1085;&#1080;&#1077; &#1087;&#1086;&#1083;&#1103; &mdash; &#1087;&#1086;&#1103;&#1074;&#1080;&#1090;&#1089;&#1103; &#1087;&#1086;&#1076;&#1089;&#1082;&#1072;&#1079;&#1082;&#1072; &#1089; &#1086;&#1087;&#1080;&#1089;&#1072;&#1085;&#1080;&#1077;&#1084;.
      </p>
      <div class="tbl-wrap"><table id="generalTable"></table></div>
    </div>
    <div class="card">
      <div class="card-title" style="margin-bottom:12px">&#128201; &#1042;&#1086;&#1076;&#1086;&#1087;&#1072;&#1076; P&amp;L &mdash; &#1089;&#1090;&#1088;&#1091;&#1082;&#1090;&#1091;&#1088;&#1072; &#1092;&#1080;&#1085;&#1072;&#1085;&#1089;&#1086;&#1074;&#1086;&#1075;&#1086; &#1088;&#1077;&#1079;&#1091;&#1083;&#1100;&#1090;&#1072;&#1090;&#1072;</div>
      <div class="chart-wrap"><canvas id="waterfallChart"></canvas></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: СРАВНЕНИЕ ═════════════════════════════════════════════════════════ -->
<div id="tab-compare" class="tab-content">
  <div class="gap">
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">&#9878;&#65039; &#1057;&#1088;&#1072;&#1074;&#1085;&#1077;&#1085;&#1080;&#1077; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;&#1086;&#1074;</div>
        <div class="ctrl" style="margin:0;flex-wrap:wrap">
          <select id="cmpTypeA"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year">&#1043;&#1086;&#1076;</option></select>
          <select id="cmpValA"></select>
          <span class="sep">&#8594;</span>
          <select id="cmpTypeB"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year">&#1043;&#1086;&#1076;</option></select>
          <select id="cmpValB"></select>
        </div>
      </div>
      <div class="tbl-wrap"><div id="compareTable"></div></div>
    </div>
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">&#128101; &#1057;&#1088;&#1072;&#1074;&#1085;&#1077;&#1085;&#1080;&#1077; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1086;&#1074;</div>
        <div class="ctrl" style="margin:0">
          <input id="projectSearch" type="text" placeholder="&#1055;&#1086;&#1080;&#1089;&#1082; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1072;..."
            style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px;min-width:220px;outline:none">
        </div>
      </div>
      <div class="tbl-wrap"><div id="projectCompare"></div></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: АНОМАЛИИ ══════════════════════════════════════════════════════════ -->
<div id="tab-anomalies" class="tab-content">
  <div class="gap">
    <div class="ctrl">
      <span class="lbl">&#1055;&#1077;&#1088;&#1080;&#1086;&#1076;:</span>
      <select id="anomPeriodType"><option value="month" selected>&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year">&#1043;&#1086;&#1076;</option></select>
      <select id="anomPeriodVal"></select>
      <span id="anomaly-stats" class="anom-stats" style="margin-left:12px"></span>
      <button class="export-btn" onclick="exportCurrentAnomalies()">&#128229; CSV</button>
    </div>
    <!-- Row 1: Выручка -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <div class="card">
        <div class="anom-block-title">&#128176; &#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;</div>
        <div class="row2" style="gap:10px">
          <div><div style="font-size:11px;font-weight:600;color:var(--green);margin-bottom:6px">&#11014; &#1058;&#1054;&#1055; &#1087;&#1086; &#1074;&#1099;&#1088;&#1091;&#1095;&#1082;&#1077;</div><div id="anomRevTop"></div></div>
          <div><div style="font-size:11px;font-weight:600;color:var(--red);margin-bottom:6px">&#11015; &#1040;&#1091;&#1090;&#1089;&#1072;&#1081;&#1076;&#1077;&#1088;&#1099;</div><div id="anomRevBot"></div></div>
        </div>
      </div>
      <div class="card">
        <div class="anom-block-title">&#128200; &#1044;&#1080;&#1085;&#1072;&#1084;&#1080;&#1082;&#1072; &#1074;&#1099;&#1088;&#1091;&#1095;&#1082;&#1080;</div>
        <div class="row2" style="gap:10px">
          <div><div style="font-size:11px;font-weight:600;color:var(--green);margin-bottom:6px">&#11014; &#1056;&#1086;&#1089;&#1090;</div><div id="anomRevGrowth"></div></div>
          <div><div style="font-size:11px;font-weight:600;color:var(--red);margin-bottom:6px">&#11015; &#1057;&#1085;&#1080;&#1078;&#1077;&#1085;&#1080;&#1077;</div><div id="anomRevDrop"></div></div>
        </div>
      </div>
    </div>
    <!-- Row 2: Прибыль -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <div class="card">
        <div class="anom-block-title">&#128184; &#1055;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100; (&#1072;&#1073;&#1089;.)</div>
        <div class="row2" style="gap:10px">
          <div><div style="font-size:11px;font-weight:600;color:var(--green);margin-bottom:6px">&#11014; &#1053;&#1072;&#1080;&#1073;&#1086;&#1083;&#1077;&#1077; &#1087;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100;&#1085;&#1099;&#1077;</div><div id="anomProfTop"></div></div>
          <div><div style="font-size:11px;font-weight:600;color:var(--red);margin-bottom:6px">&#11015; &#1059;&#1073;&#1099;&#1090;&#1086;&#1095;&#1085;&#1099;&#1077;</div><div id="anomProfBot"></div></div>
        </div>
      </div>
      <div class="card">
        <div class="anom-block-title">&#128185; &#1044;&#1080;&#1085;&#1072;&#1084;&#1080;&#1082;&#1072; &#1087;&#1088;&#1080;&#1073;&#1099;&#1083;&#1080;</div>
        <div class="row2" style="gap:10px">
          <div><div style="font-size:11px;font-weight:600;color:var(--green);margin-bottom:6px">&#11014; &#1056;&#1086;&#1089;&#1090;</div><div id="anomProfGrowth"></div></div>
          <div><div style="font-size:11px;font-weight:600;color:var(--red);margin-bottom:6px">&#11015; &#1057;&#1085;&#1080;&#1078;&#1077;&#1085;&#1080;&#1077;</div><div id="anomProfDrop"></div></div>
        </div>
      </div>
    </div>
    <!-- Row 3: Рентабельность -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <div class="card">
        <div class="anom-block-title">&#128202; &#1056;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100;</div>
        <div class="row2" style="gap:10px">
          <div><div style="font-size:11px;font-weight:600;color:var(--green);margin-bottom:6px">&#11014; &#1042;&#1099;&#1089;&#1086;&#1082;&#1072;&#1103;</div><div id="anomMarTop"></div></div>
          <div><div style="font-size:11px;font-weight:600;color:var(--red);margin-bottom:6px">&#11015; &#1053;&#1080;&#1079;&#1082;&#1072;&#1103;</div><div id="anomMarBot"></div></div>
        </div>
      </div>
      <div class="card">
        <div class="anom-block-title">&#128260; &#1044;&#1080;&#1085;&#1072;&#1084;&#1080;&#1082;&#1072; &#1088;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1080;</div>
        <div class="row2" style="gap:10px">
          <div><div style="font-size:11px;font-weight:600;color:var(--green);margin-bottom:6px">&#11014; &#1056;&#1086;&#1089;&#1090;</div><div id="anomMarGrowth"></div></div>
          <div><div style="font-size:11px;font-weight:600;color:var(--red);margin-bottom:6px">&#11015; &#1057;&#1085;&#1080;&#1078;&#1077;&#1085;&#1080;&#1077;</div><div id="anomMarDrop"></div></div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ═══ TAB: ПРОЕКТЫ ═══════════════════════════════════════════════════════════ -->
<div id="tab-projects" class="tab-content">
  <div class="gap">
    <div class="ctrl">
      <span class="lbl">&#1055;&#1077;&#1088;&#1080;&#1086;&#1076;:</span>
      <select id="projPeriodType"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year" selected>&#1043;&#1086;&#1076;</option></select>
      <select id="projPeriodVal"></select>
      <input id="projSearchInput" type="text" placeholder="&#1055;&#1086;&#1080;&#1089;&#1082; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1072;..."
        style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px;min-width:240px;outline:none">
      <span id="proj-stats" style="font-size:12px;color:var(--muted)"></span>
      <button class="export-btn" onclick="exportCurrentProjects()">&#128229; CSV</button>
    </div>
    <div class="card">
      <div class="card-hdr" style="margin-bottom:10px">
        <div class="card-title">&#128193; &#1042;&#1089;&#1077; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1099; &#8212; &#1085;&#1072;&#1078;&#1084;&#1080;&#1090;&#1077; &#1085;&#1072; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090; &#1076;&#1083;&#1103; &#1087;&#1086;&#1076;&#1088;&#1086;&#1073;&#1085;&#1086;&#1089;&#1090;&#1077;&#1081;</div>
        <div style="font-size:11px;color:var(--muted)">&#1089;&#1086;&#1088;&#1090;&#1080;&#1088;&#1086;&#1074;&#1082;&#1072; &#1087;&#1086; &#1074;&#1099;&#1088;&#1091;&#1095;&#1082;&#1077;</div>
      </div>
      <div class="tbl-wrap"><div id="projTableDiv"></div></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: РРП ══════════════════════════════════════════════════════════════ -->
<div id="tab-divisions" class="tab-content">
  <div class="gap">
    <div class="ctrl">
      <span class="lbl">&#1055;&#1077;&#1088;&#1080;&#1086;&#1076;:</span>
      <select id="divPeriodType"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year" selected>&#1043;&#1086;&#1076;</option></select>
      <select id="divPeriodVal"></select>
      <span id="div-stats" style="font-size:12px;color:var(--muted)"></span>
      <button class="export-btn" onclick="exportCurrentDivisions()">&#128229; CSV</button>
    </div>
    <div class="row2">
      <div class="card">
        <div class="card-title" style="margin-bottom:12px">&#128100; &#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072; &#1087;&#1086; &#1056;&#1056;&#1055;</div>
        <div class="chart-wrap"><canvas id="divRevenueChart"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title" style="margin-bottom:12px">&#128200; &#1056;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100; &#1087;&#1086; &#1056;&#1056;&#1055;</div>
        <div class="chart-wrap"><canvas id="divMarginChart"></canvas></div>
      </div>
    </div>
    <div class="card">
      <div class="card-hdr" style="margin-bottom:10px">
        <div class="card-title">&#128203; &#1056;&#1056;&#1055; &mdash; &#1085;&#1072;&#1078;&#1084;&#1080;&#1090;&#1077; &#1085;&#1072; &#1056;&#1056;&#1055; &#1076;&#1083;&#1103; &#1087;&#1086;&#1076;&#1088;&#1086;&#1073;&#1085;&#1086;&#1089;&#1090;&#1077;&#1081;</div>
      </div>
      <div class="tbl-wrap"><div id="divTable"></div></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: КЛИЕНТЫ & ГОРОДА ══════════════════════════════════════════════════ -->
<div id="tab-clients" class="tab-content">
  <div class="gap">
    <div class="ctrl">
      <span class="lbl">&#1055;&#1077;&#1088;&#1080;&#1086;&#1076;:</span>
      <select id="clientsPeriodType"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year" selected>&#1043;&#1086;&#1076;</option></select>
      <select id="clientsPeriodVal"></select>
    </div>
    <div class="row2">
      <div class="card">
        <div class="card-title" style="margin-bottom:12px">&#127962; &#1058;&#1086;&#1087; &#1082;&#1083;&#1080;&#1077;&#1085;&#1090;&#1086;&#1074; <span id="clientsPeriodLabel" style="font-size:11px;font-weight:400;color:var(--muted)"></span></div>
        <div class="chart-bar"><canvas id="clientsChart"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title" style="margin-bottom:12px">&#128506;&#65039; &#1058;&#1086;&#1087; &#1075;&#1086;&#1088;&#1086;&#1076;&#1086;&#1074; <span id="citiesPeriodLabel" style="font-size:11px;font-weight:400;color:var(--muted)"></span></div>
        <div class="chart-bar"><canvas id="citiesChart"></canvas></div>
      </div>
    </div>
    <div class="card">
      <div class="card-hdr" style="margin-bottom:10px">
        <div class="card-title">&#127962; &#1042;&#1089;&#1077; &#1082;&#1083;&#1080;&#1077;&#1085;&#1090;&#1099; &mdash; &#1085;&#1072;&#1078;&#1084;&#1080;&#1090;&#1077; &#1076;&#1083;&#1103; &#1087;&#1086;&#1076;&#1088;&#1086;&#1073;&#1085;&#1086;&#1089;&#1090;&#1077;&#1081;</div>
        <button class="export-btn" onclick="exportCurrentClients()">&#128229; CSV</button>
      </div>
      <div class="tbl-wrap"><div id="clientsTableDiv"></div></div>
    </div>
  </div>
</div>

</div><!-- /page-rent -->

<!-- ═══ PAGE: P&L (заглушка) ══════════════════════════════════════════════════ -->
<div id="page-pl" style="display:none">
  <div class="placeholder-wrap">
    <span class="ph-icon">&#128203;</span>
    <h2>P&amp;L &#1054;&#1090;&#1095;&#1105;&#1090;</h2>
    <p>&#1056;&#1072;&#1079;&#1076;&#1077;&#1083; &#1074; &#1088;&#1072;&#1079;&#1088;&#1072;&#1073;&#1086;&#1090;&#1082;&#1077;. &#1047;&#1076;&#1077;&#1089;&#1100; &#1087;&#1086;&#1103;&#1074;&#1080;&#1090;&#1089;&#1103; &#1087;&#1086;&#1083;&#1085;&#1099;&#1081; &#1086;&#1090;&#1095;&#1105;&#1090; &#1086; &#1087;&#1088;&#1080;&#1073;&#1099;&#1083;&#1103;&#1093; &#1080; &#1091;&#1073;&#1099;&#1090;&#1082;&#1072;&#1093; &#1087;&#1086;&#1089;&#1083;&#1077; &#1079;&#1072;&#1075;&#1088;&#1091;&#1079;&#1082;&#1080; &#1076;&#1072;&#1085;&#1085;&#1099;&#1093;.</p>
  </div>
</div>

<!-- ═══ MODAL ════════════════════════════════════════════════════════════════ -->
<div id="modal-overlay" onclick="closeModal()"
  style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:800;overflow-y:auto;padding:30px 16px">
  <div id="modal-box" onclick="event.stopPropagation()"
    style="background:var(--card);border:1px solid var(--border);border-radius:14px;max-width:900px;
           margin:0 auto;padding:28px;position:relative">
    <button onclick="closeModal()"
      style="position:absolute;top:14px;right:16px;background:none;border:none;color:var(--muted);
             font-size:20px;cursor:pointer;line-height:1;padding:4px 8px">&times;</button>
    <div id="modal-content"></div>
  </div>
</div>

<div class="footer" id="footer-main">STH-group Financial Dashboard</div>

<script>
// ── Data ──────────────────────────────────────────────────────────────────────
const D = ___DATA___;

// ── Utils ─────────────────────────────────────────────────────────────────────
const fmtM   = v => v==null?'—': Math.abs(v)>=1e6 ? (v/1e6).toFixed(1)+'M' : Math.abs(v)>=1e3 ? Math.round(v/1e3)+'K' : Math.round(v);
const fmtRub = v => v==null?'—': Math.round(v).toLocaleString('ru-RU');
const fmtPct = v => v==null?'—': v.toFixed(1)+'%';
const sign   = v => v>0?'+':'';
const css    = p => getComputedStyle(document.documentElement).getPropertyValue(p).trim();

// ── Thresholds ────────────────────────────────────────────────────────────────
const THRESHOLDS = (function(){
  try{ const s=localStorage.getItem('sth_thr'); if(s) return JSON.parse(s); }catch(e){}
  return {good:15,warn:8};
})();

function chipCls(v){ return v>=THRESHOLDS.good?'chip-g':v>=THRESHOLDS.warn?'chip-a':'chip-r'; }

function applyThresholds(){
  THRESHOLDS.good = +document.getElementById('threshGood').value||15;
  THRESHOLDS.warn = +document.getElementById('threshWarn').value||8;
  try{ localStorage.setItem('sth_thr',JSON.stringify(THRESHOLDS)); }catch(e){}
  // Force re-render visible tab
  const active=document.querySelector('.tab-content.active');
  if(active){ const name=active.id.replace('tab-',''); inited[name]=false; if(initFns[name]){initFns[name]();inited[name]=true;} }
}

// ── CSV export ────────────────────────────────────────────────────────────────
function downloadCSV(headers, rows, filename){
  const bom='\uFEFF';
  const esc=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
  const csv=bom+[headers,...rows].map(r=>r.map(esc).join(';')).join('\n');
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=filename; a.click(); URL.revokeObjectURL(a.href);
}

function exportCurrentGeneral(){
  const type=document.getElementById('generalPeriodType').value;
  const val=document.getElementById('generalPeriodVal').value;
  const rows=aggregatePlRows(getMonthKeys(type,val)); if(!rows.length) return;
  downloadCSV(['Код','Показатель','Значение'],
    rows.map(r=>{const ip=Math.abs(r.value)<=1.5&&r.value!==0; return [r.b,r.label,ip?(r.value*100).toFixed(2)+'%':Math.round(r.value)];}),
    `pl_${val}.csv`);
}

function exportCurrentProjects(){
  const type=document.getElementById('projPeriodType').value;
  const val=document.getElementById('projPeriodVal').value;
  const fl=document.getElementById('projSearchInput').value.toLowerCase();
  const keys=getMonthKeys(type,val); const acc={};
  keys.forEach(k=>(D.projects_per_month[k]||[]).forEach(p=>{
    if(!acc[p.name])acc[p.name]={rev:0,fin:0};
    acc[p.name].rev+=p.rev; acc[p.name].fin+=p.rev*(p.margin/100);
  }));
  const rows=Object.entries(acc).map(([n,a])=>({n,rev:a.rev,fin:a.fin,mar:a.rev>0?a.fin/a.rev*100:0}))
    .filter(r=>!fl||r.n.toLowerCase().includes(fl)).sort((a,b)=>b.rev-a.rev);
  downloadCSV(['Проект','Выручка','Прибыль','Рентабельность %'],
    rows.map(r=>[r.n,Math.round(r.rev),Math.round(r.fin),r.mar.toFixed(1)]),
    `\u043f\u0440\u043e\u0435\u043a\u0442\u044b_${val}.csv`);
}

function exportCurrentDivisions(){
  const type=document.getElementById('divPeriodType').value;
  const val=document.getElementById('divPeriodVal').value;
  const divs=aggregateDivisions(getMonthKeys(type,val));
  downloadCSV(['\u0420\u0420\u041f','\u0412\u044b\u0440\u0443\u0447\u043a\u0430','\u041f\u0440\u0438\u0431\u044b\u043b\u044c','\u0420\u0435\u043d\u0442\u0430\u0431\u0435\u043b\u044c\u043d\u043e\u0441\u0442\u044c %','\u0413\u043e\u0440\u043e\u0434\u043e\u0432','\u041f\u0440\u043e\u0435\u043a\u0442\u043e\u0432'],
    divs.map(d=>[d.div,Math.round(d.revSum),Math.round(d.finSum),d.margin.toFixed(1),d.cities.length,d.projects.length]),
    `\u0434\u0438\u0432\u0438\u0437\u0438\u043e\u043d\u044b_${val}.csv`);
}

function exportCurrentAnomalies(){
  const type=document.getElementById('anomPeriodType').value;
  const val=document.getElementById('anomPeriodVal').value;
  const keys=getMonthKeys(type,val); const acc={};
  keys.forEach(k=>(D.projects_per_month[k]||[]).filter(p=>p.rev>0).forEach(p=>{
    if(!acc[p.name])acc[p.name]={rev:0,fin:0};
    acc[p.name].rev+=p.rev; acc[p.name].fin+=p.rev*(p.margin/100);
  }));
  const rows=Object.entries(acc).map(([name,a])=>({n:name,rev:a.rev,fin:a.fin,mar:a.rev>0?a.fin/a.rev*100:0}))
    .sort((a,b)=>b.rev-a.rev);
  downloadCSV(['\u041f\u0440\u043e\u0435\u043a\u0442','\u0412\u044b\u0440\u0443\u0447\u043a\u0430','\u041f\u0440\u0438\u0431\u044b\u043b\u044c','\u0420\u0435\u043d\u0442\u0430\u0431\u0435\u043b\u044c\u043d\u043e\u0441\u0442\u044c %'],
    rows.map(r=>[r.n,Math.round(r.rev),Math.round(r.fin),r.mar.toFixed(1)]),
    `\u0430\u043d\u043e\u043c\u0430\u043b\u0438\u0438_${val}.csv`);
}

// ── Changelog ─────────────────────────────────────────────────────────────────
function openChangelog(){
  const first=D.monthly[0], last=D.monthly[D.monthly.length-1];
  const now=new Date().toLocaleDateString('ru-RU');
  const projCnt=new Set(Object.values(D.projects_per_month).flat().map(p=>p.name)).size;
  document.getElementById('modal-content').innerHTML=`
    <h2 style="margin-bottom:12px">&#9432; &#1054; &#1076;&#1072;&#1096;&#1073;&#1086;&#1088;&#1076;&#1077;</h2>
    <div style="display:flex;flex-direction:column;gap:14px">
      <div class="alert-box alert-b">
        <div class="alert-title">&#128202; &#1044;&#1072;&#1085;&#1085;&#1099;&#1077;</div>
        <div class="alert-val" style="font-size:16px">${first.label} &#8594; ${last.label}</div>
        <div class="alert-sub">${D.monthly.length} &#1084;&#1077;&#1089;&#1103;&#1094;&#1077;&#1074; &#183; &#1086;&#1090;&#1082;&#1088;&#1099;&#1090;&#1086;: ${now}</div>
      </div>
      <ul style="list-style:disc;padding-left:20px;line-height:2;color:var(--muted);font-size:13px">
        <li>${D.monthly.length} &#1084;&#1077;&#1089;&#1103;&#1094;&#1077;&#1074; &#1092;&#1080;&#1085;&#1072;&#1085;&#1089;&#1086;&#1074;&#1099;&#1093; &#1076;&#1072;&#1085;&#1085;&#1099;&#1093;</li>
        <li>${Object.keys(D.pl_rows).length} &#1084;&#1077;&#1089;&#1103;&#1094;&#1077;&#1074; &#1089; P&amp;L-&#1089;&#1090;&#1088;&#1091;&#1082;&#1090;&#1091;&#1088;&#1086;&#1081;</li>
        <li>${projCnt} &#1091;&#1085;&#1080;&#1082;&#1072;&#1083;&#1100;&#1085;&#1099;&#1093; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1086;&#1074;</li>
        <li>${D.clients.length} &#1082;&#1083;&#1080;&#1077;&#1085;&#1090;&#1086;&#1074; &middot; ${D.cities.length} &#1075;&#1086;&#1088;&#1086;&#1076;&#1086;&#1074; &#1074; &#1057;&#1042;&#1054;&#1044;</li>
      </ul>
      <div style="font-size:11px;color:var(--muted);border-top:1px solid var(--border);padding-top:10px">
        STH-group Financial Dashboard &middot; &#1075;&#1077;&#1085;&#1077;&#1088;&#1072;&#1094;&#1080;&#1103; &#1072;&#1074;&#1090;&#1086;&#1084;&#1072;&#1090;&#1080;&#1095;&#1077;&#1089;&#1082;&#1072;&#1103; &#1080;&#1079; Excel-&#1092;&#1072;&#1081;&#1083;&#1086;&#1074;
      </div>
    </div>`;
  document.getElementById('modal-overlay').style.display='block';
}

// ── Period utilities ──────────────────────────────────────────────────────────
const MONTH_KEYS = D.monthly.map(m => m.key);

function getMonthKeys(type, val) {
  if (type === 'month') return [val];
  if (type === 'quarter') {
    const [q, yr] = val.split('-');
    const qn = +q[1];
    const ms = [(qn-1)*3+1,(qn-1)*3+2,(qn-1)*3+3]
      .map(m => String(m).padStart(2,'0')+'-'+yr);
    return ms.filter(k => MONTH_KEYS.includes(k));
  }
  if (type === 'year') return MONTH_KEYS.filter(k => k.slice(3) === val);
  return [];
}

function getPeriodLabel(type, val) {
  if (type === 'month') { const m = D.monthly.find(m => m.key===val); return m ? m.label : val; }
  if (type === 'quarter') return val.replace('-',' ');
  return val;
}

function fillPeriodSelect(selType, selVal, onChange) {
  const sKey = 'pd_' + selType.id;
  function rebuild() {
    const t = selType.value;
    selVal.innerHTML = '';
    if (t === 'month') {
      D.monthly.slice().reverse().forEach(m => {
        selVal.innerHTML += `<option value="${m.key}">${m.label}</option>`;
      });
    } else if (t === 'quarter') {
      const seen = new Set();
      D.monthly.slice().reverse().forEach(m => {
        const yr=m.key.slice(3), mo=+m.key.slice(0,2), qn=Math.ceil(mo/3);
        const qk=`Q${qn}-${yr}`;
        if (!seen.has(qk)) { seen.add(qk); selVal.innerHTML+=`<option value="${qk}">Q${qn} ${yr}</option>`; }
      });
    } else {
      const seen = new Set();
      D.monthly.slice().reverse().forEach(m => {
        const yr=m.key.slice(3);
        if (!seen.has(yr)) { seen.add(yr); selVal.innerHTML+=`<option value="${yr}">${yr}</option>`; }
      });
    }
    // Restore saved value for this type
    try {
      const sv=JSON.parse(sessionStorage.getItem(sKey)||'null');
      if(sv&&sv.t===t){const o=[...selVal.options].find(x=>x.value===sv.v);if(o)selVal.value=sv.v;}
    } catch(e){}
    onChange();
  }
  selType.onchange = function(){
    try{sessionStorage.setItem(sKey,JSON.stringify({t:selType.value,v:''}));}catch(e){}
    rebuild();
  };
  selVal.onchange  = function(){
    try{sessionStorage.setItem(sKey,JSON.stringify({t:selType.value,v:selVal.value}));}catch(e){}
    onChange();
  };
  // Restore saved type
  try {
    const sv=JSON.parse(sessionStorage.getItem(sKey)||'null');
    if(sv&&sv.t&&['month','quarter','year'].includes(sv.t)) selType.value=sv.t;
  } catch(e){}
  rebuild();
}

function aggregatePlRows(keys) {
  if (!keys || !keys.length) return [];
  if (keys.length === 1) return (D.pl_rows[keys[0]] || []).map(r=>({...r}));
  const order = [], seen = new Set();
  keys.forEach(k => (D.pl_rows[k]||[]).forEach(r => {
    if (!seen.has(r.b)) { seen.add(r.b); order.push({...r}); }
  }));
  const sum = {};
  keys.forEach(k => (D.pl_rows[k]||[]).forEach(r => { sum[r.b]=(sum[r.b]||0)+r.value; }));
  const pctRatio = {'19':['18','1'],'5':['4','1'],'9':['8','1'],'13':['12','1'],'22':['21','1']};
  return order.map(r => {
    let val = sum[r.b]||0;
    if (pctRatio[r.b]) {
      const [n,d] = pctRatio[r.b];
      val = sum[d] ? (sum[n]||0)/sum[d] : 0;
    } else if (Math.abs(r.value)<=1.5 && r.value!==0) {
      val = val / keys.length;
    }
    return {...r, value: val};
  }).filter(r => r.value !== 0);
}

function sortBCode(a, b) {
  const pa = a.split('.').map(Number);
  const pb = b.split('.').map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const na = i < pa.length ? pa[i] : 0;
    const nb = i < pb.length ? pb[i] : 0;
    if (na !== nb) return na - nb;
  }
  return 0;
}

function aggMonthly(keys) {
  const ms = D.monthly.filter(m => keys.includes(m.key));
  if (!ms.length) return null;
  const rev   = ms.reduce((s,m)=>s+m.revenue,0);
  const costs = ms.reduce((s,m)=>s+m.project_costs,0);
  const zp    = ms.reduce((s,m)=>s+m.zp_massive,0);
  const fin   = ms.reduce((s,m)=>s+m.fin_res_biz,0);
  return { revenue:rev, project_costs:costs, zp_massive:zp, fin_res_biz:fin,
           margin_pct: rev>0?fin/rev*100:0, zp_rev_pct: rev>0?zp/rev*100:0 };
}

function prevPeriod(type, val) {
  if (type === 'month') {
    const idx = MONTH_KEYS.indexOf(val);
    return idx > 0 ? MONTH_KEYS[idx-1] : null;
  }
  if (type === 'quarter') {
    const [q,yr] = val.split('-'); const qn=+q[1];
    if (qn===1) return `Q4-${+yr-1}`; return `Q${qn-1}-${yr}`;
  }
  if (type === 'year') return String(+val-1);
  return null;
}

function marginChip(m){
  if(m===null||m===undefined) return '<span class="chip chip-b">—</span>';
  return `<span class="chip ${chipCls(m)}">${fmtPct(m)}</span>`;
}

// ── Charts registry ───────────────────────────────────────────────────────────
const charts = {};
const inited = {};

function destroyChart(id){ if(charts[id]){ charts[id].destroy(); delete charts[id]; } }

// ── Tab switching ─────────────────────────────────────────────────────────────
const initFns = {};
function showTab(name, btn){
  document.querySelectorAll('.tab-content').forEach(el=>el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el=>el.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  btn.classList.add('active');
  if(!inited[name] && initFns[name]){ initFns[name](); inited[name]=true; }
}

// ── Page navigation ───────────────────────────────────────────────────────────
const PAGE_NAMES = {
  cf:   'CF &#1054;&#1090;&#1095;&#1105;&#1090;',
  rent: '&#1054;&#1090;&#1095;&#1105;&#1090; &#1087;&#1086; &#1088;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1080;',
  pl:   'P&amp;L &#1054;&#1090;&#1095;&#1105;&#1090;',
};

function showPage(name) {
  ['home','cf','rent','pl'].forEach(p => {
    const el = document.getElementById('page-'+p);
    if (el) el.style.display = (p === name) ? '' : 'none';
  });

  const isHome = name === 'home';
  document.getElementById('hdr-home-state').style.display   = isHome ? '' : 'none';
  document.getElementById('hdr-report-state').style.display = isHome ? 'none' : 'flex';
  document.getElementById('main-tabs').style.display        = name === 'rent' ? '' : 'none';
  document.getElementById('period-badge').style.display     = name === 'rent' ? '' : 'none';

  if (!isHome) {
    document.getElementById('hdr-report-name').innerHTML = PAGE_NAMES[name] || name;
  }

  if (name === 'rent' && !inited.main) {
    initFns.main(); inited.main = true;
  }

  document.documentElement.scrollTop = 0;
}

// ── Theme ─────────────────────────────────────────────────────────────────────
function toggleTheme(){
  const th = document.documentElement.dataset.theme;
  document.documentElement.dataset.theme = th==='dark'?'light':'dark';
  document.getElementById('theme-btn').textContent = th==='dark'?'🌙 Тёмная':'☀ Светлая';
  setTimeout(()=>Object.values(charts).forEach(c=>c&&c.update()), 30);
}

// ── Chart defaults ────────────────────────────────────────────────────────────
function chartDefaults(){
  return {
    responsive:true, maintainAspectRatio:false,
    interaction:{mode:'index',intersect:false},
    plugins:{
      legend:{labels:{color:()=>css('--text'),font:{size:11},boxWidth:11,padding:14}},
      tooltip:{bodyFont:{size:12}}
    }
  };
}
function scaleX(){ return {ticks:{color:()=>css('--muted'),font:{size:11}},grid:{color:()=>css('--grid')}}; }
function scaleY(label){ return {ticks:{color:()=>css('--muted'),font:{size:11}},grid:{color:()=>css('--grid')},title:{display:!!label,text:label||'',color:()=>css('--muted')}}; }

// ══════════════════════════════════════════════════════════════════════════════
// TAB: ГЛАВНАЯ
// ══════════════════════════════════════════════════════════════════════════════
function initMain(){
  // Badge + footer (one-time)
  const firstLabel = D.monthly[0].label;
  const lastLabel  = D.monthly[D.monthly.length-1].label;
  document.getElementById('period-badge').textContent = `Данные: ${firstLabel} — ${lastLabel}`;
  document.getElementById('footer-main').textContent =
    `STH-group Financial Dashboard  |  Данные: ${firstLabel} — ${lastLabel}`;
  function renderMain() {
    const type = document.getElementById('mainPeriodType').value;
    const val  = document.getElementById('mainPeriodVal').value;
    if (!val) return;
    const keys = getMonthKeys(type, val);
    const cur  = aggMonthly(keys);
    if (!cur) return;
    const pv   = prevPeriod(type, val);
    const prevKeys = pv ? getMonthKeys(type, pv) : null;
    const prv  = prevKeys ? aggMonthly(prevKeys) : null;

    const ytdByYear = {};
    D.monthly.forEach(m => {
      const yr = m.key.slice(3);
      ytdByYear[yr] = (ytdByYear[yr]||0) + m.revenue;
    });
    const years = Object.keys(ytdByYear).sort();
    const lastYr = years[years.length-1];
    const prevYr = years[years.length-2];

    const kpis = [
      {label:'Выручка', val:'₽'+fmtM(cur.revenue),
       delta: prv ? sign(cur.revenue-prv.revenue)+'₽'+fmtM(cur.revenue-prv.revenue) : '—',
       cls: prv?(cur.revenue>=prv.revenue?'up':'down'):'neutral',
       sub: prv?`Пред: ₽${fmtM(prv.revenue)}`:'', accent:'var(--accent)'},
      {label:'Рентабельность', val:fmtPct(cur.margin_pct),
       delta: prv?sign(cur.margin_pct-prv.margin_pct)+fmtPct(cur.margin_pct-prv.margin_pct):'—',
       cls: prv?(cur.margin_pct>=prv.margin_pct?'up':'down'):'neutral',
       sub: prv?`Пред: ${fmtPct(prv.margin_pct)}`:'', accent:'var(--green)'},
      {label:'ЗП / Выручка', val:fmtPct(cur.zp_rev_pct),
       delta: prv?sign(cur.zp_rev_pct-prv.zp_rev_pct)+fmtPct(cur.zp_rev_pct-prv.zp_rev_pct):'—',
       cls: prv?(cur.zp_rev_pct<=prv.zp_rev_pct?'up':'down'):'neutral',
       sub:`ЗП масс: ₽${fmtM(cur.zp_massive)}`, accent:'var(--amber)'},
      {label:`YTD ${lastYr}`, val:'₽'+fmtM(ytdByYear[lastYr]),
       delta:`${prevYr}: ₽${fmtM(ytdByYear[prevYr])}`,
       cls:'neutral',
       sub:`Ср/мес: ₽${fmtM(Math.round(ytdByYear[lastYr]/D.monthly.filter(m=>m.key.includes(lastYr)).length))}`,
       accent:'var(--purple)'},
    ];
    document.getElementById('kpi-main').innerHTML = kpis.map(k=>`
      <div class="kpi" style="border-left-color:${k.accent}">
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value">${k.val}</div>
        <div class="kpi-delta ${k.cls}">${k.delta}</div>
        <div class="kpi-sub">${k.sub}</div>
      </div>`).join('');

    const best  = D.monthly.reduce((a,b)=>a.margin_pct>b.margin_pct?a:b);
    const worst = D.monthly.reduce((a,b)=>a.margin_pct<b.margin_pct?a:b);
    const zpTrend = D.monthly[D.monthly.length-1].zp_rev_pct - D.monthly[0].zp_rev_pct;
    document.getElementById('alerts-main').innerHTML = `
      <div class="alert-box alert-g">
        <div class="alert-title">🏆 Лучший месяц по рент-ти</div>
        <div class="alert-val">${best.label} — ${fmtPct(best.margin_pct)}</div>
        <div class="alert-sub">Выручка: ₽${fmtM(best.revenue)}</div>
      </div>
      <div class="alert-box alert-r">
        <div class="alert-title">⚠️ Худший месяц по рент-ти</div>
        <div class="alert-val">${worst.label} — ${fmtPct(worst.margin_pct)}</div>
        <div class="alert-sub">Выручка: ₽${fmtM(worst.revenue)}</div>
      </div>
      <div class="alert-box ${zpTrend<0?'alert-g':'alert-r'}">
        <div class="alert-title">📉 ЗП/Выручка: тренд за всё время</div>
        <div class="alert-val">${sign(zpTrend)}${fmtPct(zpTrend)}</div>
        <div class="alert-sub">${D.monthly[0].label}: ${fmtPct(D.monthly[0].zp_rev_pct)} → ${lastLabel}: ${fmtPct(D.monthly[D.monthly.length-1].zp_rev_pct)}</div>
      </div>`;
  }

  // Sync threshold inputs with stored values
  const tgEl=document.getElementById('threshGood'), twEl=document.getElementById('threshWarn');
  if(tgEl) tgEl.value=THRESHOLDS.good;
  if(twEl) twEl.value=THRESHOLDS.warn;

  const selType = document.getElementById('mainPeriodType');
  const selVal  = document.getElementById('mainPeriodVal');
  fillPeriodSelect(selType, selVal, renderMain);

  // Mini chart (last 6 months — always static)
  const mini = D.monthly.slice(-6);
  const ctx = document.getElementById('miniChart').getContext('2d');
  charts.mini = new Chart(ctx, {
    data:{
      labels: mini.map(m=>m.label),
      datasets:[
        {type:'bar',label:'Выручка, тыс.₽',data:mini.map(m=>Math.round(m.revenue/1e3)),
         backgroundColor:'rgba(92,107,192,.6)',yAxisID:'y',order:2},
        {type:'line',label:'Рент-ть %',data:mini.map(m=>m.margin_pct),
         borderColor:'#66bb6a',backgroundColor:'rgba(102,187,106,.1)',fill:true,
         tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
      ]
    },
    options:{...chartDefaults(),scales:{
      x:scaleX(),
      y:{...scaleY('тыс. ₽'),position:'left'},
      y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false},min:0},
    }}
  });
}
initFns.main = initMain;

// ══════════════════════════════════════════════════════════════════════════════
// TAB: ТРЕНДЫ
// ══════════════════════════════════════════════════════════════════════════════
initFns.trends = function(){
  // ── Date range selects ────────────────────────────────────────────────────
  const selFrom = document.getElementById('trendsFrom');
  const selTo   = document.getElementById('trendsTo');
  D.monthly.forEach(m => {
    selFrom.innerHTML += `<option value="${m.key}">${m.label}</option>`;
    selTo.innerHTML   += `<option value="${m.key}">${m.label}</option>`;
  });
  selTo.selectedIndex = D.monthly.length - 1;

  function getSlice() {
    const fi = MONTH_KEYS.indexOf(selFrom.value);
    const ti = MONTH_KEYS.indexOf(selTo.value);
    const lo = Math.min(fi,ti), hi = Math.max(fi,ti);
    return D.monthly.filter((_,i)=>i>=lo&&i<=hi);
  }

  // ── Standard charts ───────────────────────────────────────────────────────
  function buildStandardCharts() {
    const slice = getSlice();
    const labels=slice.map(m=>m.label);
    const rev   =slice.map(m=>Math.round(m.revenue/1e3));
    const costs =slice.map(m=>Math.round(m.project_costs/1e3));
    const margin=slice.map(m=>m.margin_pct);
    const zpRev =slice.map(m=>m.zp_rev_pct);
    const finBiz=slice.map(m=>Math.round(m.fin_res_biz/1e3));

    destroyChart('trends'); destroyChart('zp');
    charts.trends=new Chart(document.getElementById('trendsChart').getContext('2d'),{
      data:{labels,datasets:[
        {type:'bar',label:'Выручка, тыс.₽',data:rev,backgroundColor:'rgba(92,107,192,.65)',yAxisID:'y',order:2},
        {type:'bar',label:'Расходы, тыс.₽',data:costs,backgroundColor:'rgba(71,85,105,.55)',yAxisID:'y',order:2},
        {type:'line',label:'Рент-ть %',data:margin,borderColor:'#66bb6a',backgroundColor:'rgba(102,187,106,.08)',
         fill:true,tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
        {type:'line',label:'ЗП/Выр %',data:zpRev,borderColor:'#ffa726',
         tension:.3,borderDash:[5,3],yAxisID:'y2',order:1,pointRadius:3,borderWidth:2},
      ]},
      options:{...chartDefaults(),scales:{
        x:scaleX(),
        y:{...scaleY('тыс. ₽'),position:'left'},
        y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false},min:0,
            max:Math.ceil(Math.max(...zpRev,100)/10)*10},
      }}
    });
    charts.zp=new Chart(document.getElementById('zpChart').getContext('2d'),{
      data:{labels,datasets:[
        {type:'bar',label:'Фин. рез. бизнеса, тыс.₽',data:finBiz,backgroundColor:'rgba(38,198,218,.5)',yAxisID:'y',order:2},
        {type:'line',label:'ЗП масс. / Выр %',data:zpRev,borderColor:'#ffa726',
         backgroundColor:'rgba(255,167,38,.1)',fill:true,tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
      ]},
      options:{...chartDefaults(),scales:{
        x:scaleX(),
        y:{...scaleY('тыс. ₽'),position:'left'},
        y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false},min:0},
      }}
    });
  }

  // ── Custom metric chart ───────────────────────────────────────────────────
  const metricSel = document.getElementById('customMetricSel');
  const metricTypeSel = document.getElementById('customMetricChart2Type');

  // Populate from all B-codes present in pl_rows + PL_DESC
  const allBCodes = new Set();
  D.monthly.forEach(m=>(D.pl_rows[m.key]||[]).forEach(r=>allBCodes.add(r.b)));
  [...allBCodes].sort(sortBCode).forEach(b=>{
    const lbl = PL_DESC[b] ? `${b} — ${PL_DESC[b].slice(0,50)}` : `Код ${b}`;
    metricSel.innerHTML += `<option value="${b}">${lbl}</option>`;
  });

  function buildCustomMetric() {
    const b = metricSel.value;
    const chartType = metricTypeSel.value;
    if (!b) return;
    const slice = getSlice();
    const data = slice.map(m=>{
      const row=(D.pl_rows[m.key]||[]).find(r=>r.b===b);
      return row ? row.value : null;
    });
    const isPct = data.some(v=>v!==null&&Math.abs(v)<=1.5&&v!==0);
    const display = data.map(v=>v!==null?(isPct?+(v*100).toFixed(2):v):null);
    const lbl = PL_DESC[b]||('Код '+b);

    destroyChart('customMetric');
    charts.customMetric=new Chart(document.getElementById('customMetricChart').getContext('2d'),{
      data:{labels:slice.map(m=>m.label),datasets:[{
        type:chartType,label:lbl.slice(0,60),data:display,
        backgroundColor:'rgba(92,107,192,.55)',
        borderColor:'#5c6bc0',fill:chartType==='line',
        tension:.3,pointRadius:4,borderWidth:2,spanGaps:true
      }]},
      options:{...chartDefaults(),plugins:{legend:{labels:{color:()=>css('--text'),font:{size:11}}}},
        scales:{x:scaleX(),y:{...scaleY(isPct?'%':'₽'),ticks:{
          color:()=>css('--muted'),
          callback:v=>isPct?v.toFixed(1)+'%':fmtM(v)
        }}}}
    });
  }
  metricSel.onchange = buildCustomMetric;
  metricTypeSel.onchange = buildCustomMetric;
  if(metricSel.options.length) buildCustomMetric();

  // ── Project / division trend ──────────────────────────────────────────────
  const filterMode = document.getElementById('trendsFilterMode');
  const filterDetail = document.getElementById('trendsFilterDetail');

  const allProjNames = [];
  {const s=new Set(); D.monthly.forEach(m=>(D.projects_per_month[m.key]||[]).forEach(p=>s.add(p.name)));
   [...s].sort().forEach(n=>allProjNames.push(n));}

  function renderFilterDetail() {
    filterDetail.innerHTML = '';
    if(filterMode.value==='projects'){
      filterDetail.innerHTML=`
        <div style="display:flex;flex-direction:column;gap:4px">
          <input id="trendsSearchProj" placeholder="&#1060;&#1080;&#1083;&#1100;&#1090;&#1088;..." oninput="filterTrendsProjs()"
            style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:5px 8px;border-radius:5px;font-size:12px;outline:none">
          <select id="trendsProjSel" multiple size="6" style="min-width:280px;background:var(--card2);border:1px solid var(--border);color:var(--text);border-radius:5px;font-size:12px"></select>
        </div>`;
      fillTrendsProjSel(allProjNames);
      document.getElementById('trendsProjSel').onchange = buildProjTrend;
    } else if(filterMode.value==='division'){
      const rrpNames=[...new Set(Object.values(D.projects_per_month).flat().map(p=>p.rrp).filter(Boolean))].sort();
      filterDetail.innerHTML=`<select id="trendsDivSel" style="min-width:200px"><option value="all">&#1042;&#1089;&#1077; &#1056;&#1056;&#1055;</option>`+
        rrpNames.map(d=>`<option value="${d}">${d}</option>`).join('')+'</select>';
      document.getElementById('trendsDivSel').onchange = buildProjTrend;
    }
    buildProjTrend();
  }

  function fillTrendsProjSel(names) {
    const sel=document.getElementById('trendsProjSel');
    if(!sel) return;
    const prev=new Set([...sel.selectedOptions].map(o=>o.value));
    sel.innerHTML=''; names.forEach(n=>{
      sel.innerHTML+=`<option value="${n}" ${prev.has(n)?'selected':''}>${n}</option>`;
    });
  }

  window.filterTrendsProjs = function(){
    const fl=document.getElementById('trendsSearchProj').value.toLowerCase();
    fillTrendsProjSel(allProjNames.filter(n=>!fl||n.toLowerCase().includes(fl)));
  };

  const PALETTE=['#5c6bc0','#66bb6a','#ffa726','#ef5350','#26c6da','#ab47bc','#8d6e63','#78909c'];

  function buildProjTrend() {
    const slice = getSlice();
    const mode = filterMode.value;
    destroyChart('projTrend2');

    if(mode==='all'){
      // aggregate all projects — show company revenue+margin
      const data=slice.map(m=>({
        rev:Math.round(m.revenue/1e3), mar:m.margin_pct
      }));
      charts.projTrend2=new Chart(document.getElementById('trendsProjChart').getContext('2d'),{
        data:{labels:slice.map(m=>m.label),datasets:[
          {type:'bar',label:'Выручка (все), тыс.₽',data:data.map(d=>d.rev),
           backgroundColor:'rgba(92,107,192,.6)',yAxisID:'y',order:2},
          {type:'line',label:'Рент-ть %',data:data.map(d=>d.mar),
           borderColor:'#66bb6a',tension:.3,fill:false,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
        ]},
        options:{...chartDefaults(),scales:{
          x:scaleX(),y:{...scaleY('тыс. ₽'),position:'left'},
          y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false}},
        }}
      });
      return;
    }

    if(mode==='division'){
      const sel=document.getElementById('trendsDivSel');
      const divSel=sel?sel.value:'all';
      const _cf=D.cities.length?buildCityExtractor(D.cities.map(c=>c.name)):null;
      const gc=n=>{if(_cf){const r=_cf(n);if(r)return r;}const w=n.split(/[\s_]+/);return w[w.length-1]||n;};
      const allRrp=[...new Set(Object.values(D.projects_per_month).flat().map(p=>p.rrp).filter(Boolean))].sort();
      const divsToShow=divSel==='all'?allRrp:[divSel];
      const PALETTE2=['rgba(92,107,192,.75)','rgba(239,83,80,.7)','rgba(38,198,218,.7)',
        'rgba(102,187,106,.7)','rgba(255,167,38,.7)','rgba(171,71,188,.7)',
        'rgba(38,166,154,.7)','rgba(120,130,150,.5)'];
      const datasets=[];
      divsToShow.forEach((div,i)=>{
        const revData=slice.map(m=>{
          const projs=(D.projects_per_month[m.key]||[]).filter(p=>p.rrp===div);
          return Math.round(projs.reduce((s,p)=>s+p.rev,0)/1e3);
        });
        if(revData.every(v=>v===0)) return;
        datasets.push({type:'bar',label:div,data:revData,
          backgroundColor:PALETTE2[i%PALETTE2.length],
          stack:'div',yAxisID:'y',order:2});
      });
      if(!datasets.length){document.getElementById('trendsProjChart').style.display='none';return;}
      document.getElementById('trendsProjChart').style.display='';
      charts.projTrend2=new Chart(document.getElementById('trendsProjChart').getContext('2d'),{
        data:{labels:slice.map(m=>m.label),datasets},
        options:{...chartDefaults(),scales:{
          x:{...scaleX(),stacked:true},
          y:{...scaleY('тыс. ₽'),position:'left',stacked:true},
        }}
      });
      return;
    }

    // projects mode
    const sel=document.getElementById('trendsProjSel');
    if(!sel) return;
    const selected=[...sel.selectedOptions].map(o=>o.value);
    if(!selected.length){
      destroyChart('projTrend2');
      document.getElementById('trendsProjChart').style.display='none'; return;
    }
    document.getElementById('trendsProjChart').style.display='';
    const datasets=[];
    selected.slice(0,8).forEach((name,i)=>{
      const revData=slice.map(m=>{
        const p=(D.projects_per_month[m.key]||[]).find(x=>x.name===name);
        return p?Math.round(p.rev/1e3):0;
      });
      const marData=slice.map(m=>{
        const p=(D.projects_per_month[m.key]||[]).find(x=>x.name===name);
        return p?p.margin:null;
      });
      const col=PALETTE[i%PALETTE.length];
      datasets.push(
        {type:'bar',label:name.slice(0,25)+' Выр.',data:revData,
         backgroundColor:col.replace(')',',0.55)').replace('rgb(','rgba('),yAxisID:'y',order:2,stack:'rev'},
        {type:'line',label:name.slice(0,25)+' Рент.',data:marData,
         borderColor:col,fill:false,tension:.3,yAxisID:'y2',order:1,
         pointRadius:3,borderWidth:2,spanGaps:true}
      );
    });
    charts.projTrend2=new Chart(document.getElementById('trendsProjChart').getContext('2d'),{
      data:{labels:slice.map(m=>m.label),datasets},
      options:{...chartDefaults(),scales:{
        x:{...scaleX(),stacked:true},
        y:{...scaleY('тыс. ₽'),position:'left',stacked:true},
        y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false}},
      }}
    });
  }

  filterMode.onchange = renderFilterDetail;
  renderFilterDetail();

  // ── Wire up date range ────────────────────────────────────────────────────
  function onRangeChange() {
    buildStandardCharts();
    buildCustomMetric();
    buildProjTrend();
  }
  selFrom.onchange = onRangeChange;
  selTo.onchange   = onRangeChange;

  buildStandardCharts();
};

// ══════════════════════════════════════════════════════════════════════════════
// TAB: ОБЩИЙ
// ══════════════════════════════════════════════════════════════════════════════

// Описания полей по коду из B-колонки
const PL_DESC = {
  '1':     'Выручка — общий доход от реализации услуг (с НДС)',
  '1.1':   'Выручка без скидок и возвратов',
  '1.2':   'Скидки, возвраты (вычитаются из выручки)',
  '2':     'Прямые расходы на персонал — итого по всем статьям',
  '2.1':   'Фонд оплаты труда производственного персонала',
  '2.1.1': 'ЗП массового персонала (рабочие, линейный персонал)',
  '2.1.2': 'ЗП бригадиров и премии',
  '2.3':   'Страховые взносы — отчисления в социальные фонды',
  '2.4':   'Полевые расходы — итого (проживание, питание, проезд и пр.)',
  '2.4.1': 'Проживание персонала на объекте',
  '2.4.2': 'Питание персонала',
  '2.4.3': 'Проезд (командировочные, билеты)',
  '2.4.4': 'Спецодежда и средства индивидуальной защиты',
  '2.4.5': 'Организация рабочей площадки (аренда, оборудование)',
  '2.4.6': 'Услуги по подбору персонала',
  '2.4.7': 'Оплата подрядчику (субподряд)',
  '2.5':   'Трудкнижки, медкнижки, ФМС — оформление документов',
  '2.6':   'Прочие прямые расходы',
  '2.7':   'Спецоборудование и оснастка',
  '2.9':   'Бонус за подбор персонала',
  '2.10':  'НДФЛ, НПД и другие налоги на доходы физлиц',
  '2.11':  'Коммерческое вознаграждение партнёрам',
  '2.12':  'Прочие расходы по деятельности',
  '2.13':  'Комиссия банка / факторинговое обслуживание',
  '2.14':  'Штрафы и санкции',
  '2.19':  'НДС к уплате',
  '2.21':  'Налог на прибыль',
  '3':     'Расходы на управленческий персонал (STH-group: управление)',
  '3.1':   'ЗП директора и управляющих (производственные)',
  '3.2':   'Прочие расходы управленческого персонала',
  '4':     'Финансовый результат уровня Супервайзера — после ЗП и прямых расходов',
  '5':     'Рентабельность по уровню Супервайзера (%)',
  '6':     'Расходы РРП — итого',
  '7':     'Прямые расходы уровня РРП (региональный руководитель производства)',
  '8':     'Финансовый результат уровня РРП',
  '9':     'Рентабельность РРП (%)',
  '10':    'Расходы ДРП — итого',
  '11':    'Прямые расходы уровня ДРП (директор по развитию производства)',
  '12':    'Финансовый результат уровня ДРП',
  '13':    'Рентабельность ДРП (%)',
  '15':    'Прочие расходы бизнеса',
  '16':    'Административные расходы',
  '17':    'Прочие операционные расходы',
  '18':    'Финансовый результат Бизнеса — итоговая прибыль до налогов',
  '19':    'РЕНТАБЕЛЬНОСТЬ — итоговый показатель эффективности бизнеса (%)',
  '20':    'Сумма налогов к уплате',
  '21':    'Финансовый результат после налогов',
  '22':    'Рентабельность после налогов (%)',
};

function renderGeneral(rows){
  if(!rows || !rows.length){
    document.getElementById('generalTable').innerHTML = '<tbody><tr><td colspan="5" style="color:var(--muted);padding:20px">Нет данных для этого периода</td></tr></tbody>';
    return;
  }
  const revRow = rows.find(r => r.b === '1');
  const rev = revRow ? revRow.value : 0;

  const parentSet = new Set();
  rows.forEach(r => {
    const parts = r.b.split('.');
    for (let i = 1; i < parts.length; i++) parentSet.add(parts.slice(0,i).join('.'));
  });

  let html = `<thead><tr>
    <th class="pl-code">Код</th>
    <th>Показатель
      <button class="btn-sm" style="margin-left:8px;font-size:10px;padding:2px 7px" onclick="toggleAllGroups('generalTable',false)">&#9658; Свернуть</button>
      <button class="btn-sm" style="font-size:10px;padding:2px 7px" onclick="toggleAllGroups('generalTable',true)">&#9660; Развернуть</button>
    </th>
    <th class="td-r">Значение</th>
    <th class="pl-share">Доля выр.<span class="alloc-note" data-tip="&#1056;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100; — &#1076;&#1086; &#1072;&#1083;&#1083;&#1086;&#1082;&#1072;&#1094;&#1080;&#1080;">&#174;&#1072;</span></th>
  </tr></thead><tbody>`;

  rows.forEach(r => {
    const isZero = r.value === 0 || r.value === 0.0;
    const indent = r.indent * 18 + 4;
    const isTopLevel = r.indent === 0;
    const isTotal = isTopLevel && /^\d+$/.test(r.b);
    const isParent = parentSet.has(r.b);
    const isDecPct = !isZero && Math.abs(r.value) <= 1.5;

    let valStr, shareStr = '—';
    if(isDecPct){
      const pv = r.value * 100;
      const cls = pv >= 15 ? 'pos' : pv >= 5 ? '' : pv < 0 ? 'neg' : '';
      valStr = `<span class="${cls}" style="font-weight:600">${fmtPct(pv)}</span>`;
    } else {
      const cls = r.value < 0 ? 'neg' : '';
      valStr = `<span class="${cls}">${fmtRub(r.value)}</span>`;
      if(rev > 0 && r.b !== '1') shareStr = fmtPct(Math.abs(r.value) / rev * 100);
    }

    const tip = PL_DESC[r.b] || ('Код ' + r.b);
    const rowCls = [
      isZero ? 'pl-zero' : '',
      isTotal ? 'pl-top0-total' : isTopLevel ? 'pl-top0' : '',
    ].filter(Boolean).join(' ');
    const togBtn = isParent
      ? `<button class="tog" data-expanded="true" data-parentb="${r.b}" onclick="toggleGroup(this,'${r.b}')">&#9660;</button>`
      : '<span style="display:inline-block;width:14px"></span>';

    html += `<tr class="${rowCls}" data-b="${r.b}">
      <td class="pl-code">${r.b}</td>
      <td style="padding-left:${indent}px">${togBtn}<span class="pl-label" data-tip="${tip}">${r.label}</span></td>
      <td class="pl-val">${valStr}</td>
      <td class="pl-share">${shareStr}</td>
    </tr>`;
  });

  html += '</tbody>';
  document.getElementById('generalTable').innerHTML = html;
}

function renderWaterfall(rows){
  const get=b=>(rows.find(r=>r.b===b)||{value:0}).value;
  const sc=v=>Math.round(v/1e3);
  const r1=get('1'),r2=get('2'),r4=get('4'),r6=get('6'),r8=get('8'),
        r10=get('10'),r12=get('12'),r15=get('15')+get('16')+get('17'),r18=get('18');
  const items=[
    {l:'Выручка',          bar:[0,sc(r1)],         col:'rgba(102,187,106,.75)'},
    {l:'Прямые расходы',   bar:[sc(r4),sc(r1)],    col:'rgba(239,83,80,.65)'},
    {l:'Рез. Супервайзер', bar:[0,sc(r4)],          col:'rgba(92,107,192,.75)'},
    {l:'Расходы РРП',      bar:[sc(r8),sc(r4)],    col:'rgba(239,83,80,.65)'},
    {l:'Рез. РРП',         bar:[0,sc(r8)],          col:'rgba(92,107,192,.75)'},
    {l:'Расходы ДРП',      bar:[sc(r12),sc(r8)],   col:'rgba(239,83,80,.65)'},
    {l:'Рез. ДРП',         bar:[0,sc(r12)],         col:'rgba(92,107,192,.75)'},
    {l:'Прочие расходы',   bar:[sc(r18),sc(r12)],  col:'rgba(239,83,80,.65)'},
    {l:'Фин. рез.',        bar:[0,sc(r18)],         col:r18>=0?'rgba(102,187,106,.75)':'rgba(239,83,80,.65)'},
  ].filter(x=>x.bar[0]!==x.bar[1]&&!isNaN(x.bar[0])&&!isNaN(x.bar[1]));
  destroyChart('waterfall');
  if(!items.length||!r1) return;
  charts.waterfall=new Chart(document.getElementById('waterfallChart').getContext('2d'),{
    type:'bar',
    data:{labels:items.map(i=>i.l),datasets:[{
      data:items.map(i=>i.bar),backgroundColor:items.map(i=>i.col),
      borderRadius:3,borderSkipped:false}]},
    options:{...chartDefaults(),
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:ctx=>{const v=ctx.raw;const d=(v[1]-v[0])*1e3;return `${sign(d)}₽${fmtM(d)}`;}}}},
      scales:{
        x:scaleX(),
        y:{...scaleY('тыс. ₽'),ticks:{color:()=>css('--muted'),callback:v=>fmtM(v*1e3)}}}
    }
  });
}

initFns.general = function(){
  const selType = document.getElementById('generalPeriodType');
  const selVal  = document.getElementById('generalPeriodVal');
  function upd() {
    const keys = getMonthKeys(selType.value, selVal.value);
    const rows = aggregatePlRows(keys);
    renderGeneral(rows);
    renderWaterfall(rows);
  }
  fillPeriodSelect(selType, selVal, upd);
};

// ══════════════════════════════════════════════════════════════════════════════
// TAB: СРАВНЕНИЕ
// ══════════════════════════════════════════════════════════════════════════════
function renderCompare(keysA, keysB, labelA, labelB) {
  const rowsA = aggregatePlRows(keysA);
  const rowsB = aggregatePlRows(keysB);
  if (!rowsA.length && !rowsB.length) {
    document.getElementById('compareTable').innerHTML = '<div style="color:var(--muted);padding:20px">Нет данных</div>'; return;
  }
  const mapA = Object.fromEntries(rowsA.map(r=>[r.b,r]));
  const mapB = Object.fromEntries(rowsB.map(r=>[r.b,r]));
  const order = [], seen = new Set();
  [...rowsA,...rowsB].forEach(r => { if(!seen.has(r.b)){seen.add(r.b);order.push(r);} });
  order.sort((a,b) => sortBCode(a.b, b.b));

  // Warn if pl_rows B='1' differs >50% from monthly aggregate
  function revWarn(keys, label) {
    const plRev = aggregatePlRows(keys).find(r=>r.b==='1')?.value||0;
    const mRev  = aggMonthly(keys)?.revenue||0;
    if(mRev>0 && plRev>0 && Math.abs(plRev-mRev)/mRev > 0.5)
      return `<span style="color:var(--amber);cursor:help;font-size:11px" data-tip="В исходном Excel значение выручки в строке B=1 (${fmtRub(plRev)}) отличается от сводного показателя (${fmtRub(mRev)}). Возможно, данные за этот период неполные."> ⚠</span>`;
    return '';
  }
  const warnA = revWarn(keysA, labelA), warnB = revWarn(keysB, labelB);

  let html = `<table><thead><tr>
    <th class="pl-code">&#1050;&#1086;&#1076;</th>
    <th>&#1055;&#1086;&#1082;&#1072;&#1079;&#1072;&#1090;&#1077;&#1083;&#1100;</th>
    <th class="td-r">${labelA}${warnA}</th>
    <th class="td-r">${labelB}${warnB}</th>
    <th class="td-r">&#916;</th>
    <th class="td-r">&#916;%<span class="alloc-note" data-tip="&#1056;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100; — &#1076;&#1086; &#1072;&#1083;&#1083;&#1086;&#1082;&#1072;&#1094;&#1080;&#1080;">&#174;&#1072;</span></th>
  </tr></thead><tbody>`;

  order.forEach(r => {
    const rA = mapA[r.b], rB = mapB[r.b];
    if (!rA && !rB) return;
    const refVal = rA ? rA.value : rB.value;
    const isDecPct = Math.abs(refVal) <= 1.5 && refVal !== 0;
    const fv = isDecPct ? v => fmtPct(v*100) : fmtRub;
    const indent = r.indent * 18 + 4;
    const tip = PL_DESC[r.b] || ('Код ' + r.b);
    const rowCls = r.indent===0 ? (/^\d+$/.test(r.b)?'pl-top0-total':'pl-top0') : '';

    const va = rA ? rA.value : 0;
    const vb = rB ? rB.value : 0;
    const d  = vb - va;
    const dp = va !== 0 ? d/Math.abs(va)*100 : (vb !== 0 ? 100 : 0);
    const cls = d >= 0 ? 'pos' : 'neg';

    const vAstr = rA ? fv(va) : `<span style="color:var(--muted)" title="нет в периоде А">0</span>`;
    const vBstr = rB ? fv(vb) : `<span style="color:var(--muted)" title="нет в периоде Б">0</span>`;
    const dStr  = `<span class="${cls}">${sign(d)}${fv(d)}</span>`;
    const dpStr = `<span class="${cls}">${sign(dp)}${dp.toFixed(1)}%</span>`;

    html += `<tr class="${rowCls}" data-b="${r.b}">
      <td class="pl-code">${r.b}</td>
      <td style="padding-left:${indent}px"><span class="pl-label" data-tip="${tip}">${r.label}</span></td>
      <td class="td-r">${vAstr}</td>
      <td class="td-r">${vBstr}</td>
      <td class="td-r">${dStr}</td>
      <td class="td-r">${dpStr}</td>
    </tr>`;
  });

  html += '</tbody></table>';
  document.getElementById('compareTable').innerHTML = html;
}

let _pcsc='revB', _pcsa=false;

function renderProjectCompare(keysA, keysB, labelA, labelB, filter) {
  const mapA = {}, mapB = {};
  keysA.forEach(k => (D.projects_per_month[k]||[]).forEach(p => {
    if (!mapA[p.name]) mapA[p.name] = {rev:0, finSum:0};
    mapA[p.name].rev += p.rev; mapA[p.name].finSum += p.rev*(p.margin/100);
  }));
  keysB.forEach(k => (D.projects_per_month[k]||[]).forEach(p => {
    if (!mapB[p.name]) mapB[p.name] = {rev:0, finSum:0};
    mapB[p.name].rev += p.rev; mapB[p.name].finSum += p.rev*(p.margin/100);
  }));
  const allNames = [...new Set([...Object.keys(mapA),...Object.keys(mapB)])].sort();
  const fl = filter.toLowerCase().trim();
  const rows = allNames
    .filter(n => !fl || n.toLowerCase().includes(fl))
    .map(n => {
      const a = mapA[n], b = mapB[n];
      const revA = a?a.rev:0, revB = b?b.rev:0;
      const marA = a&&a.rev>0?a.finSum/a.rev*100:0;
      const marB = b&&b.rev>0?b.finSum/b.rev*100:0;
      return {n, revA, revB, marA, marB, dr:revB-revA, dm:marB-marA};
    })
    .sort((a,b)=>{
      const av=a[_pcsc]??0, bv=b[_pcsc]??0;
      const cmp=typeof av==='string'?av.localeCompare(bv,'ru'):(av-bv);
      return _pcsa?cmp:-cmp;
    });
  function sortPC(col){if(_pcsc===col){_pcsa=!_pcsa;}else{_pcsc=col;_pcsa=false;}renderProjectCompare(keysA,keysB,labelA,labelB,filter);}
  window._sortPC=sortPC;
  if (!rows.length) {
    document.getElementById('projectCompare').innerHTML =
      '<div style="color:var(--muted);padding:16px;font-size:12px">Нет проектов</div>'; return;
  }
  const sa=c=>c===_pcsc?(_pcsa?' ↑':' ↓'):'';
  let html = `<table><thead><tr>
    <th style="cursor:pointer" onclick="_sortPC('n')">&#1055;&#1088;&#1086;&#1077;&#1082;&#1090;${sa('n')}</th>
    <th class="td-r" style="cursor:pointer" onclick="_sortPC('revA')">&#1042;&#1099;&#1088;. ${labelA}${sa('revA')}</th>
    <th class="td-r" style="cursor:pointer" onclick="_sortPC('revB')">&#1042;&#1099;&#1088;. ${labelB}${sa('revB')}</th>
    <th class="td-r" style="cursor:pointer" onclick="_sortPC('dr')">&#916; &#1074;&#1099;&#1088;.${sa('dr')}</th>
    <th class="td-r" style="cursor:pointer" onclick="_sortPC('marA')">&#1056;&#1077;&#1085;&#1090;. ${labelA}${sa('marA')}</th>
    <th class="td-r" style="cursor:pointer" onclick="_sortPC('marB')">&#1056;&#1077;&#1085;&#1090;. ${labelB}${sa('marB')}</th>
    <th class="td-r" style="cursor:pointer" onclick="_sortPC('dm')">&#916; &#1088;&#1077;&#1085;&#1090;.${sa('dm')}</th>
  </tr></thead><tbody>`;
  rows.forEach(r => {
    const dr = r.revB-r.revA, dm = r.marB-r.marA;
    const chipA = chipCls(r.marA);
    const chipB = chipCls(r.marB);
    const noA = !mapA[r.n], noB = !mapB[r.n];
    html += `<tr>
      <td style="font-size:12px;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.n}">${r.n}</td>
      <td class="td-r ${noA?'':''}"><span ${noA?'style="color:var(--muted)"':''}>${noA?'0':fmtRub(r.revA)}</span></td>
      <td class="td-r"><span ${noB?'style="color:var(--muted)"':''}>${noB?'0':fmtRub(r.revB)}</span></td>
      <td class="td-r"><span class="${dr>=0?'pos':'neg'}">${sign(dr)}${fmtRub(dr)}</span></td>
      <td class="td-r">${r.marA>0?`<span class="chip ${chipA}">${fmtPct(r.marA)}</span>`:'<span style="color:var(--muted)">—</span>'}</td>
      <td class="td-r">${r.marB>0?`<span class="chip ${chipB}">${fmtPct(r.marB)}</span>`:'<span style="color:var(--muted)">—</span>'}</td>
      <td class="td-r"><span class="${dm>=0?'pos':'neg'}">${sign(dm)}${fmtPct(dm)}</span></td>
    </tr>`;
  });
  html += `</tbody></table>`;
  document.getElementById('projectCompare').innerHTML = html;
}

initFns.compare = function(){
  const tA=document.getElementById('cmpTypeA'), vA=document.getElementById('cmpValA');
  const tB=document.getElementById('cmpTypeB'), vB=document.getElementById('cmpValB');
  const srch=document.getElementById('projectSearch');
  function upd() {
    const keysA=getMonthKeys(tA.value,vA.value), keysB=getMonthKeys(tB.value,vB.value);
    const lA=getPeriodLabel(tA.value,vA.value), lB=getPeriodLabel(tB.value,vB.value);
    renderCompare(keysA,keysB,lA,lB);
    renderProjectCompare(keysA,keysB,lA,lB,srch.value);
  }
  srch.oninput = upd;
  fillPeriodSelect(tA, vA, upd);
  fillPeriodSelect(tB, vB, upd);
  if (vB.options.length > 1) { vB.selectedIndex = 1; upd(); }
};

// ══════════════════════════════════════════════════════════════════════════════
// TAB: АНОМАЛИИ
// ══════════════════════════════════════════════════════════════════════════════
function anomSmallTable(projects, col, top, n){
  if(!projects.length) return '<div style="color:var(--muted);padding:8px;font-size:11px">Нет данных</div>';
  let html='<table style="font-size:11px"><thead><tr><th>Проект</th>';
  if(col==='rev')   html+='<th class="td-r">Выручка</th>';
  if(col==='prof')  html+='<th class="td-r">Прибыль</th><th class="td-r">Рент.</th>';
  if(col==='mar')   html+='<th class="td-r">Рент-ть</th>';
  if(col==='drev')  html+='<th class="td-r">&#916; &#8381;</th><th class="td-r">&#916; %</th>';
  if(col==='dprof') html+='<th class="td-r">&#916; &#8381;</th><th class="td-r">&#916; %</th>';
  if(col==='dmar')  html+='<th class="td-r">&#916; п.п.</th>';
  html+='</tr></thead><tbody>';
  const rows = top ? projects.slice(0,n) : [...projects].reverse().slice(0,n);
  rows.forEach(p=>{
    html+=`<tr><td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${p.name}"><a class="proj-link" onclick="openProjectCard('${p.name.replace(/'/g,"\\'")}')">${p.name}</a></td>`;
    if(col==='rev')   html+=`<td class="td-r">&#8381;${fmtM(p.rev)}</td>`;
    if(col==='prof'){
      html+=`<td class="td-r">&#8381;${fmtM(p.fin)}</td><td class="td-r"><span class="chip ${chipCls(p.margin)}">${fmtPct(p.margin)}</span></td>`;
    }
    if(col==='mar'){
      const chip=chipCls(p.margin);
      html+=`<td class="td-r"><span class="chip ${chip}">${fmtPct(p.margin)}</span></td>`;
    }
    if(col==='drev'){
      const cls=p.drev>=0?'pos':'neg';
      const pct=p.prevRev>0?p.drev/p.prevRev*100:0;
      html+=`<td class="td-r"><span class="${cls}">${sign(p.drev)}&#8381;${fmtM(p.drev)}</span></td>`;
      html+=`<td class="td-r"><span class="${cls}">${sign(pct)}${pct.toFixed(1)}%</span></td>`;
    }
    if(col==='dprof'){
      const cls=p.dprof>=0?'pos':'neg';
      const pct=Math.abs(p.prevFin)>1000?p.dprof/Math.abs(p.prevFin)*100:0;
      html+=`<td class="td-r"><span class="${cls}">${sign(p.dprof)}&#8381;${fmtM(p.dprof)}</span></td>`;
      html+=`<td class="td-r"><span class="${cls}">${sign(pct)}${pct.toFixed(1)}%</span></td>`;
    }
    if(col==='dmar'){
      const cls=p.dmar>=0?'pos':'neg';
      html+=`<td class="td-r"><span class="${cls}">${sign(p.dmar)}${fmtPct(p.dmar)}</span></td>`;
    }
    html+='</tr>';
  });
  return html+'</tbody></table>';
}

function renderAnomalies(type, val){
  const keys = getMonthKeys(type, val);
  const acc = {};
  keys.forEach(k => {
    (D.projects_per_month[k]||[]).filter(p=>p.rev>0).forEach(p=>{
      if(!acc[p.name]) acc[p.name]={name:p.name,rev:0,finSum:0};
      acc[p.name].rev+=p.rev; acc[p.name].finSum+=p.rev*(p.margin/100);
    });
  });
  const projects = Object.values(acc).map(a=>({
    name:a.name, rev:a.rev, fin:a.finSum,
    margin: a.rev>0 ? a.finSum/a.rev*100 : 0
  })).filter(p=>p.rev>0);

  const pv = prevPeriod(type, val);
  const prevKeys = pv ? getMonthKeys(type, pv) : [];
  const prevAcc = {};
  prevKeys.forEach(k=>{
    (D.projects_per_month[k]||[]).filter(p=>p.rev>0).forEach(p=>{
      if(!prevAcc[p.name]) prevAcc[p.name]={rev:0,finSum:0};
      prevAcc[p.name].rev+=p.rev; prevAcc[p.name].finSum+=p.rev*(p.margin/100);
    });
  });
  const prevMap = Object.fromEntries(
    Object.entries(prevAcc).map(([n,a])=>[n,{rev:a.rev, fin:a.finSum, margin:a.rev>0?a.finSum/a.rev*100:0}])
  );

  const mean=projects.length?projects.reduce((s,p)=>s+p.margin,0)/projects.length:0;
  document.getElementById('anomaly-stats').textContent =
    `${projects.length} проектов · ср. рент-ть ${fmtPct(mean)}`;

  // Блок 1: Выручка
  const byRev = [...projects].sort((a,b)=>b.rev-a.rev);
  document.getElementById('anomRevTop').innerHTML = anomSmallTable(byRev,'rev',true,6);
  document.getElementById('anomRevBot').innerHTML  = anomSmallTable(byRev,'rev',false,6);

  // Блок 2: Прибыль абс.
  const byFin = [...projects].sort((a,b)=>b.fin-a.fin);
  document.getElementById('anomProfTop').innerHTML = anomSmallTable(byFin,'prof',true,6);
  document.getElementById('anomProfBot').innerHTML  = anomSmallTable(byFin,'prof',false,6);

  // Блок 3: Рентабельность
  const byMar = [...projects].sort((a,b)=>b.margin-a.margin);
  document.getElementById('anomMarTop').innerHTML = anomSmallTable(byMar,'mar',true,6);
  document.getElementById('anomMarBot').innerHTML  = anomSmallTable(byMar,'mar',false,6);

  const noPrev = '<div style="color:var(--muted);font-size:11px;padding:8px">Нет предыдущего периода</div>';
  if(!prevKeys.length){
    ['anomRevGrowth','anomRevDrop','anomProfGrowth','anomProfDrop','anomMarGrowth','anomMarDrop']
      .forEach(id=>document.getElementById(id).innerHTML=noPrev);
    return;
  }

  const dynamics = projects.filter(p=>prevMap[p.name]).map(p=>{
    const pr=prevMap[p.name];
    return {...p, prevRev:pr.rev, prevFin:pr.fin,
      drev:p.rev-pr.rev, dprof:p.fin-pr.fin,
      dmar:+(p.margin-pr.margin).toFixed(2)};
  });

  const byDrev = [...dynamics].sort((a,b)=>b.drev-a.drev);
  document.getElementById('anomRevGrowth').innerHTML = anomSmallTable(byDrev,'drev',true,6);
  document.getElementById('anomRevDrop').innerHTML   = anomSmallTable(byDrev,'drev',false,6);

  const byDprof = [...dynamics].sort((a,b)=>b.dprof-a.dprof);
  document.getElementById('anomProfGrowth').innerHTML = anomSmallTable(byDprof,'dprof',true,6);
  document.getElementById('anomProfDrop').innerHTML   = anomSmallTable(byDprof,'dprof',false,6);

  const byDmar = [...dynamics].sort((a,b)=>b.dmar-a.dmar);
  document.getElementById('anomMarGrowth').innerHTML = anomSmallTable(byDmar,'dmar',true,6);
  document.getElementById('anomMarDrop').innerHTML   = anomSmallTable(byDmar,'dmar',false,6);
}

initFns.anomalies = function(){
  const selType = document.getElementById('anomPeriodType');
  const selVal  = document.getElementById('anomPeriodVal');
  function upd(){ if(selVal.value) renderAnomalies(selType.value, selVal.value); }
  fillPeriodSelect(selType, selVal, upd);
};

// ══════════════════════════════════════════════════════════════════════════════
// TAB: ПРОЕКТЫ
// ══════════════════════════════════════════════════════════════════════════════
let _psc='rev', _psa=false; // sort col, sort asc

initFns.projects = function(){
  const selType = document.getElementById('projPeriodType');
  const selVal  = document.getElementById('projPeriodVal');
  const srch    = document.getElementById('projSearchInput');

  function sortProjs(col){ if(_psc===col){_psa=!_psa;}else{_psc=col;_psa=false;} renderProjects(); }
  window._sortProjs = sortProjs;

  function renderProjects() {
    const type=selType.value, val=selVal.value;
    const keys=getMonthKeys(type,val);
    const pv=prevPeriod(type,val);
    const prevKeys=pv?getMonthKeys(type,pv):[];

    const acc={}, prevAcc={};
    keys.forEach(k=>(D.projects_per_month[k]||[]).forEach(p=>{
      if(!acc[p.name]) acc[p.name]={rev:0,finSum:0};
      acc[p.name].rev+=p.rev; acc[p.name].finSum+=p.rev*(p.margin/100);
    }));
    prevKeys.forEach(k=>(D.projects_per_month[k]||[]).forEach(p=>{
      if(!prevAcc[p.name]) prevAcc[p.name]={rev:0,finSum:0};
      prevAcc[p.name].rev+=p.rev; prevAcc[p.name].finSum+=p.rev*(p.margin/100);
    }));

    const fl=srch.value.toLowerCase();
    const rows=Object.entries(acc)
      .map(([name,a])=>{
        const margin=a.rev>0?a.finSum/a.rev*100:0;
        const fin=a.finSum;
        const pr=prevAcc[name];
        return {name,rev:a.rev,fin,margin,
          drev:pr?a.rev-pr.rev:null,
          dmar:pr?+(margin-(pr.rev>0?pr.finSum/pr.rev*100:0)).toFixed(2):null};
      })
      .filter(r=>!fl||r.name.toLowerCase().includes(fl))
      .sort((a,b)=>{
        const av=a[_psc]??0, bv=b[_psc]??0;
        const cmp=typeof av==='string'?av.localeCompare(bv,'ru'):(av-bv);
        return _psa?cmp:-cmp;
      });

    const totalRev=rows.reduce((s,r)=>s+r.rev,0);
    document.getElementById('proj-stats').textContent=
      `${rows.length} проектов · ₽${fmtM(totalRev)} выручка`;

    if(!rows.length){document.getElementById('projTableDiv').innerHTML='<div style="color:var(--muted);padding:20px">Нет данных</div>';return;}

    const sa=c=>c===_psc?(_psa?' ↑':' ↓'):'';
    let html=`<table><thead><tr>
      <th style="cursor:pointer" onclick="_sortProjs('name')">&#1055;&#1088;&#1086;&#1077;&#1082;&#1090;${sa('name')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortProjs('rev')">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;${sa('rev')}</th>
      <th class="td-r">&#1044;&#1086;&#1083;&#1103;</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortProjs('fin')">&#1055;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100;${sa('fin')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortProjs('margin')">&#1056;&#1077;&#1085;&#1090;-&#1090;&#1100;${sa('margin')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortProjs('drev')">&#916; &#1042;&#1099;&#1088;.${sa('drev')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortProjs('dmar')">&#916; &#1056;&#1077;&#1085;&#1090;.${sa('dmar')}</th>
    </tr></thead><tbody>`;

    rows.forEach(r=>{
      const chip=chipCls(r.margin);
      const share=totalRev>0?fmtPct(r.rev/totalRev*100):'—';
      const dRevStr=r.drev!==null?`<span class="${r.drev>=0?'pos':'neg'}">${sign(r.drev)}&#8381;${fmtM(r.drev)}</span>`:'—';
      const dMarStr=r.dmar!==null?`<span class="${r.dmar>=0?'pos':'neg'}">${sign(r.dmar)}${fmtPct(r.dmar)}</span>`:'—';
      const esc=r.name.replace(/'/g,"\\'");
      html+=`<tr>
        <td style="font-size:12px;max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.name}">
          <a class="proj-link" onclick="openProjectCard('${esc}')">${r.name}</a>
        </td>
        <td class="td-r">&#8381;${fmtM(r.rev)}</td>
        <td class="td-r" style="color:var(--muted);font-size:11px">${share}</td>
        <td class="td-r">&#8381;${fmtM(r.fin)}</td>
        <td class="td-r"><span class="chip ${chip}">${fmtPct(r.margin)}<span class="alloc-note" data-tip="&#1044;&#1086; &#1072;&#1083;&#1083;&#1086;&#1082;&#1072;&#1094;&#1080;&#1080;">&#174;&#1072;</span></span></td>
        <td class="td-r">${dRevStr}</td>
        <td class="td-r">${dMarStr}</td>
      </tr>`;
    });
    document.getElementById('projTableDiv').innerHTML=html+'</tbody></table>';
  }

  srch.oninput=renderProjects;
  fillPeriodSelect(selType, selVal, renderProjects);
};

// ══════════════════════════════════════════════════════════════════════════════
// TAB: ДИВИЗИОНЫ
// ══════════════════════════════════════════════════════════════════════════════
const CITY_DIVISION = {
  // Урал
  'Екатеринбург':'Урал','Челябинск':'Урал','Тюмень':'Урал','Пермь':'Урал','Уфа':'Урал',
  'Магнитогорск':'Урал','Стерлитамак':'Урал','Нижний Тагил':'Урал','Курган':'Урал',
  'Сургут':'Урал','Нефтеюганск':'Урал','Нижневартовск':'Урал','Тобольск':'Урал',
  // Юг
  'Ростов-на-Дону':'Юг','Краснодар':'Юг','Ставрополь':'Юг','Волгоград':'Юг',
  'Сочи':'Юг','Новороссийск':'Юг','Астрахань':'Юг','Симферополь':'Юг',
  'Севастополь':'Юг','Махачкала':'Юг','Нальчик':'Юг','Армавир':'Юг',
  // Центр
  'Москва':'Центр','Тула':'Центр','Ярославль':'Центр','Воронеж':'Центр',
  'Рязань':'Центр','Тверь':'Центр','Калуга':'Центр','Липецк':'Центр',
  'Брянск':'Центр','Орёл':'Центр','Курск':'Центр','Белгород':'Центр',
  'Иваново':'Центр','Кострома':'Центр','Владимир':'Центр','Смоленск':'Центр',
  // Сибирь
  'Новосибирск':'Сибирь','Омск':'Сибирь','Красноярск':'Сибирь','Иркутск':'Сибирь',
  'Томск':'Сибирь','Кемерово':'Сибирь','Барнаул':'Сибирь','Новокузнецк':'Сибирь',
  'Братск':'Сибирь','Улан-Удэ':'Сибирь','Чита':'Сибирь','Абакан':'Сибирь',
  // Поволжье
  'Казань':'Поволжье','Самара':'Поволжье','Нижний Новгород':'Поволжье',
  'Саратов':'Поволжье','Ульяновск':'Поволжье','Чебоксары':'Поволжье',
  'Тольятти':'Поволжье','Набережные Челны':'Поволжье','Пенза':'Поволжье',
  'Киров':'Поволжье','Йошкар-Ола':'Поволжье','Саранск':'Поволжье',
  // Северо-Запад
  'Санкт-Петербург':'Северо-Запад','Вологда':'Северо-Запад','Псков':'Северо-Запад',
  'Великий Новгород':'Северо-Запад','Мурманск':'Северо-Запад','Петрозаводск':'Северо-Запад',
  'Архангельск':'Северо-Запад','Калининград':'Северо-Запад','Сыктывкар':'Северо-Запад',
  // Дальний Восток
  'Владивосток':'Дальний Восток','Хабаровск':'Дальний Восток','Якутск':'Дальний Восток',
  'Магадан':'Дальний Восток','Южно-Сахалинск':'Дальний Восток','Петропавловск-Камчатский':'Дальний Восток',
};

const DIV_COLORS = {
  'Урал':'rgba(92,107,192,.75)','Юг':'rgba(239,83,80,.7)','Центр':'rgba(38,198,218,.7)',
  'Сибирь':'rgba(102,187,106,.7)','Поволжье':'rgba(255,167,38,.7)',
  'Северо-Запад':'rgba(171,71,188,.7)','Дальний Восток':'rgba(38,166,154,.7)',
  'Прочее':'rgba(120,130,150,.5)'
};

// Helper: aggregate projects → RRP for given keys
function aggregateDivisions(keys) {
  const _cityFn = D.cities.length ? buildCityExtractor(D.cities.map(c=>c.name)) : null;
  const getCity = n => {
    if(_cityFn){const r=_cityFn(n);if(r)return r;}
    const w=n.split(/[\s_]+/);return w[w.length-1]||n;
  };
  const acc={};
  keys.forEach(k=>{
    (D.projects_per_month[k]||[]).forEach(p=>{
      const div=p.rrp||CITY_DIVISION[getCity(p.name)]||'Не назначен';
      if(!acc[div]) acc[div]={div,revSum:0,finSum:0,cnt:0,cities:new Set(),projects:new Set()};
      acc[div].revSum+=p.rev; acc[div].finSum+=p.rev*(p.margin/100);
      acc[div].cnt++; acc[div].cities.add(getCity(p.name)); acc[div].projects.add(p.name);
    });
  });
  return Object.values(acc).map(d=>({
    ...d,margin:d.revSum>0?d.finSum/d.revSum*100:0,
    cities:[...d.cities],projects:[...d.projects]
  })).sort((a,b)=>b.revSum-a.revSum);
}

let _dsc='revSum', _dsa=false;

initFns.divisions = function(){
  const selType = document.getElementById('divPeriodType');
  const selVal  = document.getElementById('divPeriodVal');

  function sortDivs(col){ if(_dsc===col){_dsa=!_dsa;}else{_dsc=col;_dsa=false;} renderDivisions(); }
  window._sortDivs = sortDivs;

  function renderDivisions() {
    const type=selType.value, val=selVal.value;
    const keys=getMonthKeys(type,val);
    if(!keys.length) return;

    const pv=prevPeriod(type,val);
    const prevKeys=pv?getMonthKeys(type,pv):[];
    const prevDivs=Object.fromEntries(aggregateDivisions(prevKeys).map(d=>[d.div,d]));
    let divs=aggregateDivisions(keys);

    const totalRev=divs.reduce((s,d)=>s+d.revSum,0);
    document.getElementById('div-stats').textContent=
      `${divs.length} РРП · ₽${fmtM(totalRev)} выручка`;

    if(!divs.length){
      document.getElementById('divTable').innerHTML='<div style="color:var(--muted);padding:20px">Нет данных — РРП не определены</div>';
      destroyChart('divRevenue'); destroyChart('divMargin'); return;
    }

    // Sort
    divs=[...divs].sort((a,b)=>{
      const av=a[_dsc]??0, bv=b[_dsc]??0;
      const cmp=typeof av==='string'?av.localeCompare(bv,'ru'):(av-bv);
      return _dsa?cmp:-cmp;
    });

    destroyChart('divRevenue'); destroyChart('divMargin');
    const RRPPAL=['rgba(92,107,192,.75)','rgba(239,83,80,.7)','rgba(38,198,218,.7)',
      'rgba(102,187,106,.7)','rgba(255,167,38,.7)','rgba(171,71,188,.7)',
      'rgba(38,166,154,.7)','rgba(120,130,150,.5)'];
    const colors=divs.map((_,i)=>RRPPAL[i%RRPPAL.length]);

    charts.divRevenue=new Chart(document.getElementById('divRevenueChart').getContext('2d'),{
      type:'bar',data:{labels:divs.map(d=>d.div),datasets:[{
        label:'Выручка, тыс.₽',data:divs.map(d=>Math.round(d.revSum/1e3)),
        backgroundColor:colors,borderRadius:5}]},
      options:{...chartDefaults(),indexAxis:'y',plugins:{legend:{display:false}},scales:{
        x:{ticks:{color:()=>css('--muted'),callback:v=>'₽'+v.toLocaleString()},grid:{color:()=>css('--grid')}},
        y:{ticks:{color:()=>css('--text'),font:{size:12}},grid:{display:false}}}}
    });
    charts.divMargin=new Chart(document.getElementById('divMarginChart').getContext('2d'),{
      type:'bar',data:{labels:divs.map(d=>d.div),datasets:[{
        label:'Рент-ть %',data:divs.map(d=>+d.margin.toFixed(1)),
        backgroundColor:divs.map(d=>d.margin>=THRESHOLDS.good?'rgba(102,187,106,.75)':d.margin>=THRESHOLDS.warn?'rgba(255,167,38,.75)':'rgba(239,83,80,.65)'),
        borderRadius:5}]},
      options:{...chartDefaults(),indexAxis:'y',plugins:{legend:{display:false}},scales:{
        x:{ticks:{color:()=>css('--muted'),callback:v=>v+'%'},grid:{color:()=>css('--grid')}},
        y:{ticks:{color:()=>css('--text'),font:{size:12}},grid:{display:false}}}}
    });

    const sa=c=>c===_dsc?(_dsa?' ↑':' ↓'):'';
    let html=`<table><thead><tr>
      <th style="cursor:pointer" onclick="_sortDivs('div')">&#1056;&#1056;&#1055;${sa('div')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortDivs('revSum')">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;${sa('revSum')}</th>
      <th class="td-r">&#1044;&#1086;&#1083;&#1103;</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortDivs('finSum')">&#1055;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100;${sa('finSum')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortDivs('margin')">&#1056;&#1077;&#1085;&#1090;-&#1090;&#1100;${sa('margin')}</th>
      <th class="td-r">&#916; &#1042;&#1099;&#1088;.</th>
      <th class="td-r">&#916; &#1056;&#1077;&#1085;&#1090;.</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortDivs('cities')">&#1043;&#1086;&#1088;.${sa('cities')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortDivs('projects')">&#1055;&#1088;&#1086;&#1077;&#1082;&#1090;.${sa('projects')}</th>
      <th></th>
    </tr></thead><tbody>`;
    divs.forEach(d=>{
      const chip=chipCls(d.margin);
      const share=totalRev>0?fmtPct(d.revSum/totalRev*100):'—';
      const pr=prevDivs[d.div];
      const drev=pr?d.revSum-pr.revSum:null;
      const dmar=pr?+(d.margin-pr.margin).toFixed(2):null;
      const dRevStr=drev!==null?`<span class="${drev>=0?'pos':'neg'}">${sign(drev)}&#8381;${fmtM(drev)}</span>`:'—';
      const dMarStr=dmar!==null?`<span class="${dmar>=0?'pos':'neg'}">${sign(dmar)}${fmtPct(dmar)}</span>`:'—';
      const esc=d.div.replace(/'/g,"\\'");
      html+=`<tr>
        <td style="font-weight:600">${d.div}</td>
        <td class="td-r">&#8381;${fmtM(d.revSum)}</td>
        <td class="td-r" style="color:var(--muted);font-size:11px">${share}</td>
        <td class="td-r">&#8381;${fmtM(d.finSum)}</td>
        <td class="td-r"><span class="chip ${chip}">${fmtPct(d.margin)}</span></td>
        <td class="td-r">${dRevStr}</td>
        <td class="td-r">${dMarStr}</td>
        <td class="td-r" style="color:var(--muted)">${d.cities.length}</td>
        <td class="td-r" style="color:var(--muted)">${d.projects.length}</td>
        <td><button class="btn-sm" style="font-size:11px;padding:3px 8px" onclick="openDivisionCard('${esc}')">&#128100;</button></td>
      </tr>`;
    });
    document.getElementById('divTable').innerHTML=html+'</tbody></table>';
  }

  fillPeriodSelect(selType, selVal, renderDivisions);
};

// ══════════════════════════════════════════════════════════════════════════════
// TAB: КЛИЕНТЫ & ГОРОДА
// ══════════════════════════════════════════════════════════════════════════════
function hBarChart(id, items, colorFn){
  const top=items.slice(0,12);
  return new Chart(document.getElementById(id).getContext('2d'),{
    type:'bar',
    data:{
      labels:top.map(c=>c.name),
      datasets:[{
        label:'Выручка, тыс.₽',
        data:top.map(c=>Math.round(c.rev/1e3)),
        backgroundColor:top.map(colorFn),
        borderRadius:4
      }]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:ctx=>{
          const c=top[ctx.dataIndex];
          return [` ₽${ctx.parsed.x.toLocaleString('ru-RU')} тыс.`,` Рент-ть: ${fmtPct(c.margin)}`];
        }}}
      },
      scales:{
        x:{ticks:{color:()=>css('--muted'),callback:v=>'₽'+v.toLocaleString()},grid:{color:()=>css('--grid')}},
        y:{ticks:{color:()=>css('--text'),font:{size:11}},grid:{display:false}},
      }
    }
  });
}

function buildClientExtractor(refs) {
  const sorted = refs.slice().sort((a,b)=>b.length-a.length);
  return function(pname) {
    const lower = pname.toLowerCase();
    for (const ref of sorted) {
      const rl = ref.toLowerCase();
      if (lower === rl || lower.startsWith(rl+' ') || lower.startsWith(rl+'_')) return ref;
    }
    return null;
  };
}

function buildCityExtractor(refs) {
  const sorted = refs.slice().sort((a,b)=>b.length-a.length);
  return function(pname) {
    const lower = pname.toLowerCase();
    for (const ref of sorted) {
      const rl = ref.toLowerCase();
      if (lower.includes(' '+rl) || lower.includes('_'+rl) || lower === rl) return ref;
    }
    return null;
  };
}

function aggByField(keys, isCity) {
  const clientRefs = D.clients.map(c=>c.name);
  const cityRefs   = D.cities.map(c=>c.name);
  const _clientFn  = clientRefs.length ? buildClientExtractor(clientRefs) : null;
  const _cityFn    = cityRefs.length   ? buildCityExtractor(cityRefs)    : null;
  // fallbacks when SVOD ref doesn't contain the name
  const getClient = n => (_clientFn&&_clientFn(n)) || n.split(/[\s_]+/)[0];
  const getCity   = n => {
    if (_cityFn) { const r=_cityFn(n); if(r) return r; }
    const w = n.split(/[\s_]+/); return w[w.length-1]||n;
  };

  const acc = {};
  keys.forEach(k => {
    (D.projects_per_month[k]||[]).forEach(p => {
      const name = isCity ? getCity(p.name) : getClient(p.name);
      if (!name) return;
      if (!acc[name]) acc[name] = {name, revSum:0, finSum:0};
      acc[name].revSum += p.rev;
      acc[name].finSum += p.rev * (p.margin/100);
    });
  });
  return Object.values(acc)
    .filter(a => a.revSum > 0)
    .map(a => ({name:a.name, rev:a.revSum, margin: a.revSum>0 ? a.finSum/a.revSum*100 : 0}))
    .sort((a,b)=>b.rev-a.rev)
    .slice(0,15);
}

function aggAllClients(keys) {
  const clientRefs = D.clients.map(c=>c.name);
  const _clientFn  = clientRefs.length ? buildClientExtractor(clientRefs) : null;
  const getClient  = n => (_clientFn&&_clientFn(n)) || n.split(/[\s_]+/)[0];
  const acc = {};
  keys.forEach(k => {
    (D.projects_per_month[k]||[]).forEach(p => {
      const name = getClient(p.name);
      if (!name) return;
      if (!acc[name]) acc[name] = {name, revSum:0, finSum:0, projects:new Set()};
      acc[name].revSum += p.rev;
      acc[name].finSum += p.rev * (p.margin/100);
      acc[name].projects.add(p.name);
    });
  });
  return Object.values(acc)
    .filter(a => a.revSum > 0)
    .map(a => ({name:a.name, rev:a.revSum, fin:a.finSum,
      margin: a.revSum>0 ? a.finSum/a.revSum*100 : 0,
      projects:[...a.projects]}))
    .sort((a,b)=>b.rev-a.rev);
}

function exportCurrentClients(){
  const type=document.getElementById('clientsPeriodType').value;
  const val=document.getElementById('clientsPeriodVal').value;
  const rows=aggAllClients(getMonthKeys(type,val));
  downloadCSV(['\u041a\u043b\u0438\u0435\u043d\u0442','\u0412\u044b\u0440\u0443\u0447\u043a\u0430','\u041f\u0440\u0438\u0431\u044b\u043b\u044c','\u0420\u0435\u043d\u0442\u0430\u0431\u0435\u043b\u044c\u043d\u043e\u0441\u0442\u044c %','\u041f\u0440\u043e\u0435\u043a\u0442\u043e\u0432'],
    rows.map(r=>[r.name,Math.round(r.rev),Math.round(r.fin),r.margin.toFixed(1),r.projects.length]),
    `\u043a\u043b\u0438\u0435\u043d\u0442\u044b_${val}.csv`);
}

let _clsc='rev', _clsa=false;

initFns.clients = function(){
  const selType = document.getElementById('clientsPeriodType');
  const selVal  = document.getElementById('clientsPeriodVal');

  function sortCl(col){ if(_clsc===col){_clsa=!_clsa;}else{_clsc=col;_clsa=false;} renderClients(); }
  window._sortCl = sortCl;

  function renderClients() {
    const type = selType.value, val = selVal.value;
    const keys = getMonthKeys(type, val);
    const lbl  = getPeriodLabel(type, val);
    document.getElementById('clientsPeriodLabel').textContent = lbl;
    document.getElementById('citiesPeriodLabel').textContent  = lbl;

    const clients = aggByField(keys, false);
    const cities  = aggByField(keys, true);

    destroyChart('clients'); destroyChart('cities');
    charts.clients = hBarChart('clientsChart', clients,
      c=>c.margin>=THRESHOLDS.good?'rgba(102,187,106,.75)':c.margin>=THRESHOLDS.warn?'rgba(255,167,38,.75)':'rgba(239,83,80,.65)');
    charts.cities  = hBarChart('citiesChart', cities,
      c=>c.margin>=THRESHOLDS.good?'rgba(92,107,192,.75)':c.margin>=THRESHOLDS.warn?'rgba(38,198,218,.65)':'rgba(148,163,184,.5)');

    // Full clients table
    const allClients = aggAllClients(keys);
    const totalRev = allClients.reduce((s,c)=>s+c.rev,0);
    const sorted = [...allClients].sort((a,b)=>{
      const av=a[_clsc]??0, bv=b[_clsc]??0;
      const cmp=typeof av==='string'?av.localeCompare(bv,'ru'):(av-bv);
      return _clsa?cmp:-cmp;
    });
    if(!sorted.length){
      document.getElementById('clientsTableDiv').innerHTML='<div style="color:var(--muted);padding:20px">Нет данных</div>'; return;
    }
    const sa=c=>c===_clsc?(_clsa?' ↑':' ↓'):'';
    let html=`<table><thead><tr>
      <th style="cursor:pointer" onclick="_sortCl('name')">&#1050;&#1083;&#1080;&#1077;&#1085;&#1090;${sa('name')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortCl('rev')">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;${sa('rev')}</th>
      <th class="td-r">&#1044;&#1086;&#1083;&#1103;</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortCl('fin')">&#1055;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100;${sa('fin')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortCl('margin')">&#1056;&#1077;&#1085;&#1090;-&#1090;&#1100;${sa('margin')}</th>
      <th class="td-r" style="cursor:pointer" onclick="_sortCl('projects')">&#1055;&#1088;&#1086;&#1077;&#1082;&#1090;&#1086;&#1074;${sa('projects')}</th>
    </tr></thead><tbody>`;
    sorted.forEach(c=>{
      const chip=chipCls(c.margin);
      const share=totalRev>0?fmtPct(c.rev/totalRev*100):'—';
      const esc=c.name.replace(/'/g,"\\'");
      const projCnt=c.projects.length;
      html+=`<tr>
        <td style="font-size:12px;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${c.name}">
          <a class="proj-link" onclick="openClientCard('${esc}')">${c.name}</a>
        </td>
        <td class="td-r">&#8381;${fmtM(c.rev)}</td>
        <td class="td-r" style="color:var(--muted);font-size:11px">${share}</td>
        <td class="td-r">&#8381;${fmtM(c.fin)}</td>
        <td class="td-r"><span class="chip ${chip}">${fmtPct(c.margin)}</span></td>
        <td class="td-r" style="color:var(--muted)">${projCnt}</td>
      </tr>`;
    });
    document.getElementById('clientsTableDiv').innerHTML=html+'</tbody></table>';
  }

  fillPeriodSelect(selType, selVal, renderClients);
};

// ══════════════════════════════════════════════════════════════════════════════
// MODAL + КАРТОЧКИ
// ══════════════════════════════════════════════════════════════════════════════
function closeModal(){
  document.getElementById('modal-overlay').style.display='none';
  destroyChart('modalChart'); destroyChart('modalChart2');
  document.getElementById('modal-content').innerHTML='';
}

function modalKpi(items){
  return '<div class="kpi-grid" style="margin:16px 0">'+items.map(k=>`
    <div class="kpi" style="border-left-color:${k.color||'var(--accent)'}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value" style="font-size:20px">${k.val}</div>
      ${k.sub?`<div class="kpi-sub">${k.sub}</div>`:''}
    </div>`).join('')+'</div>';
}

function openProjectCard(name){
  const months=D.monthly;
  const entries=months.map(m=>{
    const p=(D.projects_per_month[m.key]||[]).find(x=>x.name===name);
    return {label:m.label,key:m.key,p};
  }).filter(e=>e.p);

  if(!entries.length){
    document.getElementById('modal-content').innerHTML=`<h2 style="margin-bottom:12px">${name}</h2><p style="color:var(--muted)">Нет данных по этому проекту</p>`;
    document.getElementById('modal-overlay').style.display='block';return;
  }

  const totalRev=entries.reduce((s,e)=>s+e.p.rev,0);
  const totalFin=entries.reduce((s,e)=>s+e.p.rev*(e.p.margin/100),0);
  const avgMar=totalRev>0?totalFin/totalRev*100:0;
  const compAvg=months.reduce((s,m)=>s+m.margin_pct,0)/months.length;
  const best=entries.reduce((a,b)=>a.p.margin>b.p.margin?a:b);
  const worst=entries.reduce((a,b)=>a.p.margin<b.p.margin?a:b);
  const lastE=entries[entries.length-1], prevE=entries.length>1?entries[entries.length-2]:null;

  let html=`<h2 style="margin-bottom:4px;font-size:18px">&#128193; ${name}</h2>
    <p style="color:var(--muted);font-size:12px;margin-bottom:0">${entries.length} активных месяцев · первый: ${entries[0].label} · последний: ${lastE.label}</p>`;

  html+=modalKpi([
    {label:'Выручка (итого)',val:'&#8381;'+fmtM(totalRev),sub:'Ср/мес: &#8381;'+fmtM(totalRev/entries.length),color:'var(--accent)'},
    {label:'Прибыль (итого)',val:'&#8381;'+fmtM(totalFin),sub:'Рент-ть: '+fmtPct(avgMar),color:'var(--green)'},
    {label:'Лучший месяц',val:best.label,sub:'Рент-ть: '+fmtPct(best.p.margin)+' · Выр: &#8381;'+fmtM(best.p.rev),color:'var(--teal)'},
    {label:'Рент. vs ср. по компании',val:(avgMar-compAvg>=0?'+':'')+fmtPct(avgMar-compAvg),
     sub:'Ср. компания: '+fmtPct(compAvg),color:avgMar>=compAvg?'var(--green)':'var(--red)'},
  ]);

  html+=`<div style="position:relative;height:260px;margin-bottom:18px"><canvas id="modalChart"></canvas></div>`;

  html+=`<div id="proj-modal-table"></div>`;

  document.getElementById('modal-content').innerHTML=html;
  document.getElementById('modal-overlay').style.display='block';

  // Sortable monthly table
  let _pmsc='key', _pmsa=false;
  const projEntries=[...entries].reverse().map((e,i,arr)=>{
    const prev=arr[i+1];
    const fin=e.p.rev*(e.p.margin/100);
    const dm=prev?+(e.p.margin-prev.p.margin).toFixed(2):null;
    return {label:e.label,key:e.key,rev:e.p.rev,fin,margin:e.p.margin,dm};
  });
  function rebuildProjModal(){
    const sorted=[...projEntries].sort((a,b)=>{
      const av=a[_pmsc]??0,bv=b[_pmsc]??0;
      const cmp=typeof av==='string'?av.localeCompare(bv,'ru'):(av-bv);
      return _pmsa?cmp:-cmp;
    });
    const sa=c=>c===_pmsc?(_pmsa?' ↑':' ↓'):'';
    let tbl=`<div class="tbl-wrap"><table style="font-size:12px"><thead><tr>
      <th style="cursor:pointer" onclick="window._sortPM('key')">&#1052;&#1077;&#1089;&#1103;&#1094;${sa('key')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortPM('rev')">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;${sa('rev')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortPM('fin')">&#1055;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100;${sa('fin')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortPM('margin')">&#1056;&#1077;&#1085;&#1090;-&#1090;&#1100;${sa('margin')}</th>
      <th class="td-r">&#916; &#1082; &#1087;&#1088;&#1077;&#1076;.</th>
    </tr></thead><tbody>`;
    sorted.forEach(r=>{
      const chip=chipCls(r.margin);
      tbl+=`<tr>
        <td>${r.label}</td>
        <td class="td-r">&#8381;${fmtM(r.rev)}</td>
        <td class="td-r">&#8381;${fmtM(r.fin)}</td>
        <td class="td-r"><span class="chip ${chip}">${fmtPct(r.margin)}</span></td>
        <td class="td-r">${r.dm!==null?`<span class="${r.dm>=0?'pos':'neg'}">${sign(r.dm)}${fmtPct(r.dm)}</span>`:'—'}</td>
      </tr>`;
    });
    tbl+='</tbody></table></div>';
    document.getElementById('proj-modal-table').innerHTML=tbl;
  }
  window._sortPM=function(col){if(_pmsc===col){_pmsa=!_pmsa;}else{_pmsc=col;_pmsa=false;}rebuildProjModal();};
  rebuildProjModal();

  setTimeout(()=>{
    const labels=entries.map(e=>e.label);
    charts.modalChart=new Chart(document.getElementById('modalChart').getContext('2d'),{
      data:{labels,datasets:[
        {type:'bar',label:'Выручка, тыс.₽',data:entries.map(e=>Math.round(e.p.rev/1e3)),
         backgroundColor:'rgba(92,107,192,.6)',yAxisID:'y',order:2},
        {type:'line',label:'Рент-ть %',data:entries.map(e=>e.p.margin),
         borderColor:'#66bb6a',backgroundColor:'rgba(102,187,106,.1)',fill:true,
         tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
        {type:'line',label:'Ср. компания %',data:entries.map(()=>+compAvg.toFixed(1)),
         borderColor:'rgba(255,167,38,.6)',borderDash:[6,4],
         fill:false,yAxisID:'y2',order:1,pointRadius:0,borderWidth:1.5},
      ]},
      options:{...chartDefaults(),scales:{
        x:scaleX(),
        y:{...scaleY('тыс. ₽'),position:'left'},
        y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false}},
      }}
    });
  },50);
}

function openDivisionCard(divName){
  // Gather all months data for this division
  const _cityFn=D.cities.length?buildCityExtractor(D.cities.map(c=>c.name)):null;
  const getCity=n=>{if(_cityFn){const r=_cityFn(n);if(r)return r;}const w=n.split(/[\s_]+/);return w[w.length-1]||n;};

  const monthData=D.monthly.map(m=>{
    const projs=(D.projects_per_month[m.key]||[]).filter(p=>{
      const d=p.rrp||CITY_DIVISION[getCity(p.name)]||'Не назначен';return d===divName;
    });
    const rev=projs.reduce((s,p)=>s+p.rev,0);
    const fin=projs.reduce((s,p)=>s+p.rev*(p.margin/100),0);
    return{label:m.label,key:m.key,rev,fin,margin:rev>0?fin/rev*100:0,projs};
  }).filter(m=>m.rev>0);

  if(!monthData.length){
    document.getElementById('modal-content').innerHTML=`<h2>${divName}</h2><p style="color:var(--muted)">Нет данных</p>`;
    document.getElementById('modal-overlay').style.display='block';return;
  }

  const totalRev=monthData.reduce((s,m)=>s+m.rev,0);
  const totalFin=monthData.reduce((s,m)=>s+m.fin,0);
  const avgMar=totalRev>0?totalFin/totalRev*100:0;

  const curYear=D.monthly[D.monthly.length-1].key.slice(3);
  const ytdM=monthData.filter(m=>m.key.slice(3)===curYear);
  const ytdRev=ytdM.reduce((s,m)=>s+m.rev,0);
  const ytdFin=ytdM.reduce((s,m)=>s+m.fin,0);
  const ytdMar=ytdRev>0?ytdFin/ytdRev*100:0;

  // Top cities
  const cityAcc={};
  monthData.forEach(m=>m.projs.forEach(p=>{
    const city=getCity(p.name);
    if(!cityAcc[city])cityAcc[city]={rev:0,fin:0};
    cityAcc[city].rev+=p.rev;cityAcc[city].fin+=p.rev*(p.margin/100);
  }));
  const topCities=Object.entries(cityAcc).map(([c,a])=>({city:c,rev:a.rev,margin:a.rev>0?a.fin/a.rev*100:0}))
    .sort((a,b)=>b.rev-a.rev).slice(0,8);

  // Top projects (last month in data)
  const projAcc={};
  monthData.forEach(m=>m.projs.forEach(p=>{
    if(!projAcc[p.name])projAcc[p.name]={rev:0,fin:0};
    projAcc[p.name].rev+=p.rev;projAcc[p.name].fin+=p.rev*(p.margin/100);
  }));
  const topProjs=Object.entries(projAcc).map(([n,a])=>({name:n,rev:a.rev,margin:a.rev>0?a.fin/a.rev*100:0}))
    .sort((a,b)=>b.rev-a.rev).slice(0,8);

  let html=`<h2 style="margin-bottom:4px;font-size:18px">&#128100; ${divName}</h2>
    <p style="color:var(--muted);font-size:12px;margin-bottom:0">&#1056;&#1056;&#1055; &middot; ${monthData.length} &#1072;&#1082;&#1090;&#1080;&#1074;&#1085;&#1099;&#1093; &#1084;&#1077;&#1089;&#1103;&#1094;&#1077;&#1074;</p>`;

  html+=modalKpi([
    {label:'Выручка (итого)',val:'&#8381;'+fmtM(totalRev),sub:'Ср/мес: &#8381;'+fmtM(totalRev/monthData.length),color:'var(--accent)'},
    {label:'Прибыль (итого)',val:'&#8381;'+fmtM(totalFin),sub:'Рент-ть: '+fmtPct(avgMar),color:'var(--green)'},
    {label:'YTD '+curYear,val:'&#8381;'+fmtM(ytdRev),sub:'Рент: '+fmtPct(ytdMar)+' · &#8381;'+fmtM(ytdFin)+' прибыль',color:'var(--teal)'},
    {label:'Проектов / Городов',val:Object.keys(projAcc).length+' / '+Object.keys(cityAcc).length,color:'var(--purple)'},
  ]);

  html+=`<div style="position:relative;height:240px;margin-bottom:18px"><canvas id="modalChart"></canvas></div>`;

  html+=`<div class="row2" style="gap:14px;margin-top:4px">
    <div>
      <div style="font-size:12px;font-weight:600;margin-bottom:8px">&#128204; &#1058;&#1086;&#1087; &#1075;&#1086;&#1088;&#1086;&#1076;&#1086;&#1074;</div>
      <table style="font-size:12px;width:100%"><thead><tr><th>&#1043;&#1086;&#1088;&#1086;&#1076;</th><th class="td-r">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;</th><th class="td-r">&#1056;&#1077;&#1085;&#1090;.</th></tr></thead><tbody>`;
  topCities.forEach(c=>{
    const chip=chipCls(c.margin);
    html+=`<tr><td>${c.city}</td><td class="td-r">&#8381;${fmtM(c.rev)}</td><td class="td-r"><span class="chip ${chip}">${fmtPct(c.margin)}</span></td></tr>`;
  });
  html+=`</tbody></table></div>
    <div>
      <div style="font-size:12px;font-weight:600;margin-bottom:8px">&#128203; &#1058;&#1086;&#1087; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1086;&#1074;</div>
      <table style="font-size:12px;width:100%"><thead><tr><th>&#1055;&#1088;&#1086;&#1077;&#1082;&#1090;</th><th class="td-r">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;</th><th class="td-r">&#1056;&#1077;&#1085;&#1090;.</th></tr></thead><tbody>`;
  topProjs.forEach(p=>{
    const chip=chipCls(p.margin);
    html+=`<tr><td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${p.name}</td>
      <td class="td-r">&#8381;${fmtM(p.rev)}</td><td class="td-r"><span class="chip ${chip}">${fmtPct(p.margin)}</span></td></tr>`;
  });
  html+=`</tbody></table></div></div>`;

  document.getElementById('modal-content').innerHTML=html;
  document.getElementById('modal-overlay').style.display='block';

  setTimeout(()=>{
    charts.modalChart=new Chart(document.getElementById('modalChart').getContext('2d'),{
      data:{labels:monthData.map(m=>m.label),datasets:[
        {type:'bar',label:'Выручка, тыс.₽',data:monthData.map(m=>Math.round(m.rev/1e3)),
         backgroundColor:'rgba(92,107,192,.6)',yAxisID:'y',order:2},
        {type:'line',label:'Рент-ть %',data:monthData.map(m=>+m.margin.toFixed(1)),
         borderColor:'#66bb6a',backgroundColor:'rgba(102,187,106,.1)',fill:true,
         tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
      ]},
      options:{...chartDefaults(),scales:{
        x:scaleX(),
        y:{...scaleY('тыс. ₽'),position:'left'},
        y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false}},
      }}
    });
  },50);
}

function openClientCard(clientName){
  const clientRefs = D.clients.map(c=>c.name);
  const _clientFn  = clientRefs.length ? buildClientExtractor(clientRefs) : null;
  const getClient  = n => (_clientFn&&_clientFn(n)) || n.split(/[\s_]+/)[0];

  // Gather monthly data for this client
  const monthData = D.monthly.map(m=>{
    const projs=(D.projects_per_month[m.key]||[]).filter(p=>getClient(p.name)===clientName);
    const rev=projs.reduce((s,p)=>s+p.rev,0);
    const fin=projs.reduce((s,p)=>s+p.rev*(p.margin/100),0);
    return {label:m.label,key:m.key,rev,fin,margin:rev>0?fin/rev*100:0,projs};
  }).filter(m=>m.rev>0);

  if(!monthData.length){
    document.getElementById('modal-content').innerHTML=`<h2 style="margin-bottom:12px">${clientName}</h2><p style="color:var(--muted)">Нет данных</p>`;
    document.getElementById('modal-overlay').style.display='block'; return;
  }

  const totalRev=monthData.reduce((s,m)=>s+m.rev,0);
  const totalFin=monthData.reduce((s,m)=>s+m.fin,0);
  const avgMar=totalRev>0?totalFin/totalRev*100:0;
  const compAvg=D.monthly.reduce((s,m)=>s+m.margin_pct,0)/D.monthly.length;
  const best=monthData.reduce((a,b)=>a.margin>b.margin?a:b);

  // Aggregate projects
  const projAcc={};
  monthData.forEach(m=>m.projs.forEach(p=>{
    if(!projAcc[p.name])projAcc[p.name]={rev:0,fin:0};
    projAcc[p.name].rev+=p.rev; projAcc[p.name].fin+=p.rev*(p.margin/100);
  }));
  let projList=Object.entries(projAcc).map(([n,a])=>({n,rev:a.rev,fin:a.fin,margin:a.rev>0?a.fin/a.rev*100:0}))
    .sort((a,b)=>b.rev-a.rev);

  let _cpsc='rev', _cpsa=false;
  function rebuildClientModal(){
    const sorted=[...projList].sort((a,b)=>{
      const av=a[_cpsc]??0,bv=b[_cpsc]??0;
      const cmp=typeof av==='string'?av.localeCompare(bv,'ru'):(av-bv);
      return _cpsa?cmp:-cmp;
    });
    const sa=c=>c===_cpsc?(_cpsa?' ↑':' ↓'):'';
    let tbl=`<div class="tbl-wrap"><table style="font-size:12px"><thead><tr>
      <th style="cursor:pointer" onclick="window._sortCP('n')">&#1055;&#1088;&#1086;&#1077;&#1082;&#1090;${sa('n')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortCP('rev')">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;${sa('rev')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortCP('fin')">&#1055;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100;${sa('fin')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortCP('margin')">&#1056;&#1077;&#1085;&#1090;-&#1090;&#1100;${sa('margin')}</th>
    </tr></thead><tbody>`;
    sorted.forEach(p=>{
      const chip=chipCls(p.margin);
      const esc=p.n.replace(/'/g,"\\'");
      tbl+=`<tr><td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${p.n}">
        <a class="proj-link" onclick="openProjectCard('${esc}')">${p.n}</a></td>
        <td class="td-r">&#8381;${fmtM(p.rev)}</td>
        <td class="td-r">&#8381;${fmtM(p.fin)}</td>
        <td class="td-r"><span class="chip ${chip}">${fmtPct(p.margin)}</span></td></tr>`;
    });
    tbl+='</tbody></table></div>';
    document.getElementById('client-proj-table').innerHTML=tbl;
  }
  window._sortCP=function(col){if(_cpsc===col){_cpsa=!_cpsa;}else{_cpsc=col;_cpsa=false;}rebuildClientModal();};

  // Monthly table sort
  let _cmsc='label', _cmsa=false;
  const mEntries=[...monthData].reverse();
  function rebuildMonthTable(){
    const sorted=[...mEntries].sort((a,b)=>{
      const av=a[_cmsc]??0, bv=b[_cmsc]??0;
      const cmp=typeof av==='string'?av.localeCompare(bv,'ru'):(av-bv);
      return _cmsa?cmp:-cmp;
    });
    const sa=c=>c===_cmsc?(_cmsa?' ↑':' ↓'):'';
    let tbl=`<div class="tbl-wrap"><table style="font-size:12px"><thead><tr>
      <th style="cursor:pointer" onclick="window._sortCM('label')">&#1052;&#1077;&#1089;&#1103;&#1094;${sa('label')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortCM('rev')">&#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;${sa('rev')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortCM('fin')">&#1055;&#1088;&#1080;&#1073;&#1099;&#1083;&#1100;${sa('fin')}</th>
      <th class="td-r" style="cursor:pointer" onclick="window._sortCM('margin')">&#1056;&#1077;&#1085;&#1090;-&#1090;&#1100;${sa('margin')}</th>
    </tr></thead><tbody>`;
    sorted.forEach(m=>{
      const chip=chipCls(m.margin);
      tbl+=`<tr><td>${m.label}</td>
        <td class="td-r">&#8381;${fmtM(m.rev)}</td>
        <td class="td-r">&#8381;${fmtM(m.fin)}</td>
        <td class="td-r"><span class="chip ${chip}">${fmtPct(m.margin)}</span></td></tr>`;
    });
    tbl+='</tbody></table></div>';
    document.getElementById('client-month-table').innerHTML=tbl;
  }
  window._sortCM=function(col){if(_cmsc===col){_cmsa=!_cmsa;}else{_cmsc=col;_cmsa=false;}rebuildMonthTable();};

  let html=`<h2 style="margin-bottom:4px;font-size:18px">&#127962; ${clientName}</h2>
    <p style="color:var(--muted);font-size:12px;margin-bottom:0">${monthData.length} активных месяцев · ${projList.length} проектов · первый: ${monthData[0].label} · последний: ${monthData[monthData.length-1].label}</p>`;

  html+=modalKpi([
    {label:'Выручка (итого)',val:'&#8381;'+fmtM(totalRev),sub:'Ср/мес: &#8381;'+fmtM(totalRev/monthData.length),color:'var(--accent)'},
    {label:'Прибыль (итого)',val:'&#8381;'+fmtM(totalFin),sub:'Рент-ть: '+fmtPct(avgMar),color:'var(--green)'},
    {label:'Лучший месяц',val:best.label,sub:'Рент-ть: '+fmtPct(best.margin)+' · &#8381;'+fmtM(best.rev),color:'var(--teal)'},
    {label:'Рент. vs ср. компания',val:(avgMar-compAvg>=0?'+':'')+fmtPct(avgMar-compAvg),
     sub:'Ср. компания: '+fmtPct(compAvg),color:avgMar>=compAvg?'var(--green)':'var(--red)'},
  ]);

  html+=`<div style="position:relative;height:240px;margin-bottom:18px"><canvas id="modalChart"></canvas></div>`;
  html+=`<div class="row2" style="gap:14px">
    <div>
      <div style="font-size:12px;font-weight:600;margin-bottom:8px">&#128203; &#1055;&#1088;&#1086;&#1077;&#1082;&#1090;&#1099;</div>
      <div id="client-proj-table"></div>
    </div>
    <div>
      <div style="font-size:12px;font-weight:600;margin-bottom:8px">&#128197; &#1055;&#1086; &#1084;&#1077;&#1089;&#1103;&#1094;&#1072;&#1084;</div>
      <div id="client-month-table"></div>
    </div>
  </div>`;

  document.getElementById('modal-content').innerHTML=html;
  document.getElementById('modal-overlay').style.display='block';
  rebuildClientModal();
  rebuildMonthTable();

  setTimeout(()=>{
    charts.modalChart=new Chart(document.getElementById('modalChart').getContext('2d'),{
      data:{labels:monthData.map(m=>m.label),datasets:[
        {type:'bar',label:'Выручка, тыс.₽',data:monthData.map(m=>Math.round(m.rev/1e3)),
         backgroundColor:'rgba(92,107,192,.6)',yAxisID:'y',order:2},
        {type:'line',label:'Рент-ть %',data:monthData.map(m=>+m.margin.toFixed(1)),
         borderColor:'#66bb6a',backgroundColor:'rgba(102,187,106,.1)',fill:true,
         tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
        {type:'line',label:'Ср. компания %',data:monthData.map(()=>+compAvg.toFixed(1)),
         borderColor:'rgba(255,167,38,.6)',borderDash:[6,4],
         fill:false,yAxisID:'y2',order:1,pointRadius:0,borderWidth:1.5},
      ]},
      options:{...chartDefaults(),scales:{
        x:scaleX(),
        y:{...scaleY('тыс. ₽'),position:'left'},
        y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false}},
      }}
    });
  },50);
}

// ── Custom tooltip ────────────────────────────────────────────────────────────
(function(){
  const tt = document.getElementById('tt');
  document.addEventListener('mouseover', e => {
    const el = e.target.closest('[data-tip]');
    if (!el) { tt.style.display='none'; return; }
    tt.textContent = el.dataset.tip;
    tt.style.display = 'block';
  });
  document.addEventListener('mousemove', e => {
    if (tt.style.display !== 'block') return;
    const x = e.clientX + 14, y = e.clientY + 14;
    const w = tt.offsetWidth, h = tt.offsetHeight;
    tt.style.left = (x + w > window.innerWidth  ? e.clientX - w - 8 : x) + 'px';
    tt.style.top  = (y + h > window.innerHeight ? e.clientY - h - 8 : y) + 'px';
  });
  document.addEventListener('mouseout', e => {
    if (e.target.closest('[data-tip]')) tt.style.display='none';
  });
})();

// ── Collapsible groups ────────────────────────────────────────────────────────
function toggleGroup(btn, parentB) {
  const table = btn.closest('table');
  const rows = table.querySelectorAll('tr[data-b]');
  const isExpanded = btn.dataset.expanded !== 'false';
  rows.forEach(tr => {
    const b = tr.dataset.b;
    if (b !== parentB && b.startsWith(parentB + '.')) {
      tr.style.display = isExpanded ? 'none' : '';
    }
  });
  btn.dataset.expanded = isExpanded ? 'false' : 'true';
  btn.textContent = isExpanded ? '\u25B6' : '\u25BC';
}

function toggleAllGroups(tableId, expand) {
  const table = document.getElementById(tableId);
  if (!table) return;
  table.querySelectorAll('.tog[data-expanded]').forEach(btn => {
    const parentB = btn.dataset.parentb;
    if (!parentB) return;
    const isExp = btn.dataset.expanded !== 'false';
    if (expand && !isExp) toggleGroup(btn, parentB);
    if (!expand && isExp) toggleGroup(btn, parentB);
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// Init
// ══════════════════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', ()=>{
  showPage('home');
});
</script>
</body>
</html>"""

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("STH-group Dashboard Generator v3")
    xl_readers = load_xl_files()
    if not xl_readers:
        print("ERROR: no Excel files found. Put .xlsx files into data/ subfolder."); return

    monthly_sheets = collect_monthly_sheets(xl_readers)
    if not monthly_sheets:
        print("ERROR: no monthly sheets (MM-YYYY format) found."); return
    print(f"  Found {len(monthly_sheets)} months across {len(xl_readers)} file(s)")

    print("  Monthly aggregates...")
    monthly = extract_monthly(monthly_sheets)
    print(f"    {len(monthly)} months")

    print("  Raw P&L rows (Общий tab)...")
    pl_rows = extract_pl_rows(monthly_sheets)
    print(f"    {sum(len(v) for v in pl_rows.values())} total rows")

    print("  Projects per month...")
    projects = extract_projects(monthly_sheets)
    print(f"    {sum(len(v) for v in projects.values())} project-month records")

    # SVOD sheets from the file that contains them (look for СВОД по клиентам)
    svod_xl = None
    for xl in reversed(xl_readers):
        if any('\u0421\u0412\u041e\u0414' in s for s in xl.sm):
            svod_xl = xl; break
    clients, cities = [], []
    if svod_xl:
        print("  СВОД по клиентам...")
        clients = extract_svod(svod_xl, '\u0421\u0412\u041e\u0414 \u043f\u043e \u043a\u043b\u0438\u0435\u043d\u0442\u0430\u043c')
        print(f"    {len(clients)} clients")
        print("  СВОД по городам...")
        cities = extract_svod(svod_xl, '\u0421\u0412\u041e\u0414 \u043f\u043e \u0433\u043e\u0440\u043e\u0434\u0430\u043c')
        print(f"    {len(cities)} cities")
    else:
        print("  СВОД sheets not found — clients/cities will be empty")

    for xl in xl_readers: xl.close()

    data = {
        'monthly': monthly,
        'pl_rows': pl_rows,
        'projects_per_month': projects,
        'clients': clients,
        'cities': cities,
    }
    json_str = json.dumps(data, ensure_ascii=False, separators=(',',':'))
    html_out = HTML.replace('___DATA___', json_str)

    with open(OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html_out)

    size_kb = os.path.getsize(OUTPUT) // 1024
    print(f"\nDone: {OUTPUT}  ({size_kb} KB)")

if __name__ == '__main__':
    main()

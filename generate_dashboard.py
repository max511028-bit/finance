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

# ── Data extraction ────────────────────────────────────────────────────────────

def extract_monthly(monthly_sheets):
    rows = []
    for xl, key, label in monthly_sheets:
        sd = xl.read(key)
        rev = g(sd,'1'); fin = g(sd,'18')
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
        projects = []
        for col, name in sd.get(3,{}).items():
            if col in ('A','B','C','D'): continue
            if not name or not isinstance(name,str) or not name.strip(): continue
            clean = name.split('_')[0].strip()
            rev = float(rev_row.get(col,0) or 0)
            if rev < 10000: continue
            mar = float(mar_row.get(col,0) or 0)
            fin = float(fin_row.get(col,0) or 0)
            mar = fix_margin(mar, rev, fin)
            projects.append({'name': clean, 'rev': round(rev), 'margin': round(mar*100, 2)})
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
  --bg:#f5f5f5;--card:#ffffff;--card2:#f8f9ff;--border:#e0e0e0;
  --text:#1a237e;--muted:#5c6bc0;--grid:rgba(0,0,0,0.06);
}
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

/* Tooltip via title - enhanced */
[title]{position:relative}

/* Anomaly sections */
.anom-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:900px){.anom-grid{grid-template-columns:1fr}}
.anom-stats{font-size:12px;color:var(--muted);margin-bottom:10px}

.footer{text-align:center;padding:16px;color:var(--muted);font-size:11px;
        border-top:1px solid var(--border);margin-top:10px}
</style>
</head>
<body>

<!-- Header -->
<header class="hdr">
  <div class="hdr-brand">
    <span style="font-size:20px">&#128202;</span>
    <div>
      <h1>STH-group</h1>
      <span class="sub">&#1056;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1086;&#1074;</span>
    </div>
  </div>
  <div class="hdr-right">
    <span class="badge" id="period-badge"></span>
    <button class="btn-sm" id="theme-btn" onclick="toggleTheme()">&#9728; &#1057;&#1074;&#1077;&#1090;&#1083;&#1072;&#1103;</button>
  </div>
</header>

<!-- Tab bar -->
<nav class="tab-bar">
  <button class="tab-btn active" onclick="showTab('main',this)">&#127968;&nbsp;&#1043;&#1083;&#1072;&#1074;&#1085;&#1072;&#1103;</button>
  <button class="tab-btn" onclick="showTab('trends',this)">&#128200;&nbsp;&#1058;&#1088;&#1077;&#1085;&#1076;&#1099;</button>
  <button class="tab-btn" onclick="showTab('general',this)">&#128203;&nbsp;&#1054;&#1073;&#1097;&#1080;&#1081;</button>
  <button class="tab-btn" onclick="showTab('compare',this)">&#9878;&#65039;&nbsp;&#1057;&#1088;&#1072;&#1074;&#1085;&#1077;&#1085;&#1080;&#1077;</button>
  <button class="tab-btn" onclick="showTab('anomalies',this)">&#128269;&nbsp;&#1040;&#1085;&#1086;&#1084;&#1072;&#1083;&#1080;&#1080;</button>
  <button class="tab-btn" onclick="showTab('clients',this)">&#127962;&nbsp;&#1050;&#1083;&#1080;&#1077;&#1085;&#1090;&#1099;&nbsp;&amp;&nbsp;&#1043;&#1086;&#1088;&#1086;&#1076;&#1072;</button>
</nav>

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
    <div class="card">
      <div class="card-title" style="margin-bottom:12px">&#128201; &#1058;&#1088;&#1077;&#1085;&#1076; (&#1087;&#1086;&#1089;&#1083;&#1077;&#1076;&#1085;&#1080;&#1077; 6 &#1084;&#1077;&#1089;&#1103;&#1094;&#1077;&#1074;)</div>
      <div class="chart-sm"><canvas id="miniChart"></canvas></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: ТРЕНДЫ ════════════════════════════════════════════════════════════ -->
<div id="tab-trends" class="tab-content">
  <div class="gap">
    <div class="card">
      <div class="card-title" style="margin-bottom:12px">&#128200; &#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072;, &#1088;&#1072;&#1089;&#1093;&#1086;&#1076;&#1099; &#1080; &#1088;&#1077;&#1085;&#1090;&#1072;&#1073;&#1077;&#1083;&#1100;&#1085;&#1086;&#1089;&#1090;&#1100; &mdash; <span id="trends-period"></span></div>
      <div class="chart-wrap"><canvas id="trendsChart"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title" style="margin-bottom:12px">&#128188; &#1047;&#1055; / &#1042;&#1099;&#1088;&#1091;&#1095;&#1082;&#1072; &#1080; &#1092;&#1080;&#1085;&#1072;&#1085;&#1089;&#1086;&#1074;&#1099;&#1081; &#1088;&#1077;&#1079;&#1091;&#1083;&#1100;&#1090;&#1072;&#1090;</div>
      <div class="chart-wrap"><canvas id="zpChart"></canvas></div>
    </div>
  </div>
</div>

<!-- ═══ TAB: ОБЩИЙ ═════════════════════════════════════════════════════════════ -->
<div id="tab-general" class="tab-content">
  <div class="gap">
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">&#128203; &#1042;&#1089;&#1077; &#1087;&#1086;&#1082;&#1072;&#1079;&#1072;&#1090;&#1077;&#1083;&#1080; &#1079;&#1072; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;</div>
        <div class="ctrl" style="margin:0">
          <span class="lbl">&#1058;&#1080;&#1087;:</span>
          <select id="generalPeriodType"><option value="month">&#1052;&#1077;&#1089;&#1103;&#1094;</option><option value="quarter">&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;</option><option value="year">&#1043;&#1086;&#1076;</option></select>
          <select id="generalPeriodVal"></select>
        </div>
      </div>
      <p style="font-size:11px;color:var(--muted);margin-bottom:10px">
        &#1053;&#1072;&#1074;&#1077;&#1076;&#1080;&#1090;&#1077; &#1085;&#1072; &#1085;&#1072;&#1079;&#1074;&#1072;&#1085;&#1080;&#1077; &#1087;&#1086;&#1083;&#1103; &mdash; &#1087;&#1086;&#1103;&#1074;&#1080;&#1090;&#1089;&#1103; &#1087;&#1086;&#1076;&#1089;&#1082;&#1072;&#1079;&#1082;&#1072; &#1089; &#1086;&#1087;&#1080;&#1089;&#1072;&#1085;&#1080;&#1077;&#1084;.
      </p>
      <div class="tbl-wrap"><table id="generalTable"></table></div>
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
  </div>
</div>

<!-- ═══ TAB: АНОМАЛИИ ══════════════════════════════════════════════════════════ -->
<div id="tab-anomalies" class="tab-content">
  <div class="gap">
    <div class="card">
      <div class="card-hdr">
        <div class="card-title">&#128269; &#1040;&#1085;&#1086;&#1084;&#1072;&#1083;&#1080;&#1080; &#1080; &#1076;&#1080;&#1085;&#1072;&#1084;&#1080;&#1082;&#1072; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1086;&#1074;</div>
        <div class="ctrl" style="margin:0">
          <span class="lbl">&#1052;&#1077;&#1089;&#1103;&#1094;:</span>
          <select id="anomalyMonth"></select>
        </div>
      </div>
      <div class="anom-stats" id="anomaly-stats"></div>
      <div class="anom-grid">
        <div>
          <div style="font-size:12px;font-weight:600;color:var(--red);margin-bottom:8px">&#11015; &#1055;&#1088;&#1086;&#1073;&#1083;&#1077;&#1084;&#1085;&#1099;&#1077; &#1087;&#1088;&#1086;&#1077;&#1082;&#1090;&#1099;</div>
          <div class="tbl-wrap"><div id="anomBad"></div></div>
        </div>
        <div>
          <div style="font-size:12px;font-weight:600;color:var(--green);margin-bottom:8px">&#11014; &#1051;&#1080;&#1076;&#1077;&#1088;&#1099;</div>
          <div class="tbl-wrap"><div id="anomGood"></div></div>
        </div>
      </div>
      <div style="margin-top:20px">
        <div style="font-size:12px;font-weight:600;color:var(--accent);margin-bottom:8px">&#128260; &#1044;&#1080;&#1085;&#1072;&#1084;&#1080;&#1082;&#1072; vs &#1087;&#1088;&#1077;&#1076;&#1099;&#1076;&#1091;&#1097;&#1080;&#1081; &#1084;&#1077;&#1089;&#1103;&#1094;</div>
        <div class="tbl-wrap"><div id="anomDynamic"></div></div>
      </div>
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

function fillPeriodSelect(selType, selVal, onChange, defaultFirst) {
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
    if (defaultFirst && selVal.options.length > 1 && !defaultFirst._set) {
      defaultFirst._set = true;
    }
    onChange();
  }
  selType.onchange = rebuild;
  selVal.onchange  = onChange;
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
  const cls = m>=15?'chip-g':m>=8?'chip-a':'chip-r';
  return `<span class="chip ${cls}">${fmtPct(m)}</span>`;
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
  document.getElementById('trends-period').textContent = `${firstLabel} → ${lastLabel}`;

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
  const labels = D.monthly.map(m=>m.label);
  const rev    = D.monthly.map(m=>Math.round(m.revenue/1e3));
  const costs  = D.monthly.map(m=>Math.round(m.project_costs/1e3));
  const margin = D.monthly.map(m=>m.margin_pct);
  const zpRev  = D.monthly.map(m=>m.zp_rev_pct);
  const finBiz = D.monthly.map(m=>Math.round(m.fin_res_biz/1e3));

  charts.trends = new Chart(document.getElementById('trendsChart').getContext('2d'),{
    data:{labels,datasets:[
      {type:'bar',label:'Выручка, тыс.₽',data:rev,backgroundColor:'rgba(92,107,192,.65)',yAxisID:'y',order:2},
      {type:'bar',label:'Расходы, тыс.₽',data:costs,backgroundColor:'rgba(71,85,105,.55)',yAxisID:'y',order:2},
      {type:'line',label:'Рент-ть %',data:margin,borderColor:'#66bb6a',backgroundColor:'rgba(102,187,106,.08)',
       fill:true,tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
      {type:'line',label:'ЗП/Выручка %',data:zpRev,borderColor:'#ffa726',
       tension:.3,borderDash:[5,3],yAxisID:'y2',order:1,pointRadius:3,borderWidth:2},
    ]},
    options:{...chartDefaults(),scales:{
      x:scaleX(),
      y:{...scaleY('тыс. ₽'),position:'left'},
      y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false},min:0,
          max:Math.ceil(Math.max(...zpRev,100)/10)*10},
    }}
  });

  charts.zp = new Chart(document.getElementById('zpChart').getContext('2d'),{
    data:{labels,datasets:[
      {type:'bar',label:'Фин. рез. бизнеса, тыс.₽',data:finBiz,backgroundColor:'rgba(38,198,218,.5)',yAxisID:'y',order:2},
      {type:'line',label:'ЗП масс. / Выручка %',data:zpRev,borderColor:'#ffa726',
       backgroundColor:'rgba(255,167,38,.1)',fill:true,tension:.3,yAxisID:'y2',order:1,pointRadius:4,borderWidth:2},
    ]},
    options:{...chartDefaults(),scales:{
      x:scaleX(),
      y:{...scaleY('тыс. ₽'),position:'left'},
      y2:{...scaleY('%'),position:'right',grid:{drawOnChartArea:false},min:0},
    }}
  });
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
    document.getElementById('generalTable').innerHTML = '<tbody><tr><td colspan="4" style="color:var(--muted);padding:20px">Нет данных для этого периода</td></tr></tbody>';
    return;
  }
  const revRow = rows.find(r => r.b === '1');
  const rev = revRow ? revRow.value : 0;

  let html = `<thead><tr>
    <th class="pl-code">Код</th>
    <th>Показатель</th>
    <th class="td-r">Значение</th>
    <th class="pl-share">Доля выручки</th>
  </tr></thead><tbody>`;

  rows.forEach(r => {
    const isZero = r.value === 0 || r.value === 0.0;

    const indent = r.indent * 18 + 4;
    const isTopLevel = r.indent === 0;
    const isTotal = isTopLevel && /^\d+$/.test(r.b);

    // Detect percentage: value between -1.5 and 1.5 (stored as decimal, e.g. 0.287 = 28.7%)
    const isDecPct = !isZero && Math.abs(r.value) <= 1.5;

    let valStr, shareStr = '—';
    if(isDecPct){
      const pv = r.value * 100;
      const cls = pv >= 15 ? 'pos' : pv >= 5 ? '' : pv < 0 ? 'neg' : '';
      valStr = `<span class="${cls}" style="font-weight:600">${fmtPct(pv)}</span>`;
    } else {
      const cls = r.value < 0 ? 'neg' : '';
      valStr = `<span class="${cls}">${fmtRub(r.value)}</span>`;
      if(rev > 0 && r.b !== '1'){
        shareStr = fmtPct(Math.abs(r.value) / rev * 100);
      }
    }

    const tip = PL_DESC[r.b] || ('Код ' + r.b);
    const rowCls = [
      isZero ? 'pl-zero' : '',
      isTotal ? 'pl-top0-total' : isTopLevel ? 'pl-top0' : '',
    ].filter(Boolean).join(' ');

    html += `<tr class="${rowCls}">
      <td class="pl-code">${r.b}</td>
      <td style="padding-left:${indent}px">
        <span class="pl-label" title="${tip}">${r.label}</span>
      </td>
      <td class="pl-val">${valStr}</td>
      <td class="pl-share">${shareStr}</td>
    </tr>`;
  });

  html += '</tbody>';
  document.getElementById('generalTable').innerHTML = html;
}

initFns.general = function(){
  const selType = document.getElementById('generalPeriodType');
  const selVal  = document.getElementById('generalPeriodVal');
  function upd() {
    const keys = getMonthKeys(selType.value, selVal.value);
    renderGeneral(aggregatePlRows(keys));
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

  let html = `<table><thead><tr>
    <th class="pl-code">&#1050;&#1086;&#1076;</th>
    <th>&#1055;&#1086;&#1082;&#1072;&#1079;&#1072;&#1090;&#1077;&#1083;&#1100;</th>
    <th class="td-r">${labelA}</th>
    <th class="td-r">${labelB}</th>
    <th class="td-r">&#916;</th>
    <th class="td-r">&#916;%</th>
  </tr></thead><tbody>`;

  order.forEach(r => {
    const rA = mapA[r.b], rB = mapB[r.b];
    if (!rA && !rB) return;
    const refVal = rA ? rA.value : rB.value;
    const isDecPct = Math.abs(refVal) <= 1.5 && refVal !== 0;
    const fv = isDecPct ? v => fmtPct(v*100) : fmtRub;
    const indent = r.indent * 18 + 4;
    const tip = PL_DESC[r.b] || ('&#1050;&#1086;&#1076; ' + r.b);
    const rowCls = r.indent===0 ? (/^\d+$/.test(r.b)?'pl-top0-total':'pl-top0') : '';

    let vAstr, vBstr, dStr = '&#8212;', dpStr = '&#8212;';
    if (!rA) {
      vAstr = `<span style="color:var(--muted)" title="&#1085;&#1077;&#1090; &#1074; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;&#1077; &#1040;">&#8212;</span>`;
      vBstr = fv(rB.value);
    } else if (!rB) {
      vAstr = fv(rA.value);
      vBstr = `<span style="color:var(--muted)" title="&#1085;&#1077;&#1090; &#1074; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;&#1077; &#1041;">&#8212;</span>`;
    } else {
      const va=rA.value, vb=rB.value, d=vb-va;
      const dp = va!==0 ? d/Math.abs(va)*100 : 0;
      const cls = d>=0?'pos':'neg';
      vAstr = fv(va); vBstr = fv(vb);
      dStr  = `<span class="${cls}">${sign(d)}${fv(d)}</span>`;
      dpStr = `<span class="${cls}">${sign(dp)}${dp.toFixed(1)}%</span>`;
    }

    html += `<tr class="${rowCls}">
      <td class="pl-code">${r.b}</td>
      <td style="padding-left:${indent}px"><span class="pl-label" title="${tip}">${r.label}</span></td>
      <td class="td-r">${vAstr}</td>
      <td class="td-r">${vBstr}</td>
      <td class="td-r">${dStr}</td>
      <td class="td-r">${dpStr}</td>
    </tr>`;
  });

  html += '</tbody></table>';
  document.getElementById('compareTable').innerHTML = html;
}

initFns.compare = function(){
  const tA=document.getElementById('cmpTypeA'), vA=document.getElementById('cmpValA');
  const tB=document.getElementById('cmpTypeB'), vB=document.getElementById('cmpValB');
  let _initB = false;
  function upd() {
    const keysA = getMonthKeys(tA.value, vA.value);
    const keysB = getMonthKeys(tB.value, vB.value);
    renderCompare(keysA, keysB, getPeriodLabel(tA.value,vA.value), getPeriodLabel(tB.value,vB.value));
  }
  // Fill A first, then B with offset -1 default
  fillPeriodSelect(tA, vA, upd);
  fillPeriodSelect(tB, vB, upd);
  // Set B default to second option (previous month)
  if (vB.options.length > 1) { vB.selectedIndex = 1; upd(); }
};

// ══════════════════════════════════════════════════════════════════════════════
// TAB: АНОМАЛИИ
// ══════════════════════════════════════════════════════════════════════════════
function projTable(projects, showDev, showDelta){
  if(!projects.length) return '<div style="color:var(--muted);padding:12px;font-size:12px">Нет данных</div>';
  let html=`<table><thead><tr><th>Проект</th><th class="td-r">Выручка</th><th class="td-r">Рент-ть</th>`;
  if(showDev)   html+='<th>Откл. от ср.</th>';
  if(showDelta) html+='<th class="td-r">Δ к пред.</th>';
  html+='</tr></thead><tbody>';
  projects.slice(0,10).forEach(p=>{
    const chip=p.margin>=15?'chip-g':p.margin>=8?'chip-a':'chip-r';
    html+=`<tr><td class="td-name" style="font-size:12px">${p.name}</td>
      <td class="td-r" style="font-size:12px">₽${fmtM(p.rev)}</td>
      <td class="td-r"><span class="chip ${chip}">${fmtPct(p.margin)}</span></td>`;
    if(showDev && p.dev!==undefined){
      const bw=Math.round(Math.abs(p.dev)/Math.max(...projects.map(x=>Math.abs(x.dev||0)),1)*60);
      const bc=p.dev>=0?'var(--green)':'var(--red)';
      html+=`<td><div class="devbar"><div class="devbar-inner" style="width:${bw}px;background:${bc}"></div>
        <span class="${p.dev>=0?'pos':'neg'}" style="font-size:11px">${sign(p.dev)}${fmtPct(p.dev)}</span></div></td>`;
    }
    if(showDelta && p.delta!==undefined){
      html+=`<td class="td-r ${p.delta>=0?'pos':'neg'}" style="font-weight:600">
        ${p.delta>=0?'▲':'▼'} ${sign(p.delta)}${fmtPct(p.delta)}</td>`;
    }
    html+='</tr>';
  });
  return html+'</tbody></table>';
}

function renderAnomalies(key){
  const projects=(D.projects_per_month[key]||[]).filter(p=>p.rev>0);
  const months=D.monthly.map(m=>m.key);
  const idx=months.indexOf(key);
  const prevKey=idx>0?months[idx-1]:null;
  const prev=prevKey?(D.projects_per_month[prevKey]||[]):[];

  const margins=projects.filter(p=>p.margin!==0).map(p=>p.margin);
  const mean=margins.reduce((s,v)=>s+v,0)/(margins.length||1);
  const std=Math.sqrt(margins.map(v=>(v-mean)**2).reduce((s,v)=>s+v,0)/(margins.length||1));

  document.getElementById('anomaly-stats').textContent =
    `${projects.length} проектов · средняя рент-ть ${fmtPct(mean)} · σ = ${fmtPct(std)}`;

  const withDev=projects.map(p=>({...p,dev:p.margin-mean}));
  const bad=withDev.filter(p=>p.dev<0).sort((a,b)=>a.dev-b.dev);
  const good=withDev.filter(p=>p.dev>0).sort((a,b)=>b.dev-a.dev);

  document.getElementById('anomBad').innerHTML=projTable(bad, true, false);
  document.getElementById('anomGood').innerHTML=projTable(good, true, false);

  let dynHTML='<div style="color:var(--muted);font-size:12px;padding:8px">Выберите месяц кроме первого для просмотра динамики</div>';
  if(prev.length>0){
    const prevMap=Object.fromEntries(prev.map(p=>[p.name,p.margin]));
    const dynamics=projects
      .filter(p=>prevMap[p.name]!==undefined)
      .map(p=>({...p,delta:+(p.margin-prevMap[p.name]).toFixed(2)}))
      .sort((a,b)=>Math.abs(b.delta)-Math.abs(a.delta));
    dynHTML=projTable(dynamics, false, true);
  }
  document.getElementById('anomDynamic').innerHTML=dynHTML;
}

initFns.anomalies = function(){
  const sel=document.getElementById('anomalyMonth');
  D.monthly.slice().reverse().forEach(m=>{
    if((D.projects_per_month[m.key]||[]).length>0)
      sel.innerHTML+=`<option value="${m.key}">${m.label}</option>`;
  });
  sel.onchange=()=>renderAnomalies(sel.value);
  if(sel.options.length) renderAnomalies(sel.value);
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

function aggByField(keys, fieldIdx) {
  const acc = {};
  keys.forEach(k => {
    (D.projects_per_month[k]||[]).forEach(p => {
      const parts = p.name.split('_');
      const name  = (parts[fieldIdx]||'').trim();
      if (!name) return;
      if (!acc[name]) acc[name] = {name, revSum:0, finSum:0};
      acc[name].revSum += p.rev;
      // Approximate fin from margin%: fin = rev * margin/100
      acc[name].finSum += p.rev * (p.margin/100);
    });
  });
  return Object.values(acc)
    .map(a => ({name:a.name, rev:a.revSum, margin: a.revSum>0 ? a.finSum/a.revSum*100 : 0}))
    .sort((a,b)=>b.rev-a.rev)
    .slice(0,15);
}

initFns.clients = function(){
  const selType = document.getElementById('clientsPeriodType');
  const selVal  = document.getElementById('clientsPeriodVal');

  function renderClients() {
    const type = selType.value, val = selVal.value;
    const keys = getMonthKeys(type, val);
    const lbl  = getPeriodLabel(type, val);
    document.getElementById('clientsPeriodLabel').textContent = lbl;
    document.getElementById('citiesPeriodLabel').textContent  = lbl;

    const clients = aggByField(keys, 0);
    const cities  = aggByField(keys, 1);

    destroyChart('clientsChart'); destroyChart('citiesChart');
    charts.clients = hBarChart('clientsChart', clients,
      c=>c.margin>=15?'rgba(102,187,106,.75)':c.margin>=8?'rgba(255,167,38,.75)':'rgba(239,83,80,.65)');
    charts.cities  = hBarChart('citiesChart', cities,
      c=>c.margin>=15?'rgba(92,107,192,.75)':c.margin>=8?'rgba(38,198,218,.65)':'rgba(148,163,184,.5)');
  }

  fillPeriodSelect(selType, selVal, renderClients);
};

// ══════════════════════════════════════════════════════════════════════════════
// Init
// ══════════════════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', ()=>{
  initFns.main(); inited.main=true;
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

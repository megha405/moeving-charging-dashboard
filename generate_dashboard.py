#!/usr/bin/env python3
"""MoEVing EV Charging Dashboard Generator — run this each morning after dropping new data files."""

import os, json, csv, re, sys
from datetime import datetime

try:
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl -q")
    import openpyxl

FOLDER = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(FOLDER, "dashboard.html")

MONTHS = {"jan":"01","feb":"02","mar":"03","apr":"04","may":"05","jun":"06",
           "jul":"07","aug":"08","sep":"09","oct":"10","nov":"11","dec":"12"}

PARTNER_COLORS = {
    "JioBP":                    "#2563eb",
    "Statiq":                   "#7c3aed",
    "ChargeZone":               "#d97706",
    "BPCL":                     "#dc2626",
    "HPCL eCharge":             "#16a34a",
    "Shell Recharge":           "#ea580c",
    "Adani":                    "#0891b2",
    "Charge IN by Mahindra":    "#db2777",
    "MoEVing Own":              "#64748b",
    "EcoPlug":                  "#84cc16",
    "Gentari":                  "#0d9488",
    "Tata Power":               "#f59e0b",
    "BESCOM":                   "#6366f1",
}

def parse_date(s):
    if not s: return None
    s = str(s).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m: return s[:10]
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m: return f"{m.group(3)}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    m = re.match(r"^(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+(\d{4})", s, re.I)
    if m: return f"{m.group(3)}-{MONTHS[m.group(2).lower()]}-{m.group(1).zfill(2)}"
    return None

def sf(v):
    if v is None: return None
    s = str(v).strip()
    if s in ("", "null", "None", "-", "N/A"): return None
    try: return float(s)
    except: return None

def norm_partner(s):
    if not s: return "Unknown"
    s = str(s)
    # Strip CPO role suffix like ", CPO_eMSP" or " [EVLINQ]"
    s = re.sub(r"\s*[\[,].*", "", s).strip()
    sl = s.lower()
    if "jio" in sl: return "JioBP"
    if "statiq" in sl: return "Statiq"
    if "chargezone" in sl or "charge zone" in sl: return "ChargeZone"
    if "bpcl" in sl: return "BPCL"
    if "hpcl" in sl: return "HPCL eCharge"
    if "shell" in sl: return "Shell Recharge"
    if "adani" in sl: return "Adani"
    if "mahindra" in sl or "charge in" in sl or "charge_in" in sl: return "Charge IN by Mahindra"
    if "moeving" in sl or "itc" in sl: return "MoEVing Own"
    if "bescom" in sl: return "BESCOM"
    if "ecoplug" in sl or "ecocharg" in sl or "earthtron" in sl: return "EcoPlug"
    if "gentari" in sl or "numocity" in sl: return "Gentari"
    if "fastwatt" in sl or "bhawna" in sl or "eesl" in sl or "glida" in sl or "charj" in sl: return s
    if "tata" in sl: return "Tata Power"
    return s

def detect_format(headers):
    hs = set(str(h).strip() for h in headers if h)
    if "TRANSACTION_ID" in hs and "SALE_OF_ENERGY" in hs: return "jiobp_south"
    if "Transaction ID" in hs and "Sale Of Energy" in hs: return "jiobp"
    if "Booking ID" in hs and "CPO/Platform" in hs: return "statiq"
    if "CPO" in hs and "Charging Session Cost" in hs: return "kpi_dashboard"
    if "PROFILE_ID" in hs and "groupId" in hs: return "tata_summary"
    return "unknown"

def parse_jiobp(data_rows, headers):
    h = {str(v).strip(): i for i, v in enumerate(headers)}
    g = lambda row, k: row[h[k]] if k in h and h[k] < len(row) else None
    out = []
    for row in data_rows:
        kwh = sf(g(row, "Units Consumed(kWh)"))
        if not kwh or kwh <= 0: continue
        cost = sf(g(row, "Sale Of Energy"))
        tariff = sf(g(row, "Tariff Rate"))
        vrn = str(g(row, "VRN") or "").strip()
        out.append({
            "date":         parse_date(g(row, "Transaction Date")),
            "partner":      "JioBP",
            "station":      str(g(row, "Station") or "").strip(),
            "city":         str(g(row, "City") or "").strip(),
            "state":        str(g(row, "State") or "").strip(),
            "vrn":          vrn if vrn not in ("", "-") else None,
            "make":         str(g(row, "Make") or "").strip(),
            "model":        str(g(row, "Model") or "").strip(),
            "driver":       str(g(row, "Customer Name") or "").strip(),
            "kwh":          kwh,
            "cost":         cost,
            "cost_per_kwh": tariff if tariff else (round(cost/kwh, 4) if cost and kwh else None),
            "duration":     str(g(row, "Session Duration(hh:mm:ss)") or "").strip(),
            "connector":    str(g(row, "Connector Type") or "").strip(),
        })
    return out

def parse_jiobp_south(data_rows, headers):
    h = {str(v).strip(): i for i, v in enumerate(headers)}
    g = lambda row, k: row[h[k]] if k in h and h[k] < len(row) else None
    out = []
    for row in data_rows:
        kwh = sf(g(row, "UNITS_CONSUMED_KWH"))
        if not kwh or kwh <= 0: continue
        cost = sf(g(row, "SALE_OF_ENERGY"))
        tariff = sf(g(row, "TARIFF_RATE"))
        vrn = str(g(row, "VRN") or "").strip()
        out.append({
            "date":         parse_date(g(row, "TRANSACTION_DATE")),
            "partner":      "JioBP",
            "station":      str(g(row, "STATION") or "").strip(),
            "city":         str(g(row, "CITY") or "").strip(),
            "state":        str(g(row, "STATE") or "").strip(),
            "vrn":          vrn if vrn not in ("", "-") else None,
            "make":         str(g(row, "MAKE") or "").strip(),
            "model":        str(g(row, "MODEL") or "").strip(),
            "driver":       str(g(row, "CUSTOMER_NAME") or "").strip(),
            "kwh":          kwh,
            "cost":         cost,
            "cost_per_kwh": tariff if tariff else (round(cost/kwh, 4) if cost and kwh else None),
            "duration":     str(g(row, "SESSION_DURATION") or "").strip(),
            "connector":    str(g(row, "CONNECTOR_TYPE") or "").strip(),
        })
    return out

def parse_statiq(data_rows, headers):
    h = {str(v).strip(): i for i, v in enumerate(headers)}
    g = lambda row, k: row[h[k]] if k in h and h[k] < len(row) else None
    out = []
    for row in data_rows:
        if str(g(row, "Status") or "").lower() != "completed": continue
        kwh = sf(g(row, "Units Consumed (kWh)"))
        if not kwh or kwh <= 0: continue
        cost = sf(g(row, "Net Price"))
        tariff = sf(g(row, "EV Charging Rate (Per Unit)"))
        partner = norm_partner(str(g(row, "CPO/Platform") or ""))
        vrn = str(g(row, "Vehicle Registration Number") or "").strip()
        out.append({
            "date":         parse_date(g(row, "Start Date")),
            "partner":      partner,
            "station":      str(g(row, "Station Name") or "").strip(),
            "city":         str(g(row, "City") or "").strip(),
            "state":        str(g(row, "State") or "").strip(),
            "vrn":          vrn if vrn not in ("", "-") else None,
            "make":         "",
            "model":        str(g(row, "Vehicle Name") or "").strip(),
            "driver":       str(g(row, "User Name") or "").strip(),
            "kwh":          kwh,
            "cost":         cost,
            "cost_per_kwh": tariff if tariff else (round(cost/kwh, 4) if cost and kwh else None),
            "duration":     str(g(row, "Charging Time") or "").strip(),
            "connector":    str(g(row, "Connector Name") or "").strip(),
        })
    return out

def parse_kpi_dashboard(data_rows, headers):
    h = {str(v).strip(): i for i, v in enumerate(headers)}
    g = lambda row, k: row[h[k]] if k in h and h[k] < len(row) else None
    out = []
    for row in data_rows:
        if str(g(row, "Status") or "").lower() != "stopped": continue
        kwh = sf(g(row, "Total Energy Delivered"))
        if not kwh or kwh <= 0: continue
        cost_raw = str(g(row, "Charging Session Cost") or "").strip()
        cost = sf(cost_raw) if cost_raw != "null" else None
        partner = norm_partner(str(g(row, "CPO") or ""))
        out.append({
            "date":         parse_date(g(row, "Start Time")),
            "partner":      partner,
            "station":      str(g(row, "Charging Station") or "").strip(),
            "city":         str(g(row, "City") or "").strip(),
            "state":        None,
            "vrn":          None,
            "make":         "",
            "model":        "",
            "driver":       str(g(row, "Phone") or "").strip(),
            "kwh":          kwh,
            "cost":         cost,
            "cost_per_kwh": round(cost/kwh, 4) if cost and kwh else None,
            "duration":     None,
            "connector":    str(g(row, "Vehicle Type") or "").strip(),
        })
    return out

def parse_tata_summary(data_rows, headers):
    h = {str(v).strip(): i for i, v in enumerate(headers)}
    g = lambda row, k: row[h[k]] if k in h and h[k] < len(row) else None
    total_kwh, total_cost = 0.0, 0.0
    for row in data_rows:
        kwh = sf(g(row, "PUBLIC chargers UNITS(Kwh)")) or 0
        cost = sf(g(row, "PUBLIC chargers Amount")) or 0
        total_kwh += kwh
        total_cost += cost
    return {"partner": "Tata Power", "kwh": round(total_kwh, 2), "cost": round(total_cost, 2),
            "note": "User-level monthly summary — individual sessions not available in this file"}

def load_file(fpath):
    ext = os.path.splitext(fpath)[1].lower()
    if ext in (".xlsx", ".xls"):
        wb = openpyxl.load_workbook(fpath, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows: return None, None, []
        return rows[0], detect_format(rows[0]), rows[1:]
    elif ext == ".csv":
        with open(fpath, encoding="utf-8-sig", errors="replace") as f:
            lines = f.read().splitlines()
        start = 0
        known = {"PROFILE_ID","Transaction ID","TRANSACTION_ID","Booking ID","Phone","Invoice Number"}
        for i, line in enumerate(lines[:10]):
            if any(k in line for k in known):
                start = i; break
        reader = list(csv.reader(lines[start:]))
        if not reader: return None, None, []
        return reader[0], detect_format(reader[0]), reader[1:]
    return None, None, []

def main():
    transactions, summaries, files_log = [], [], []

    for fname in sorted(os.listdir(FOLDER)):
        if fname.startswith(".") or fname in ("generate_dashboard.py", "dashboard.html"):
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext not in (".csv", ".xlsx", ".xls"):
            continue
        fpath = os.path.join(FOLDER, fname)
        headers, fmt, data_rows = load_file(fpath)

        if fmt == "jiobp":
            txns = parse_jiobp(data_rows, headers)
            transactions.extend(txns)
            files_log.append({"file": fname, "format": "JioBP XLSX", "sessions": len(txns)})
        elif fmt == "jiobp_south":
            txns = parse_jiobp_south(data_rows, headers)
            transactions.extend(txns)
            files_log.append({"file": fname, "format": "JioBP South CSV", "sessions": len(txns)})
        elif fmt == "statiq":
            txns = parse_statiq(data_rows, headers)
            transactions.extend(txns)
            files_log.append({"file": fname, "format": "Statiq CSV", "sessions": len(txns)})
        elif fmt == "kpi_dashboard":
            txns = parse_kpi_dashboard(data_rows, headers)
            transactions.extend(txns)
            files_log.append({"file": fname, "format": "KPI Dashboard CSV", "sessions": len(txns)})
        elif fmt == "tata_summary":
            summ = parse_tata_summary(data_rows, headers)
            summaries.append(summ)
            files_log.append({"file": fname, "format": "Tata Summary CSV", "sessions": 0, "note": "Monthly summary only"})
        else:
            files_log.append({"file": fname, "format": "Unrecognised", "sessions": 0})

    print(f"\nFiles processed:")
    for f in files_log:
        note = f" ({f.get('note','')})" if f.get('note') else ""
        print(f"  {f['file']}: {f['format']} → {f['sessions']} sessions{note}")

    with_cost    = [t for t in transactions if t["cost"] is not None and t["cost"] > 0]
    no_cost      = [t for t in transactions if t["cost"] is None]
    total_kwh    = sum(t["kwh"] for t in transactions)
    total_spend  = sum(t["cost"] for t in with_cost)
    print(f"\nTotal sessions: {len(transactions)} | kWh: {total_kwh:.1f} | Spend excl. GST: ₹{total_spend:,.0f}")
    print(f"Sessions with cost data: {len(with_cost)} | Without: {len(no_cost)}")

    payload = {
        "transactions": transactions,
        "summaries": summaries,
        "files": files_log,
        "partner_colors": PARTNER_COLORS,
        "generated_at": datetime.now().strftime("%d %b %Y %H:%M"),
    }

    html = build_html(payload)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n✓  Dashboard → {OUTPUT_FILE}")
    print("   Open in Chrome to view.")

# ── HTML ──────────────────────────────────────────────────────────────────────

def build_html(payload):
    data_json = json.dumps(payload, ensure_ascii=False, default=str)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MoEVing | EV Charging Cost Dashboard</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  body {{ font-family: system-ui, -apple-system, sans-serif; }}
  select {{ appearance:none; background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20' fill='%236b7280'%3E%3Cpath fill-rule='evenodd' d='M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z'/%3E%3C/svg%3E"); background-repeat:no-repeat; background-position:right 8px center; background-size:16px; padding-right:2rem; }}
  .chart-box {{ position:relative; }}
  #filterBar {{ position:sticky; top:0; z-index:20; }}
  th {{ cursor:pointer; user-select:none; white-space:nowrap; }}
  th:hover {{ background:#e2e8f0; }}
  .sort-asc::after {{ content:' ↑'; color:#2563eb; }}
  .sort-desc::after {{ content:' ↓'; color:#2563eb; }}
  .dot {{ display:inline-block; width:10px; height:10px; border-radius:50%; margin-right:5px; }}
  .no-cost {{ background:#fef9c3; }}
  .zero-cost {{ background:#f0fdf4; }}
</style>
</head>
<body class="bg-slate-100 min-h-screen text-slate-800">

<!-- HEADER -->
<header class="bg-slate-800 text-white px-6 py-4 flex items-center justify-between shadow-lg">
  <div>
    <div class="text-xs text-slate-400 uppercase tracking-widest font-semibold">MoEVing Urban Technologies</div>
    <h1 class="text-xl font-bold mt-0.5">EV Fast Charging — Cost Intelligence Dashboard</h1>
  </div>
  <div class="text-right">
    <div class="text-xs text-slate-400" id="genTime"></div>
    <div class="text-xs text-slate-300 mt-1" id="dataRange"></div>
  </div>
</header>

<!-- FILTER BAR -->
<div id="filterBar" class="bg-white border-b border-slate-200 shadow-sm px-4 py-2.5">
  <div class="flex flex-wrap items-center gap-2 text-sm">
    <span class="text-xs font-bold text-slate-400 uppercase tracking-wide mr-1">Filter</span>
    <select id="fPartner"  onchange="applyFilters()" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"></select>
    <select id="fVehicle"  onchange="applyFilters()" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"></select>
    <select id="fCity"     onchange="applyFilters()" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"></select>
    <select id="fState"    onchange="applyFilters()" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"></select>
    <select id="fCustomer" onchange="applyFilters()" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none opacity-50 cursor-not-allowed" disabled></select>
    <input type="date" id="fFrom" onchange="applyFilters()" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
    <span class="text-slate-400 text-xs">–</span>
    <input type="date" id="fTo"   onchange="applyFilters()" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
    <button onclick="resetFilters()" class="ml-auto text-xs text-blue-600 hover:text-blue-800 font-semibold underline">Reset</button>
    <button onclick="exportCSV()"   class="text-xs border border-slate-300 rounded-lg px-3 py-1.5 hover:bg-slate-100 font-medium">Export CSV</button>
  </div>
</div>

<!-- KPI CARDS -->
<div class="px-4 pt-4 pb-2 grid grid-cols-2 lg:grid-cols-5 gap-3">
  <div class="bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <div class="text-xs text-slate-400 font-semibold uppercase tracking-wide">Sessions</div>
    <div class="text-3xl font-bold text-slate-800 mt-1" id="kSessions">—</div>
    <div class="text-xs text-slate-400 mt-0.5">Charging events</div>
  </div>
  <div class="bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <div class="text-xs text-slate-400 font-semibold uppercase tracking-wide">Energy</div>
    <div class="text-3xl font-bold text-slate-800 mt-1" id="kKwh">—</div>
    <div class="text-xs text-slate-400 mt-0.5">Total kWh</div>
  </div>
  <div class="bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <div class="text-xs text-slate-400 font-semibold uppercase tracking-wide">Total Spend</div>
    <div class="text-3xl font-bold text-green-600 mt-1" id="kSpend">—</div>
    <div class="text-xs text-slate-400 mt-0.5">Excl. GST</div>
  </div>
  <div class="bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <div class="text-xs text-slate-400 font-semibold uppercase tracking-wide">Avg ₹/kWh</div>
    <div class="text-3xl font-bold text-blue-600 mt-1" id="kRate">—</div>
    <div class="text-xs text-slate-400 mt-0.5">Blended · excl. GST</div>
  </div>
  <div class="bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <div class="text-xs text-slate-400 font-semibold uppercase tracking-wide">Unknown Cost</div>
    <div class="text-3xl font-bold text-amber-500 mt-1" id="kNoCost">—</div>
    <div class="text-xs text-slate-400 mt-0.5">Sessions w/o cost data</div>
  </div>
</div>

<!-- TATA POWER SUMMARY BANNER (conditional) -->
<div id="tataBanner" class="hidden mx-4 mb-2 bg-amber-50 border border-amber-200 rounded-xl p-3 text-sm text-amber-800 flex items-center gap-3">
  <span class="text-lg">⚡</span>
  <span id="tataText"></span>
</div>

<!-- ROW 1: Cost/kWh by Partner + Daily Spend -->
<div class="px-4 pb-3 grid grid-cols-1 lg:grid-cols-5 gap-3">
  <div class="lg:col-span-3 bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <h3 class="text-sm font-semibold text-slate-700">Cost per kWh by Charging Partner</h3>
    <p class="text-xs text-slate-400 mb-3">Average effective rate (excl. GST) — sorted lowest → highest. Sessions without cost excluded.</p>
    <div class="chart-box" style="height:260px"><canvas id="cPartner"></canvas></div>
  </div>
  <div class="lg:col-span-2 bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <h3 class="text-sm font-semibold text-slate-700">Daily Spend Trend</h3>
    <p class="text-xs text-slate-400 mb-3">Total cost per day, excl. GST</p>
    <div class="chart-box" style="height:260px"><canvas id="cDaily"></canvas></div>
  </div>
</div>

<!-- ROW 2: Cost/kWh by Station + Vehicle Cost -->
<div class="px-4 pb-3 grid grid-cols-1 lg:grid-cols-5 gap-3">
  <div class="lg:col-span-3 bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <h3 class="text-sm font-semibold text-slate-700">Cost per kWh by Station</h3>
    <p class="text-xs text-slate-400 mb-3">Colour = charging partner · Lower bars = cheaper · Targets for cost negotiation</p>
    <div class="chart-box" style="height:360px"><canvas id="cStation"></canvas></div>
  </div>
  <div class="lg:col-span-2 bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <h3 class="text-sm font-semibold text-slate-700">Spend by City</h3>
    <p class="text-xs text-slate-400 mb-3">Total excl. GST · top 12</p>
    <div class="chart-box" style="height:360px"><canvas id="cCity"></canvas></div>
  </div>
</div>

<!-- ROW 3: Vehicle + Partner Table -->
<div class="px-4 pb-3 grid grid-cols-1 lg:grid-cols-5 gap-3">
  <div class="lg:col-span-2 bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <h3 class="text-sm font-semibold text-slate-700">Cost per Vehicle</h3>
    <p class="text-xs text-slate-400 mb-3">Total & avg per session, excl. GST</p>
    <div class="chart-box" style="height:240px"><canvas id="cVehicle"></canvas></div>
  </div>
  <div class="lg:col-span-3 bg-white rounded-xl shadow-sm p-4 border border-slate-100">
    <h3 class="text-sm font-semibold text-slate-700">Partner Rate Summary</h3>
    <p class="text-xs text-slate-400 mb-3">Min · Max · Avg rate and volume per charging partner — use for cost negotiations</p>
    <div class="overflow-x-auto">
      <table class="w-full text-xs">
        <thead><tr class="bg-slate-50 text-slate-500 uppercase tracking-wide text-left">
          <th class="px-2 py-2">Partner</th>
          <th class="px-2 py-2 text-right">Sessions</th>
          <th class="px-2 py-2 text-right">kWh</th>
          <th class="px-2 py-2 text-right">Avg ₹/kWh</th>
          <th class="px-2 py-2 text-right">Min ₹/kWh</th>
          <th class="px-2 py-2 text-right">Max ₹/kWh</th>
          <th class="px-2 py-2 text-right">Total Spend</th>
        </tr></thead>
        <tbody id="partnerTable" class="divide-y divide-slate-100"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TRANSACTION TABLE -->
<div class="px-4 pb-8">
  <div class="bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden">
    <div class="p-4 border-b border-slate-100 flex items-center justify-between">
      <div>
        <h3 class="text-sm font-semibold text-slate-700">Transaction Detail</h3>
        <p class="text-xs text-slate-400" id="tableCount"></p>
      </div>
      <div class="flex gap-2 text-xs">
        <span class="inline-flex items-center gap-1"><span class="w-3 h-3 rounded bg-yellow-100 border border-yellow-300"></span> No cost data</span>
        <span class="inline-flex items-center gap-1 ml-2"><span class="w-3 h-3 rounded bg-green-50 border border-green-200"></span> Free / own charger</span>
      </div>
    </div>
    <div class="overflow-x-auto max-h-96 overflow-y-auto">
      <table class="w-full text-xs">
        <thead class="sticky top-0">
          <tr class="bg-slate-50 text-slate-500 font-semibold uppercase tracking-wide text-left">
            <th class="px-3 py-2" onclick="sortBy('date')">Date</th>
            <th class="px-3 py-2" onclick="sortBy('partner')">Partner</th>
            <th class="px-3 py-2" onclick="sortBy('station')">Station</th>
            <th class="px-3 py-2" onclick="sortBy('city')">City</th>
            <th class="px-3 py-2" onclick="sortBy('vrn')">Vehicle</th>
            <th class="px-3 py-2 text-right" onclick="sortBy('kwh')">kWh</th>
            <th class="px-3 py-2 text-right" onclick="sortBy('cost_per_kwh')">₹/kWh</th>
            <th class="px-3 py-2 text-right" onclick="sortBy('cost')">Cost (excl. GST)</th>
            <th class="px-3 py-2" onclick="sortBy('duration')">Duration</th>
          </tr>
        </thead>
        <tbody id="txTable" class="divide-y divide-slate-100"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- DATA SOURCES -->
<div class="px-4 pb-6">
  <div class="bg-white rounded-xl shadow-sm border border-slate-100 p-4">
    <h3 class="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Data Sources Loaded</h3>
    <div id="sourcesList" class="text-xs text-slate-500 space-y-1"></div>
  </div>
</div>

<script>
const PAYLOAD = {data_json};

// ── State ────────────────────────────────────────────────────────────────────
let ALL = PAYLOAD.transactions;
let FILTERED = [];
const PC = PAYLOAD.partner_colors;
let sortState = {{col:'date', dir:'asc'}};

// ── Helpers ──────────────────────────────────────────────────────────────────
const inr = v => '₹' + (v||0).toLocaleString('en-IN', {{maximumFractionDigits:0}});
const fmt2 = v => v == null ? '—' : v.toFixed(2);
const partnerColor = p => PC[p] || '#94a3b8';
const CITY_COLORS = ['#3b82f6','#7c3aed','#d97706','#dc2626','#16a34a','#ea580c','#0891b2','#db2777','#84cc16','#0d9488','#6366f1','#64748b'];

function uniq(arr) {{ return [...new Set(arr.filter(Boolean))].sort(); }}

const charts = {{}};
function destroyChart(id) {{ if(charts[id]){{charts[id].destroy(); delete charts[id];}} }}

// ── Init ─────────────────────────────────────────────────────────────────────
window.onload = function() {{
  document.getElementById('genTime').textContent = 'Generated ' + PAYLOAD.generated_at;

  const dates = ALL.map(r=>r.date).filter(Boolean).sort();
  if(dates.length) {{
    document.getElementById('dataRange').textContent = dates[0] + ' → ' + dates[dates.length-1];
    document.getElementById('fFrom').value = dates[0];
    document.getElementById('fTo').value = dates[dates.length-1];
  }}

  // Populate filters
  fillSelect('fPartner', uniq(ALL.map(r=>r.partner)), 'All Partners');
  fillSelect('fVehicle', uniq(ALL.map(r=>r.vrn).filter(v=>v&&v!=='null')), 'All Vehicles');
  fillSelect('fCity', uniq(ALL.map(r=>r.city)), 'All Cities');
  fillSelect('fState', uniq(ALL.map(r=>r.state).filter(Boolean)), 'All States/Regions');

  // Customer mapping note
  document.getElementById('fCustomer').innerHTML = '<option value="all">Customer (load mapping)</option>';

  // Tata banner
  if(PAYLOAD.summaries && PAYLOAD.summaries.length) {{
    const tataBanner = document.getElementById('tataBanner');
    const parts = PAYLOAD.summaries.map(s =>
      `${{s.partner}}: ${{s.kwh.toFixed(1)}} kWh · ${{inr(s.cost)}} excl. GST (${{s.note}})`);
    document.getElementById('tataText').textContent = parts.join(' | ');
    tataBanner.classList.remove('hidden');
  }}

  // Data sources
  const srcEl = document.getElementById('sourcesList');
  PAYLOAD.files.forEach(f => {{
    const note = f.note ? ` — ${{f.note}}` : '';
    srcEl.innerHTML += `<div>📄 <strong>${{f.file}}</strong> · ${{f.format}} · ${{f.sessions}} sessions${{note}}</div>`;
  }});

  applyFilters();
}};

function fillSelect(id, items, placeholder) {{
  const sel = document.getElementById(id);
  sel.innerHTML = `<option value="all">${{placeholder}}</option>`;
  items.forEach(v => sel.innerHTML += `<option value="${{v}}">${{v}}</option>`);
}}

// ── Filters ──────────────────────────────────────────────────────────────────
function applyFilters() {{
  const partner = document.getElementById('fPartner').value;
  const vehicle = document.getElementById('fVehicle').value;
  const city    = document.getElementById('fCity').value;
  const state   = document.getElementById('fState').value;
  const from    = document.getElementById('fFrom').value;
  const to      = document.getElementById('fTo').value;

  FILTERED = ALL.filter(r => {{
    if(partner !== 'all' && r.partner !== partner) return false;
    if(vehicle !== 'all' && r.vrn !== vehicle) return false;
    if(city    !== 'all' && r.city !== city)   return false;
    if(state   !== 'all' && r.state !== state) return false;
    if(from && r.date < from) return false;
    if(to   && r.date > to)   return false;
    return true;
  }});

  updateKPIs();
  updateAllCharts();
  renderTable();
}}

function resetFilters() {{
  ['fPartner','fVehicle','fCity','fState'].forEach(id => document.getElementById(id).value='all');
  const dates = ALL.map(r=>r.date).filter(Boolean).sort();
  if(dates.length) {{
    document.getElementById('fFrom').value = dates[0];
    document.getElementById('fTo').value   = dates[dates.length-1];
  }}
  applyFilters();
}}

// ── KPIs ─────────────────────────────────────────────────────────────────────
function updateKPIs() {{
  const withCost = FILTERED.filter(r => r.cost != null && r.cost > 0);
  const sessions  = FILTERED.length;
  const totalKwh  = FILTERED.reduce((s,r)=>s+r.kwh, 0);
  const totalSpend= withCost.reduce((s,r)=>s+r.cost, 0);
  const totalKwhC = withCost.reduce((s,r)=>s+r.kwh, 0);
  const avgRate   = totalKwhC > 0 ? totalSpend/totalKwhC : 0;
  const noCost    = FILTERED.filter(r => r.cost == null).length;

  document.getElementById('kSessions').textContent = sessions.toLocaleString('en-IN');
  document.getElementById('kKwh').textContent      = totalKwh.toLocaleString('en-IN', {{maximumFractionDigits:1}}) + ' kWh';
  document.getElementById('kSpend').textContent    = inr(totalSpend);
  document.getElementById('kRate').textContent     = avgRate > 0 ? '₹' + avgRate.toFixed(2) : '—';
  document.getElementById('kNoCost').textContent   = noCost;
}}

// ── Charts ────────────────────────────────────────────────────────────────────
Chart.defaults.font.family = 'system-ui, sans-serif';
Chart.defaults.font.size = 11;
Chart.defaults.color = '#64748b';

function updateAllCharts() {{
  buildPartnerChart();
  buildDailyChart();
  buildStationChart();
  buildCityChart();
  buildVehicleChart();
  buildPartnerTable();
}}

function buildPartnerChart() {{
  const map = {{}};
  FILTERED.filter(r=>r.cost>0).forEach(r => {{
    if(!map[r.partner]) map[r.partner] = {{cost:0, kwh:0}};
    map[r.partner].cost += r.cost;
    map[r.partner].kwh  += r.kwh;
  }});
  const entries = Object.entries(map)
    .map(([p,d]) => ({{p, rate: d.kwh>0 ? d.cost/d.kwh : 0}}))
    .filter(e=>e.rate>0).sort((a,b)=>a.rate-b.rate);

  destroyChart('partner');
  const ctx = document.getElementById('cPartner').getContext('2d');
  charts['partner'] = new Chart(ctx, {{
    type:'bar',
    data: {{
      labels: entries.map(e=>e.p),
      datasets: [{{ label:'₹/kWh', data:entries.map(e=>e.rate),
        backgroundColor:entries.map(e=>partnerColor(e.p)), borderRadius:5, barThickness:26 }}]
    }},
    options: {{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}},
        tooltip:{{ callbacks:{{ label:c=>' ₹'+c.parsed.x.toFixed(2)+'/kWh' }} }} }},
      scales:{{
        x:{{ beginAtZero:false, min:10, title:{{display:true,text:'₹ per kWh (excl. GST)'}}, grid:{{color:'#f1f5f9'}} }},
        y:{{ grid:{{display:false}} }}
      }}
    }}
  }});
}}

function buildDailyChart() {{
  const map = {{}};
  FILTERED.filter(r=>r.cost>0&&r.date).forEach(r => {{
    map[r.date] = (map[r.date]||0) + r.cost;
  }});
  const dates = Object.keys(map).sort();
  const labels = dates.map(d => {{const p=d.split('-'); return p[2]+'/'+p[1];}});

  destroyChart('daily');
  const ctx = document.getElementById('cDaily').getContext('2d');
  charts['daily'] = new Chart(ctx, {{
    type:'line',
    data:{{ labels, datasets:[{{
      label:'Daily Spend (₹)', data:dates.map(d=>map[d]),
      fill:true, tension:0.3,
      borderColor:'#2563eb', backgroundColor:'rgba(37,99,235,0.07)',
      pointBackgroundColor:'#2563eb', pointRadius:3
    }}] }},
    options:{{
      responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}}, tooltip:{{ callbacks:{{ label:c=>' '+inr(c.parsed.y) }} }} }},
      scales:{{
        y:{{ beginAtZero:true, ticks:{{callback:v=>'₹'+v.toLocaleString('en-IN')}}, grid:{{color:'#f1f5f9'}} }},
        x:{{ grid:{{display:false}}, ticks:{{maxRotation:45}} }}
      }}
    }}
  }});
}}

function buildStationChart() {{
  const map = {{}};
  FILTERED.filter(r=>r.cost>0).forEach(r => {{
    const key = r.station || 'Unknown';
    if(!map[key]) map[key] = {{cost:0, kwh:0, partner:r.partner}};
    map[key].cost += r.cost; map[key].kwh += r.kwh;
  }});
  const entries = Object.entries(map)
    .map(([name,d]) => ({{name, rate:d.kwh>0?d.cost/d.kwh:0, partner:d.partner}}))
    .filter(e=>e.rate>0).sort((a,b)=>a.rate-b.rate).slice(0,25);

  destroyChart('station');
  const ctx = document.getElementById('cStation').getContext('2d');
  charts['station'] = new Chart(ctx, {{
    type:'bar',
    data:{{
      labels:entries.map(e=>e.name.length>35?e.name.slice(0,33)+'…':e.name),
      datasets:[{{ label:'₹/kWh', data:entries.map(e=>e.rate),
        backgroundColor:entries.map(e=>partnerColor(e.partner)), borderRadius:3, barThickness:14 }}]
    }},
    options:{{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}},
        tooltip:{{ callbacks:{{ label:(c)=>` ₹${{c.parsed.x.toFixed(2)}}/kWh · ${{entries[c.dataIndex].partner}}` }} }} }},
      scales:{{
        x:{{ beginAtZero:false, min:10, title:{{display:true,text:'₹/kWh (excl. GST)'}}, grid:{{color:'#f1f5f9'}} }},
        y:{{ ticks:{{font:{{size:10}}}}, grid:{{display:false}} }}
      }}
    }}
  }});
}}

function buildCityChart() {{
  const map = {{}};
  FILTERED.filter(r=>r.cost>0&&r.city).forEach(r => {{
    map[r.city] = (map[r.city]||0) + r.cost;
  }});
  const entries = Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,12);

  destroyChart('city');
  const ctx = document.getElementById('cCity').getContext('2d');
  charts['city'] = new Chart(ctx, {{
    type:'doughnut',
    data:{{
      labels:entries.map(e=>e[0]),
      datasets:[{{ data:entries.map(e=>e[1]),
        backgroundColor:CITY_COLORS.slice(0,entries.length), borderWidth:2, borderColor:'#fff' }}]
    }},
    options:{{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'right', labels:{{boxWidth:10, padding:5, font:{{size:10}}}}}},
        tooltip:{{ callbacks:{{ label:c=>` ${{inr(c.parsed)}}` }} }}
      }}
    }}
  }});
}}

function buildVehicleChart() {{
  const map = {{}};
  FILTERED.filter(r=>r.cost>0&&r.vrn).forEach(r => {{
    if(!map[r.vrn]) map[r.vrn] = {{cost:0, kwh:0, sessions:0, label:r.vrn+(r.model?' ('+r.model+')':'')}};
    map[r.vrn].cost += r.cost; map[r.vrn].kwh += r.kwh; map[r.vrn].sessions++;
  }});
  const entries = Object.entries(map).sort((a,b)=>b[1].cost-a[1].cost);
  if(!entries.length) {{ destroyChart('vehicle'); return; }}

  destroyChart('vehicle');
  const ctx = document.getElementById('cVehicle').getContext('2d');
  charts['vehicle'] = new Chart(ctx, {{
    type:'bar',
    data:{{
      labels:entries.map(e=>e[0]),
      datasets:[
        {{label:'Total Spend (₹)', data:entries.map(e=>e[1].cost),
          backgroundColor:CITY_COLORS.slice(0,entries.length), borderRadius:4, yAxisID:'y'}},
        {{label:'Avg/session (₹)', data:entries.map(e=>e[1].sessions?e[1].cost/e[1].sessions:0),
          backgroundColor:CITY_COLORS.slice(0,entries.length).map(c=>c+'55'), borderRadius:4, yAxisID:'y1'}}
      ]
    }},
    options:{{
      responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{position:'bottom',labels:{{boxWidth:10,padding:6}}}},
        tooltip:{{ callbacks:{{ label:c=>` ${{c.dataset.label}}: ${{inr(c.parsed.y)}}` }} }} }},
      scales:{{
        y:{{beginAtZero:true, position:'left', ticks:{{callback:v=>'₹'+v.toLocaleString('en-IN')}}, grid:{{color:'#f1f5f9'}}}},
        y1:{{beginAtZero:true, position:'right', grid:{{drawOnChartArea:false}}, ticks:{{callback:v=>'₹'+v.toLocaleString('en-IN')}}}},
        x:{{grid:{{display:false}}}}
      }}
    }}
  }});
}}

function buildPartnerTable() {{
  const map = {{}};
  FILTERED.filter(r=>r.cost>0&&r.cost_per_kwh).forEach(r => {{
    if(!map[r.partner]) map[r.partner] = {{cost:0, kwh:0, sessions:0, rates:[]}};
    map[r.partner].cost += r.cost;
    map[r.partner].kwh  += r.kwh;
    map[r.partner].sessions++;
    map[r.partner].rates.push(r.cost_per_kwh);
  }});
  const entries = Object.entries(map).sort((a,b)=>{{
    const ar = a[1].kwh>0?a[1].cost/a[1].kwh:99;
    const br = b[1].kwh>0?b[1].cost/b[1].kwh:99;
    return ar-br;
  }});
  const tbody = document.getElementById('partnerTable');
  tbody.innerHTML = entries.map(([p,d]) => {{
    const avg = d.kwh>0?d.cost/d.kwh:0;
    const min = Math.min(...d.rates); const max = Math.max(...d.rates);
    const dot = `<span class="dot" style="background:${{partnerColor(p)}}"></span>`;
    return `<tr class="hover:bg-slate-50">
      <td class="px-2 py-2 font-medium">${{dot}}${{p}}</td>
      <td class="px-2 py-2 text-right">${{d.sessions}}</td>
      <td class="px-2 py-2 text-right">${{d.kwh.toFixed(1)}}</td>
      <td class="px-2 py-2 text-right font-semibold text-blue-600">₹${{avg.toFixed(2)}}</td>
      <td class="px-2 py-2 text-right text-green-600">₹${{min.toFixed(2)}}</td>
      <td class="px-2 py-2 text-right text-red-500">₹${{max.toFixed(2)}}</td>
      <td class="px-2 py-2 text-right font-medium">${{inr(d.cost)}}</td>
    </tr>`;
  }}).join('');
}}

// ── Table ─────────────────────────────────────────────────────────────────────
function renderTable() {{
  const sorted = [...FILTERED].sort((a,b)=>{{
    const dir = sortState.dir==='asc'?1:-1;
    const va=a[sortState.col], vb=b[sortState.col];
    if(va==null) return 1; if(vb==null) return -1;
    return (typeof va==='number'?va-vb:String(va).localeCompare(String(vb)))*dir;
  }});
  document.getElementById('tableCount').textContent = sorted.length.toLocaleString('en-IN') + ' transactions';
  const dot = p => `<span class="dot" style="background:${{partnerColor(p)}}"></span>`;
  document.getElementById('txTable').innerHTML = sorted.map(r => {{
    const rowClass = r.cost==null?'no-cost':(r.cost===0?'zero-cost':'');
    const costCell = r.cost==null?'<span class="text-amber-500 font-medium">No data</span>':
                     r.cost===0 ?'<span class="text-green-600">₹0 (own)</span>':inr(r.cost);
    const rateCell = r.cost_per_kwh ? '₹'+r.cost_per_kwh.toFixed(2) : '—';
    return `<tr class="${{rowClass}} hover:opacity-90 transition-opacity">
      <td class="px-3 py-1.5 text-slate-500">${{r.date||'—'}}</td>
      <td class="px-3 py-1.5">${{dot(r.partner)}}${{r.partner}}</td>
      <td class="px-3 py-1.5 max-w-[160px] truncate text-slate-600" title="${{r.station}}">${{r.station||'—'}}</td>
      <td class="px-3 py-1.5 text-slate-500">${{r.city||'—'}}</td>
      <td class="px-3 py-1.5 font-mono text-slate-600">${{r.vrn||'—'}}</td>
      <td class="px-3 py-1.5 text-right">${{r.kwh.toFixed(2)}}</td>
      <td class="px-3 py-1.5 text-right font-medium">${{rateCell}}</td>
      <td class="px-3 py-1.5 text-right font-semibold">${{costCell}}</td>
      <td class="px-3 py-1.5 text-slate-400">${{r.duration||'—'}}</td>
    </tr>`;
  }}).join('');
}}

function sortBy(col) {{
  sortState.dir = sortState.col===col ? (sortState.dir==='asc'?'desc':'asc') : 'asc';
  sortState.col = col;
  // Update header indicators
  document.querySelectorAll('th').forEach(th => th.classList.remove('sort-asc','sort-desc'));
  renderTable();
}}

// ── Export ────────────────────────────────────────────────────────────────────
function exportCSV() {{
  const cols = ['Date','Partner','Station','City','State','VRN','Model','Driver','kWh','Cost/kWh','Cost excl GST','Duration'];
  const rows = FILTERED.map(r=>[r.date,r.partner,r.station,r.city,r.state||'',r.vrn||'',r.model||'',r.driver||'',r.kwh,r.cost_per_kwh??'',r.cost??'',r.duration||'']);
  const csv = [cols,...rows].map(r=>r.map(v=>`"${{v}}"`).join(',')).join('\\n');
  const a = Object.assign(document.createElement('a'), {{href:URL.createObjectURL(new Blob([csv],{{type:'text/csv'}})), download:'moeving_charging_'+new Date().toISOString().slice(0,10)+'.csv'}});
  a.click();
}}
</script>
</body></html>"""

if __name__ == "__main__":
    main()

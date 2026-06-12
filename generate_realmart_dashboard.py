#!/usr/bin/env python3
"""
Realmart Sales Dashboard Generator
Senior Data Analyst Edition — Dracula Theme + Fully Interactive Visuals
"""

from collections import defaultdict
from datetime import date, datetime, timedelta
from html import escape
from pathlib import Path
from zipfile import ZipFile
import json
import statistics
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "Realmart_Sales_Dataset.xlsx"
OUTPUT = ROOT / "Realmart_Sales_Dashboard.html"
NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def column_index(cell_ref):
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for char in letters:
        index = index * 26 + ord(char.upper()) - 64
    return index - 1


def shared_strings(zip_file):
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    return [
        "".join(text.text or "" for text in item.findall(".//a:t", NS))
        for item in root.findall("a:si", NS)
    ]


def cell_value(cell, shared):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//a:t", NS))
    value = cell.find("a:v", NS)
    if value is None:
        return ""
    raw = value.text or ""
    if cell_type == "s":
        return shared[int(raw)] if raw else ""
    if cell_type == "b":
        return "TRUE" if raw == "1" else "FALSE"
    return raw


def read_xlsx_rows(path):
    with ZipFile(path) as zip_file:
        shared = shared_strings(zip_file)
        workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
        sheet_name = workbook.find(".//a:sheets/a:sheet", NS).attrib["name"]
        root = ET.fromstring(zip_file.read("xl/worksheets/sheet1.xml"))
        sparse_rows = []
        width = 0
        for row in root.findall(".//a:sheetData/a:row", NS):
            values = {}
            for cell in row.findall("a:c", NS):
                idx = column_index(cell.attrib["r"])
                values[idx] = cell_value(cell, shared)
            if values:
                sparse_rows.append(values)
                width = max(width, max(values) + 1)
    rows = [[row.get(i, "") for i in range(width)] for row in sparse_rows]
    headers = rows[0]
    return sheet_name, [dict(zip(headers, row)) for row in rows[1:]]


def as_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def as_date(value):
    if isinstance(value, str) and "/" in value:
        return datetime.strptime(value, "%m/%d/%Y").date()
    try:
        return date(1899, 12, 30) + timedelta(days=float(value))
    except (TypeError, ValueError):
        return None


def as_hour(value):
    try:
        minutes = round((float(value) % 1) * 24 * 60)
        return min(minutes // 60, 23)
    except (TypeError, ValueError):
        return None


def money(value):
    abs_value = abs(value)
    if abs_value >= 1_000_000:
        return f"₹{value / 1_000_000:.2f}M"
    if abs_value >= 1_000:
        return f"₹{value / 1_000:.1f}K"
    return f"₹{value:,.0f}"


def build_dashboard(sheet_name, rows):
    # Enrich rows
    for row in rows:
        row["_total"] = as_float(row["Total"])
        row["_gross"] = as_float(row["gross income"])
        row["_cogs"] = as_float(row["cogs"])
        row["_qty"] = as_float(row["Quantity"])
        row["_rating"] = as_float(row["Rating"])
        row["_date"] = as_date(row["Date"])
        row["_hour"] = as_hour(row["Time"])

    # Global KPIs
    total_sales = sum(row["_total"] for row in rows)
    gross_income = sum(row["_gross"] for row in rows)
    cogs = sum(row["_cogs"] for row in rows)
    quantity = sum(row["_qty"] for row in rows)
    avg_order = total_sales / max(len(rows), 1)
    avg_rating = statistics.mean(row["_rating"] for row in rows) if rows else 0
    margin = (gross_income / total_sales * 100) if total_sales > 0 else 0
    dates = [row["_date"] for row in rows if row["_date"]]
    date_range = f"{min(dates):%d %b %Y} - {max(dates):%d %b %Y}" if dates else ""

    # Prepare compact data for client-side interactivity
    data_rows = []
    for row in rows:
        d = row["_date"]
        data_rows.append({
            "id": row.get("Invoice ID", ""),
            "date": d.strftime("%Y-%m-%d") if d else "",
            "city": row.get("City", ""),
            "branch": row.get("Branch", ""),
            "pline": row.get("Product line", ""),
            "payment": row.get("Payment", ""),
            "ctype": row.get("Customer type", ""),
            "gender": row.get("Gender", ""),
            "total": round(row["_total"], 2),
            "gross": round(row["_gross"], 2),
            "qty": int(row["_qty"]),
            "rating": round(row["_rating"], 2),
            "hour": row["_hour"] if row["_hour"] is not None else -1,
        })

    # Unique values for filters (sorted for nice UX)
    def uniq(key, transform=None):
        s = sorted({r[key] for r in rows if r.get(key)})
        return s

    cities = uniq("City")
    branches = uniq("Branch")
    plines = uniq("Product line")
    payments = uniq("Payment")
    ctypes = uniq("Customer type")
    genders = uniq("Gender")

    # For initial server-side insight text (static)
    city_sales = defaultdict(float)
    for r in rows:
        city_sales[r["City"]] += r["_total"]
    top_city = max(city_sales, key=city_sales.get)

    pline_sales = defaultdict(float)
    for r in rows:
        pline_sales[r["Product line"]] += r["_total"]
    top_pline = max(pline_sales, key=pline_sales.get)

    payment_sales = defaultdict(float)
    for r in rows:
        payment_sales[r["Payment"]] += r["_total"]
    top_pay = max(payment_sales, key=payment_sales.get)

    # Compact JSON for embedding (no escaping issues inside script)
    data_json = json.dumps(data_rows, separators=(",", ":"))
    meta_json = json.dumps({
        "cities": cities,
        "branches": branches,
        "plines": plines,
        "payments": payments,
        "ctypes": ctypes,
        "genders": genders,
        "recordCount": len(rows),
        "dateRange": date_range,
        "sheet": sheet_name,
    }, separators=(",", ":"))

    # Dracula + modern dashboard CSS (beautiful, high contrast, interactive)
    css = """
:root {
  color-scheme: dark;
  --bg: #282a36;
  --bg-2: #21232e;
  --panel: #343746;
  --panel-2: #3a3c4e;
  --border: #44475a;
  --text: #f8f8f2;
  --muted: #6272a4;
  --cyan: #8be9fd;
  --green: #50fa7b;
  --orange: #ffb86c;
  --pink: #ff79c6;
  --purple: #bd93f9;
  --red: #ff5555;
  --yellow: #f1fa8c;
  --accent: #bd93f9;
}

* { box-sizing: border-box; }
body {
  margin: 0;
  background: #282a36;
  color: var(--text);
  font-family: Inter, ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  line-height: 1.5;
}
.app {
  max-width: 1480px;
  margin: 0 auto;
  padding: 28px 24px 60px;
}
header {
  display: flex;
  align-items: flex-end;
  justify-content: space-between;
  gap: 16px;
  padding-bottom: 18px;
  border-bottom: 1px solid var(--border);
  margin-bottom: 18px;
}
h1 {
  margin: 0;
  font-size: clamp(28px, 4.2vw, 42px);
  letter-spacing: -.02em;
  font-weight: 700;
}
.subtitle {
  color: var(--muted);
  font-size: 15px;
  margin-top: 4px;
}
.badge {
  background: #3a3c4e;
  color: var(--purple);
  border: 1px solid var(--border);
  padding: 6px 14px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 600;
  letter-spacing: .5px;
  white-space: nowrap;
}

.controls {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 16px 10px;
  margin-bottom: 18px;
}
.filter-group {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px;
  margin-bottom: 10px;
}
.filter-group:last-child { margin-bottom: 0; }
.filter-label {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: .08em;
  color: var(--muted);
  width: 92px;
  flex-shrink: 0;
  font-weight: 600;
}
.pill {
  background: #21232e;
  color: #e6e6e6;
  border: 1px solid var(--border);
  padding: 5px 11px;
  border-radius: 999px;
  font-size: 12.5px;
  cursor: pointer;
  user-select: none;
  transition: all .1s ease;
  display: inline-flex;
  align-items: center;
  gap: 5px;
}
.pill:hover { border-color: var(--purple); color: #fff; }
.pill.active {
  background: var(--purple);
  color: #282a36;
  border-color: var(--purple);
  font-weight: 600;
}
.pill .count {
  font-size: 10px;
  opacity: .75;
  background: rgba(0,0,0,.18);
  padding: 1px 6px;
  border-radius: 999px;
}
.pill.active .count { background: rgba(255,255,255,.25); color: #282a36; }

.actions {
  display: flex;
  gap: 8px;
  align-items: center;
  flex-wrap: wrap;
}
.btn {
  background: transparent;
  color: var(--text);
  border: 1px solid var(--border);
  padding: 7px 14px;
  border-radius: 8px;
  font-size: 13px;
  cursor: pointer;
  transition: all .1s ease;
}
.btn:hover { border-color: var(--cyan); color: var(--cyan); }
.btn.primary {
  background: var(--purple);
  color: #282a36;
  border-color: var(--purple);
  font-weight: 600;
}
.btn.primary:hover { filter: brightness(1.1); }

.kpis {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(168px, 1fr));
  gap: 12px;
  margin-bottom: 20px;
}
.kpi {
  background: linear-gradient(180deg, #343746 0%, #2f313f 100%);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 16px;
  position: relative;
  overflow: hidden;
}
.kpi::before {
  content: "";
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 3px;
  background: linear-gradient(to right, var(--purple), var(--cyan));
}
.kpi span {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: .07em;
  color: var(--muted);
}
.kpi strong {
  display: block;
  font-size: 22px;
  font-weight: 700;
  margin: 6px 0 2px;
  letter-spacing: -.01em;
}
.kpi small { color: var(--muted); font-size: 12px; }

.viz-grid {
  display: grid;
  grid-template-columns: 2.1fr 1fr;
  gap: 14px;
  margin-bottom: 14px;
}
@media (max-width: 1100px) {
  .viz-grid { grid-template-columns: 1fr; }
}
.panel {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 16px 18px;
  min-width: 0;
}
.panel h3 {
  margin: 0 0 12px;
  font-size: 15px;
  font-weight: 600;
  color: #e6e6e6;
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.panel h3 .sub { font-size: 11px; color: var(--muted); font-weight: 400; }

#trend-chart {
  width: 100%;
  height: 268px;
  display: block;
}

.bars { display: grid; gap: 7px; }
.bar-row {
  display: grid;
  grid-template-columns: 138px 1fr 78px;
  align-items: center;
  gap: 9px;
  font-size: 13px;
  cursor: pointer;
  padding: 2px 0;
  border-radius: 6px;
}
.bar-row:hover { background: rgba(189,147,249,.08); }
.bar-label { color: #e0e0e0; font-size: 13px; padding-left: 2px; }
.bar-track {
  height: 9px;
  background: #21232e;
  border-radius: 999px;
  overflow: hidden;
}
.bar-track span {
  display: block;
  height: 100%;
  background: linear-gradient(90deg, var(--purple), var(--cyan));
  border-radius: 999px;
  transition: width .2s ease;
}
.bar-value { text-align: right; font-variant-numeric: tabular-nums; color: #c9c9c9; font-size: 12.5px; }

.donut-wrap {
  display: flex;
  gap: 22px;
  align-items: center;
  flex-wrap: wrap;
}
.donut {
  width: 168px;
  height: 168px;
  border-radius: 50%;
  position: relative;
  flex-shrink: 0;
}
.donut svg { width: 100%; height: 100%; display: block; transform: rotate(-90deg); }
.legend {
  display: grid;
  gap: 6px;
  font-size: 13px;
  min-width: 130px;
}
.legend-item {
  display: flex;
  align-items: center;
  gap: 8px;
  cursor: pointer;
  padding: 2px 4px;
  border-radius: 6px;
}
.legend-item:hover { background: rgba(189,147,249,.1); }
.legend-swatch {
  width: 13px; height: 13px; border-radius: 3px; flex-shrink: 0;
}
.legend-item.active { color: var(--purple); font-weight: 600; }

.three-col {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 14px;
}
@media (max-width: 920px) { .three-col { grid-template-columns: 1fr 1fr; } }
@media (max-width: 620px) { .three-col { grid-template-columns: 1fr; } }

.hour-grid {
  display: grid;
  grid-template-columns: repeat(12, 1fr);
  gap: 5px;
  height: 138px;
  align-items: end;
  padding: 6px 2px 0;
}
.hour-col {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
}
.hour-bar {
  width: 100%;
  background: linear-gradient(180deg, var(--cyan), var(--purple));
  border-radius: 4px 4px 2px 2px;
  min-height: 6px;
  transition: height .2s ease;
  cursor: pointer;
}
.hour-col small {
  font-size: 9.5px;
  color: var(--muted);
  text-align: center;
}

.insights {
  background: #21232e;
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 14px 16px;
  margin-bottom: 16px;
}
.insights h4 { margin: 0 0 8px; font-size: 13px; color: var(--yellow); }
.insight-list {
  display: grid;
  gap: 5px;
  font-size: 13.2px;
  color: #d8d8d8;
}
.insight-list li { margin-left: 4px; }

.data-table-wrap {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 12px;
  overflow: hidden;
}
.table-toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 14px;
  border-bottom: 1px solid var(--border);
  flex-wrap: wrap;
}
.table-toolbar input {
  background: #21232e;
  border: 1px solid var(--border);
  color: var(--text);
  padding: 7px 11px;
  border-radius: 8px;
  font-size: 13px;
  width: 240px;
}
.table-toolbar input:focus { outline: none; border-color: var(--purple); }
.table-stats { color: var(--muted); font-size: 12.5px; }

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
}
th, td { padding: 9px 11px; text-align: left; border-bottom: 1px solid #3f4150; }
th {
  background: #2c2e3a;
  color: var(--muted);
  font-weight: 600;
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: .06em;
  cursor: pointer;
  user-select: none;
}
th:hover { color: var(--cyan); }
td { color: #e8e8e8; font-variant-numeric: tabular-nums; }
tr:hover td { background: rgba(189,147,249,.05); }
td.right, th.right { text-align: right; }

footer {
  margin-top: 22px;
  font-size: 12px;
  color: #55596b;
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 8px;
}
a, .link { color: var(--cyan); text-decoration: none; }
a:hover { text-decoration: underline; }

.tooltip {
  position: absolute;
  pointer-events: none;
  background: #21232e;
  border: 1px solid var(--border);
  color: #fff;
  padding: 6px 10px;
  border-radius: 6px;
  font-size: 12px;
  box-shadow: 0 6px 18px rgba(0,0,0,.5);
  z-index: 999;
  white-space: nowrap;
}
    """

    # The full interactive HTML document
    html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Realmart • Sales Intelligence Dashboard</title>
  <style>{css}</style>
</head>
<body>
<div class="app">
  <header>
    <div>
      <h1>Realmart</h1>
      <div class="subtitle">Sales Intelligence Dashboard &nbsp;•&nbsp; Senior Data Analyst View &nbsp;•&nbsp; 2019</div>
    </div>
    <div>
      <span class="badge">{escape(date_range)}</span>
    </div>
  </header>

  <!-- FILTERS -->
  <div class="controls" id="controls">
    <div class="filter-group" data-dim="city">
      <div class="filter-label">City</div>
      <div class="pills" id="pills-city"></div>
    </div>
    <div class="filter-group" data-dim="branch">
      <div class="filter-label">Branch</div>
      <div class="pills" id="pills-branch"></div>
    </div>
    <div class="filter-group" data-dim="pline">
      <div class="filter-label">Product Line</div>
      <div class="pills" id="pills-pline"></div>
    </div>
    <div class="filter-group" data-dim="payment">
      <div class="filter-label">Payment</div>
      <div class="pills" id="pills-payment"></div>
    </div>
    <div class="filter-group" data-dim="ctype">
      <div class="filter-label">Customer</div>
      <div class="pills" id="pills-ctype"></div>
    </div>
    <div class="filter-group" data-dim="gender">
      <div class="filter-label">Gender</div>
      <div class="pills" id="pills-gender"></div>
    </div>

    <div class="actions" style="margin-top:8px">
      <button class="btn primary" id="reset-btn">Reset All Filters</button>
      <button class="btn" id="select-all-btn">Select All</button>
      <input id="search-input" type="text" placeholder="Search Invoice ID..." style="margin-left:12px; background:#21232e; border:1px solid #44475a; color:#f8f8f2; padding:7px 11px; border-radius:8px; font-size:13px; width:210px;">
      <span style="color:#55596b; font-size:12px; margin-left:auto">Click pills or bars to filter • Multi-select supported</span>
    </div>
  </div>

  <!-- KPIs -->
  <div class="kpis" id="kpis">
    <div class="kpi"><span>Total Revenue</span><strong id="kpi-revenue">—</strong><small id="kpi-revenue-sub">of filtered sales</small></div>
    <div class="kpi"><span>Gross Income</span><strong id="kpi-gross">—</strong><small id="kpi-margin">margin</small></div>
    <div class="kpi"><span>Transactions</span><strong id="kpi-tx">—</strong><small id="kpi-avg">avg order value</small></div>
    <div class="kpi"><span>Units Sold</span><strong id="kpi-units">—</strong><small id="kpi-units-sub">per invoice</small></div>
    <div class="kpi"><span>Avg Rating</span><strong id="kpi-rating">—</strong><small>customer satisfaction</small></div>
    <div class="kpi"><span>COGS</span><strong id="kpi-cogs">—</strong><small>cost of goods sold</small></div>
  </div>

  <!-- MAIN VIZ -->
  <div class="viz-grid">
    <!-- Trend -->
    <div class="panel">
      <h3>Daily Sales Trend <span class="sub" id="trend-sub"></span></h3>
      <div id="trend-chart" style="position:relative"></div>
    </div>

    <!-- Product Lines (clickable bars) -->
    <div class="panel">
      <h3>Product Line Performance <span class="sub">click to filter</span></h3>
      <div id="product-bars" class="bars"></div>
    </div>
  </div>

  <div class="three-col">
    <!-- City + Branch -->
    <div class="panel">
      <h3>City Performance</h3>
      <div id="city-bars" class="bars"></div>
      <div style="height:10px"></div>
      <h3 style="font-size:13.5px;margin:8px 0 6px">Branch</h3>
      <div id="branch-bars" class="bars"></div>
    </div>

    <!-- Payment Donut -->
    <div class="panel">
      <h3>Payment Mix <span class="sub">click segments</span></h3>
      <div class="donut-wrap">
        <div id="payment-donut" class="donut"></div>
        <div id="payment-legend" class="legend"></div>
      </div>
    </div>

    <!-- Gender + Customer + Hour -->
    <div class="panel">
      <h3 style="margin-bottom:8px">Gender Split</h3>
      <div id="gender-bars" class="bars" style="margin-bottom:14px"></div>

      <h3 style="margin:2px 0 8px">Customer Type</h3>
      <div id="ctype-bars" class="bars" style="margin-bottom:14px"></div>

      <h3 style="margin:2px 0 6px">Hourly Revenue Pattern</h3>
      <div id="hour-grid" class="hour-grid"></div>
    </div>
  </div>

  <!-- Dynamic Analyst Insights -->
  <div class="insights">
    <h4>📊 Senior Analyst Insights (live)</h4>
    <ul class="insight-list" id="insights"></ul>
  </div>

  <!-- Data Table -->
  <div class="data-table-wrap">
    <div class="table-toolbar">
      <div>
        <strong style="font-size:14px">Filtered Transactions</strong>
        <span id="table-stats" class="table-stats"></span>
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        <input id="table-search" type="text" placeholder="Filter visible rows...">
        <button class="btn" id="export-btn">Export CSV</button>
      </div>
    </div>
    <div style="max-height: 380px; overflow:auto">
      <table id="data-table">
        <thead>
          <tr>
            <th data-sort="id">Invoice</th>
            <th data-sort="date">Date</th>
            <th data-sort="city">City</th>
            <th data-sort="pline">Product Line</th>
            <th data-sort="payment">Payment</th>
            <th class="right" data-sort="total">Total (₹)</th>
            <th class="right" data-sort="gross">Gross Income</th>
            <th class="right" data-sort="rating">Rating</th>
          </tr>
        </thead>
        <tbody id="table-body"></tbody>
      </table>
    </div>
  </div>

  <footer>
    <div>
      Source: <strong>{escape(sheet_name)}</strong> • {len(rows):,} records embedded for full interactivity.
      Generated with pure Python + vanilla JS. No external runtime dependencies.
    </div>
    <div>
      Dracula theme • Click any bar / pill / donut segment to instantly filter the entire dashboard
    </div>
  </footer>
</div>

<script>
// ==================== EMBEDDED DATA + META ====================
const DATA = {data_json};
const META = {meta_json};

// ==================== STATE ====================
let activeFilters = {{
  city: new Set(META.cities),
  branch: new Set(META.branches),
  pline: new Set(META.plines),
  payment: new Set(META.payments),
  ctype: new Set(META.ctypes),
  gender: new Set(META.genders),
}};
let searchTerm = "";
let tableSearchTerm = "";
let sortState = {{ key: "date", dir: -1 }}; // default newest first-ish

// ==================== HELPERS ====================
function money(v) {{
  const a = Math.abs(v);
  if (a >= 1e6) return "₹" + (v/1e6).toFixed(2) + "M";
  if (a >= 1e3) return "₹" + (v/1e3).toFixed(1) + "K";
  return "₹" + Math.round(v).toLocaleString();
}}
function pct(a,b) {{ return b ? ((a/b)*100).toFixed(1) + "%" : "0%"; }}
function getFilteredRows() {{
  return DATA.filter(r => {{
    if (searchTerm && !r.id.toLowerCase().includes(searchTerm.toLowerCase())) return false;
    if (!activeFilters.city.has(r.city)) return false;
    if (!activeFilters.branch.has(r.branch)) return false;
    if (!activeFilters.pline.has(r.pline)) return false;
    if (!activeFilters.payment.has(r.payment)) return false;
    if (!activeFilters.ctype.has(r.ctype)) return false;
    if (!activeFilters.gender.has(r.gender)) return false;
    return true;
  }});
}}

// ==================== FILTER PILL RENDER ====================
function renderPills(dim, containerId) {{
  const cont = document.getElementById(containerId);
  cont.innerHTML = "";
  const metaMap = {{ pline: "plines", city: "cities", branch: "branches", payment: "payments", ctype: "ctypes", gender: "genders" }};
  const metaKey = metaMap[dim] || (dim + "s");
  const vals = META[metaKey] || [];
  vals.forEach(val => {{
    const btn = document.createElement("button");
    btn.className = "pill";
    btn.textContent = val;
    const countSpan = document.createElement("span");
    countSpan.className = "count";
    const keyForRow = (dim === "pline") ? "pline" : dim;
    const fullCount = DATA.filter(r => r[keyForRow] === val).length;
    countSpan.textContent = fullCount;
    btn.appendChild(countSpan);

    if (activeFilters[dim].has(val)) btn.classList.add("active");

    btn.onclick = () => {{
      const set = activeFilters[dim];
      if (set.has(val)) {{
        if (set.size > 1) set.delete(val);
      }} else {{
        set.add(val);
      }}
      syncAll();
    }};
    cont.appendChild(btn);
  }});
}}

function renderAllPills() {{
  renderPills("city", "pills-city");
  renderPills("branch", "pills-branch");
  renderPills("pline", "pills-pline");
  renderPills("payment", "pills-payment");
  renderPills("ctype", "pills-ctype");
  renderPills("gender", "pills-gender");
}}

// ==================== KPIs ====================
function updateKPIs(filtered) {{
  const n = filtered.length || 1;
  const rev = filtered.reduce((s,r)=>s+r.total, 0);
  const gr = filtered.reduce((s,r)=>s+r.gross, 0);
  const units = filtered.reduce((s,r)=>s+r.qty, 0);
  const avg = rev / n;
  const avgR = filtered.reduce((s,r)=>s+r.rating,0) / n;

  document.getElementById("kpi-revenue").textContent = money(rev);
  document.getElementById("kpi-gross").textContent = money(gr);
  document.getElementById("kpi-margin").textContent = pct(gr, rev) + " margin";
  document.getElementById("kpi-tx").textContent = filtered.length.toLocaleString();
  document.getElementById("kpi-avg").textContent = "avg " + money(avg);
  document.getElementById("kpi-units").textContent = units.toLocaleString();
  document.getElementById("kpi-units-sub").textContent = (units / n).toFixed(1) + " units / invoice";
  document.getElementById("kpi-rating").textContent = avgR.toFixed(2);
  document.getElementById("kpi-cogs").textContent = money(rev - gr);
}}

// ==================== CHARTS ====================
function renderTrend(filtered) {{
  const container = document.getElementById("trend-chart");
  container.innerHTML = "";

  // Aggregate daily
  const byDay = {{}};
  filtered.forEach(r => {{
    if (!r.date) return;
    byDay[r.date] = (byDay[r.date] || 0) + r.total;
  }});

  const sortedDays = Object.keys(byDay).sort();
  if (!sortedDays.length) {{
    container.innerHTML = "<div style='color:#6272a4;padding:40px;text-align:center'>No data in current filter</div>";
    return;
  }}

  const values = sortedDays.map(d => byDay[d]);
  const maxV = Math.max(...values);
  const minV = Math.min(...values);
  const w = 820, h = 260, padL=48, padR=16, padT=18, padB=28;

  const svgNS = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNS, "svg");
  svg.setAttribute("viewBox", `0 0 ${{w}} ${{h}}`);
  svg.style.width = "100%";
  svg.style.height = "100%";

  // grid + labels
  for (let i=0; i<=4; i++) {{
    const y = padT + (h - padT - padB) * i / 4;
    const line = document.createElementNS(svgNS, "line");
    line.setAttribute("x1", padL);
    line.setAttribute("x2", w - padR);
    line.setAttribute("y1", y);
    line.setAttribute("y2", y);
    line.setAttribute("stroke", "#44475a");
    line.setAttribute("stroke-width", "1");
    svg.appendChild(line);

    const val = maxV - (maxV-minV) * i / 4;
    const txt = document.createElementNS(svgNS, "text");
    txt.setAttribute("x", 6);
    txt.setAttribute("y", y + 4);
    txt.setAttribute("fill", "#6272a4");
    txt.setAttribute("font-size", "11");
    txt.textContent = money(val);
    svg.appendChild(txt);
  }}

  // points
  const pts = [];
  sortedDays.forEach((d, i) => {{
    const x = padL + (w - padL - padR) * i / Math.max(1, sortedDays.length-1);
    const y = padT + (h-padT-padB) * (1 - (byDay[d]-minV) / Math.max(1, maxV-minV));
    pts.push({{x, y, d, v: byDay[d]}});
  }});

  // area
  let areaD = `M ${{pts[0].x}} ${{h-padB}} `;
  pts.forEach(p => areaD += `L ${{p.x}} ${{p.y}} `);
  areaD += `L ${{pts[pts.length-1].x}} ${{h-padB}} Z`;

  const area = document.createElementNS(svgNS, "path");
  area.setAttribute("d", areaD);
  area.setAttribute("fill", "#bd93f9");
  area.setAttribute("fill-opacity", "0.18");
  svg.appendChild(area);

  // line
  let lineD = `M ${{pts[0].x}} ${{pts[0].y}}`;
  pts.forEach(p => lineD += ` L ${{p.x}} ${{p.y}}`);
  const line = document.createElementNS(svgNS, "path");
  line.setAttribute("d", lineD);
  line.setAttribute("fill", "none");
  line.setAttribute("stroke", "#bd93f9");
  line.setAttribute("stroke-width", "2.5");
  line.setAttribute("stroke-linejoin", "round");
  line.setAttribute("stroke-linecap", "round");
  svg.appendChild(line);

  // dots + hover
  const tip = document.createElement("div");
  tip.className = "tooltip";
  tip.style.display = "none";
  container.appendChild(tip);

  pts.forEach((p, idx) => {{
    const c = document.createElementNS(svgNS, "circle");
    c.setAttribute("cx", p.x);
    c.setAttribute("cy", p.y);
    c.setAttribute("r", "3.5");
    c.setAttribute("fill", "#50fa7b");
    c.setAttribute("stroke", "#282a36");
    c.setAttribute("stroke-width", "1.5");
    c.style.cursor = "pointer";

    c.onmouseenter = () => {{
      tip.style.display = "block";
      tip.innerHTML = `<strong>${{p.d}}</strong><br>${{money(p.v)}}`;
      const rect = container.getBoundingClientRect();
      tip.style.left = (p.x + 12) + "px";
      tip.style.top = (p.y - 6) + "px";
    }};
    c.onmouseleave = () => {{ tip.style.display = "none"; }};
    svg.appendChild(c);
  }});

  // x labels (sparse)
  const labelIdxs = [0, Math.floor(pts.length/3), Math.floor(pts.length*2/3), pts.length-1];
  labelIdxs.forEach(i => {{
    if (!pts[i]) return;
    const tx = document.createElementNS(svgNS, "text");
    tx.setAttribute("x", pts[i].x);
    tx.setAttribute("y", h - 6);
    tx.setAttribute("fill", "#6272a4");
    tx.setAttribute("font-size", "10");
    tx.setAttribute("text-anchor", "middle");
    tx.textContent = pts[i].d.slice(5);
    svg.appendChild(tx);
  }});

  container.appendChild(svg);
  document.getElementById("trend-sub").textContent = `${{sortedDays.length}} days • ${{filtered.length}} tx`;
}}

function renderBars(containerId, groups, dimKey, color = "#bd93f9") {{
  const el = document.getElementById(containerId);
  el.innerHTML = "";
  const total = groups.reduce((s,g)=>s+g.value,0) || 1;

  groups.forEach(g => {{
    const row = document.createElement("div");
    row.className = "bar-row";
    row.innerHTML = `
      <div class="bar-label">${{g.label}}</div>
      <div class="bar-track"><span style="width:${{(g.value/total*100).toFixed(1)}}%; background:${{color}}"></span></div>
      <div class="bar-value">${{money(g.value)}}</div>
    `;
    row.onclick = () => {{
      const set = activeFilters[dimKey];
      if (set.has(g.label)) {{
        if (set.size > 1) set.delete(g.label);
      }} else {{
        set.add(g.label);
      }}
      syncAll();
    }};
    el.appendChild(row);
  }});
}}

function renderDonut(filtered) {{
  const wrap = document.getElementById("payment-donut");
  const legend = document.getElementById("payment-legend");
  wrap.innerHTML = "";
  legend.innerHTML = "";

  const payAgg = {{}};
  filtered.forEach(r => payAgg[r.payment] = (payAgg[r.payment]||0) + r.total);
  const entries = Object.entries(payAgg).sort((a,b)=>b[1]-a[1]);
  const total = entries.reduce((s,e)=>s+e[1],0) || 1;

  const colors = ["#bd93f9", "#8be9fd", "#50fa7b", "#ffb86c", "#ff79c6"];

  // SVG donut
  const size = 168;
  const svgNS = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNS, "svg");
  svg.setAttribute("width", size);
  svg.setAttribute("height", size);
  svg.setAttribute("viewBox", "0 0 42 42");

  let offset = 0;
  entries.forEach((e, i) => {{
    const [name, val] = e;
    const frac = val / total;
    const dash = frac * 100;
    const c = document.createElementNS(svgNS, "circle");
    c.setAttribute("cx", "21");
    c.setAttribute("cy", "21");
    c.setAttribute("r", "15.9155");
    c.setAttribute("fill", "none");
    c.setAttribute("stroke", colors[i % colors.length]);
    c.setAttribute("stroke-width", "7");
    c.setAttribute("stroke-dasharray", `${{dash.toFixed(3)}} ${{100-dash}}`);
    c.setAttribute("stroke-dashoffset", (-offset).toFixed(3));
    c.style.cursor = "pointer";

    c.onclick = () => {{
      const set = activeFilters.payment;
      if (set.has(name) && set.size > 1) set.delete(name);
      else set.add(name);
      syncAll();
    }};
    svg.appendChild(c);
    offset += dash;
  }});

  // center hole label
  const centerText = document.createElementNS(svgNS, "text");
  centerText.setAttribute("x","21");
  centerText.setAttribute("y","22.5");
  centerText.setAttribute("text-anchor","middle");
  centerText.setAttribute("fill","#f8f8f2");
  centerText.setAttribute("font-size","5.8");
  centerText.setAttribute("font-weight","600");
  centerText.textContent = "PAY";
  svg.appendChild(centerText);

  wrap.appendChild(svg);

  // legend
  entries.forEach((e, i) => {{
    const [name, val] = e;
    const item = document.createElement("div");
    item.className = "legend-item";
    item.innerHTML = `
      <span class="legend-swatch" style="background:${{colors[i%colors.length]}}"></span>
      <span style="flex:1">${{name}}</span>
      <span style="font-variant-numeric:tabular-nums;color:#888">${{money(val)}} • ${{pct(val,total)}}</span>
    `;
    item.onclick = () => {{
      const set = activeFilters.payment;
      if (set.has(name) && set.size > 1) set.delete(name);
      else set.add(name);
      syncAll();
    }};
    legend.appendChild(item);
  }});
}}

function renderHourly(filtered) {{
  const container = document.getElementById("hour-grid");
  container.innerHTML = "";

  const hours = Array(24).fill(0);
  filtered.forEach(r => {{ if (r.hour >= 0) hours[r.hour] += r.total; }});

  const maxH = Math.max(...hours) || 1;
  for (let h=0; h<24; h+=2) {{ // every other for space
    const col = document.createElement("div");
    col.className = "hour-col";
    const barH = Math.max(6, (hours[h] / maxH) * 100);
    col.innerHTML = `
      <div class="hour-bar" style="height:${{barH}}%" title="${{h.toString().padStart(2,'0')}}:00 • ${{money(hours[h])}}"></div>
      <small>${{h.toString().padStart(2,'0')}}</small>
    `;
    col.onclick = () => {{
      // simple: show only this hour's contribution (we don't have hour filter yet, just highlight insight)
      alert(`Hour ${{h.toString().padStart(2,'0')}}:00 contributed ${{money(hours[h])}} in current filter`);
    }};
    container.appendChild(col);
  }}
}}

function renderSimpleBars(containerId, dimKey, color) {{
  const filtered = getFilteredRows();
  const agg = {{}};
  filtered.forEach(r => {{
    const k = r[dimKey];
    agg[k] = (agg[k] || 0) + r.total;
  }});
  const groups = Object.entries(agg).map(([label, value]) => ({{label, value}})).sort((a,b)=>b.value-a.value);
  renderBars(containerId, groups, dimKey === "pline" ? "pline" : dimKey, color);
}}

function updateAllCharts(filtered) {{
  // Trend
  renderTrend(filtered);

  // Product bars (purple-cyan)
  const plAgg = {{}};
  filtered.forEach(r => plAgg[r.pline] = (plAgg[r.pline]||0) + r.total);
  const plGroups = Object.entries(plAgg).map(([l,v])=>({{label:l,value:v}})).sort((a,b)=>b.value-a.value);
  renderBars("product-bars", plGroups, "pline", "#bd93f9");

  // City
  renderSimpleBars("city-bars", "city", "#8be9fd");
  // Branch
  renderSimpleBars("branch-bars", "branch", "#50fa7b");

  // Donut
  renderDonut(filtered);

  // Gender
  const gAgg = {{}};
  filtered.forEach(r => gAgg[r.gender]=(gAgg[r.gender]||0)+r.total);
  renderBars("gender-bars", Object.entries(gAgg).map(([l,v])=>({{label:l,value:v}})), "gender", "#ff79c6");

  // Customer type
  const ctAgg = {{}};
  filtered.forEach(r => ctAgg[r.ctype]=(ctAgg[r.ctype]||0)+r.total);
  renderBars("ctype-bars", Object.entries(ctAgg).map(([l,v])=>({{label:l,value:v}})), "ctype", "#ffb86c");

  // Hourly
  renderHourly(filtered);
}}

// ==================== INSIGHTS (dynamic senior analyst) ====================
function updateInsights(filtered) {{
  const el = document.getElementById("insights");
  if (!filtered.length) {{
    el.innerHTML = "<li>No records match the current filters.</li>";
    return;
  }}
  const rev = filtered.reduce((s,r)=>s + r.total, 0);
  const n = filtered.length;

  // Top contributors
  const pl = {{}}; filtered.forEach(r=>pl[r.pline]=(pl[r.pline]||0)+r.total);
  const topPl = Object.entries(pl).sort((a,b)=>b[1]-a[1])[0];

  const cty = {{}}; filtered.forEach(r=>cty[r.city]=(cty[r.city]||0)+r.total);
  const topC = Object.entries(cty).sort((a,b)=>b[1]-a[1])[0];

  const pay = {{}}; filtered.forEach(r=>pay[r.payment]=(pay[r.payment]||0)+r.total);
  const topP = Object.entries(pay).sort((a,b)=>b[1]-a[1])[0];

  const avgR = filtered.reduce((s,r)=>s+r.rating,0)/n;

  const cashShare = (pay["Cash"] || 0) / rev * 100;

  const items = [
    `<li><strong>${{topPl[0]}}</strong> leads with ${{money(topPl[1])}} (${{pct(topPl[1], rev)}}) in current selection.</li>`,
    `<li><strong>${{topC[0]}}</strong> accounts for ${{money(topC[1])}} — ${{pct(topC[1], rev)}} of filtered revenue.</li>`,
    `<li>Cash payments represent <strong>${{cashShare.toFixed(1)}}%</strong> of revenue. ${{topP[0]}} is next at ${{pct(topP[1],rev)}}.</li>`,
    `<li>Average customer rating in view: <strong>${{avgR.toFixed(2)}}</strong>. ${{avgR < 6.3 ? "Consider service improvements." : "Solid satisfaction."}}</li>`,
    `<li>${{n.toLocaleString()}} transactions • average basket ${{money(rev/n)}}.</li>`,
  ];
  el.innerHTML = items.join("");
}}

// ==================== DATA TABLE ====================
let visibleRowsCache = [];

function renderTable(filtered) {{
  // apply secondary table search
  let rows = filtered;
  if (tableSearchTerm) {{
    const q = tableSearchTerm.toLowerCase();
    rows = rows.filter(r => 
      r.id.toLowerCase().includes(q) ||
      r.pline.toLowerCase().includes(q) ||
      r.city.toLowerCase().includes(q)
    );
  }}

  // sort
  const key = sortState.key;
  rows = [...rows].sort((a,b) => {{
    let va = a[key], vb = b[key];
    if (key === "total" || key === "gross" || key === "rating") {{
      return (va - vb) * sortState.dir;
    }}
    if (key === "date") {{
      return (va > vb ? 1 : -1) * sortState.dir;
    }}
    return String(va).localeCompare(String(vb)) * sortState.dir;
  }});

  visibleRowsCache = rows;

  const tbody = document.getElementById("table-body");
  tbody.innerHTML = "";

  const show = rows.slice(0, 18); // cap for perf
  show.forEach(r => {{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td style="font-family:monospace;font-size:11.5px">${{r.id}}</td>
      <td>${{r.date}}</td>
      <td>${{r.city}}</td>
      <td>${{r.pline}}</td>
      <td>${{r.payment}}</td>
      <td class="right">${{money(r.total)}}</td>
      <td class="right">${{money(r.gross)}}</td>
      <td class="right" style="color:${{r.rating >= 7 ? '#50fa7b' : r.rating >= 6 ? '#f1fa8c' : '#ffb86c'}}">${{r.rating.toFixed(1)}}</td>
    `;
    tbody.appendChild(tr);
  }});

  document.getElementById("table-stats").textContent = 
    `— showing ${{show.length}} of ${{rows.length}} filtered (total ${{filtered.length}} records)`;
}}

// ==================== MASTER SYNC ====================
function syncAll() {{
  const filtered = getFilteredRows();

  // KPIs + charts
  updateKPIs(filtered);
  updateAllCharts(filtered);
  updateInsights(filtered);

  // Pills (re-render to reflect active state + live counts)
  renderAllPills();

  // Table
  renderTable(filtered);
}}

// ==================== BOOTSTRAP ====================
function initFilters() {{
  // Pills already rendered in renderAllPills
  document.getElementById("reset-btn").onclick = () => {{
    activeFilters.city = new Set(META.cities);
    activeFilters.branch = new Set(META.branches);
    activeFilters.pline = new Set(META.plines);
    activeFilters.payment = new Set(META.payments);
    activeFilters.ctype = new Set(META.ctypes);
    activeFilters.gender = new Set(META.genders);
    searchTerm = "";
    document.getElementById("search-input").value = "";
    tableSearchTerm = "";
    document.getElementById("table-search").value = "";
    syncAll();
  }};

  document.getElementById("select-all-btn").onclick = () => {{
    activeFilters.city = new Set(META.cities);
    activeFilters.branch = new Set(META.branches);
    activeFilters.pline = new Set(META.plines);
    activeFilters.payment = new Set(META.payments);
    activeFilters.ctype = new Set(META.ctypes);
    activeFilters.gender = new Set(META.genders);
    syncAll();
  }};

  const s = document.getElementById("search-input");
  s.oninput = () => {{
    searchTerm = s.value.trim();
    syncAll();
  }};

  const ts = document.getElementById("table-search");
  ts.oninput = () => {{
    tableSearchTerm = ts.value.trim();
    const filtered = getFilteredRows();
    renderTable(filtered);
  }};

  // Table header sorting
  document.querySelectorAll("#data-table th[data-sort]").forEach(th => {{
    th.onclick = () => {{
      const k = th.dataset.sort;
      if (sortState.key === k) {{
        sortState.dir = -sortState.dir;
      }} else {{
        sortState.key = k;
        sortState.dir = (k === "date" || k === "total" || k === "gross") ? -1 : 1;
      }}
      const filtered = getFilteredRows();
      renderTable(filtered);
    }};
  }});

  // Export
  document.getElementById("export-btn").onclick = () => {{
    if (!visibleRowsCache.length) return;
    const headers = ["Invoice","Date","City","Branch","ProductLine","Payment","CustomerType","Gender","Total","GrossIncome","Qty","Rating"];
    let csv = headers.join(",") + "\\n";
    visibleRowsCache.forEach(r => {{
      csv += [
        r.id, r.date, r.city, r.branch || "", r.pline, r.payment, r.ctype, r.gender,
        r.total, r.gross, r.qty, r.rating
      ].map(v => typeof v === "string" && v.includes(",") ? `"${{v}}"` : v).join(",") + "\\n";
    }});
    const blob = new Blob([csv], {{type: "text/csv"}});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "realmart_filtered.csv";
    a.click();
    URL.revokeObjectURL(url);
  }};
}}

function boot() {{
  // Initial pill UI
  renderAllPills();
  initFilters();

  // First full render (all data)
  const initial = getFilteredRows();
  updateKPIs(initial);
  updateAllCharts(initial);
  updateInsights(initial);
  renderTable(initial);

  // Subtle welcome note in console for analysts
  console.log("%c[Realmart] Dracula interactive dashboard ready. All filters are live.", "color:#6272a4");
}}

boot();
</script>
</body>
</html>"""
    return html


def main():
    if not WORKBOOK.exists():
        raise FileNotFoundError(f"Workbook not found: {WORKBOOK}")
    sheet_name, rows = read_xlsx_rows(WORKBOOK)
    html = build_dashboard(sheet_name, rows)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"✅ Wrote beautiful Dracula interactive dashboard → {OUTPUT}")
    print(f"   Records: {len(rows):,}")
    print(f"   Open the .html in any browser for full interactivity.")


if __name__ == "__main__":
    main()

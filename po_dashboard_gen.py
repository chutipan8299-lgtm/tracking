"""
po_business.py  (Scalable edition)
------------------------------------
ออกแบบใหม่รองรับ PO จำนวนมาก:
- แทน bar chart แน่นๆ ด้วย Fill Rate heatmap + scrollable table
- PO Timeline แสดงเป็น scatter/gantt ไม่แน่นเมื่อเพิ่มขึ้น
- กด PO ใน chart → filter table ทันที
"""

import sys, json
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

EXCEL_PATH = sys.argv[1] if len(sys.argv) > 1 else "PO_Items.xlsx"
OUT_PATH   = "business_dashboard.html"

STATUS_COLORS = {
    "Ordered":            "#94a3b8",
    "Stuffed":            "#fbbf24",
    "In Transit":         "#3b82f6",
    "Arrived":            "#8b5cf6",
    "Partially Received": "#f97316",
    "Fully Received":     "#22c55e",
    "Unknown":            "#cbd5e1",
}

# =============================================================================
# LOAD
# =============================================================================
def load_data(path):
    df = pd.read_excel(path)

    # ── พิมพ์ชื่อคอลัมน์จริงออกมาให้เห็น (ใช้ debug ครั้งแรก) ──────────────
    print("\n📋 คอลัมน์จริงในไฟล์:")
    for c in df.columns:
        print(f"   '{c}'")

    # ── Map ชื่อไทย → ชื่ออังกฤษ (ครอบคลุมทั้งชื่อเต็มและชื่อย่อที่อาจต่างกัน) ──
    col_map = {
        # PO Number
        "รหัสใบจัดซื้อสินค้า":          "PO Number",
        "เลขที่ใบสั่งซื้อ":              "PO Number",
        "PO Number":                      "PO Number",
        # Order Date
        "เวลาที่สั่ง":                   "Order Date",
        "วันที่สั่งซื้อ":               "Order Date",
        "Order Date":                     "Order Date",
        # Product SKU
        "รหัสสินค้า":                    "Product SKU",
        "SKU":                            "Product SKU",
        "Product SKU":                    "Product SKU",
        # Product Name
        "ชื่อสินค้า":                    "Product Name",
        "Product Name":                   "Product Name",
        # Ordered Qty
        "จำนวนการจัดซื้อสินค้า":        "Ordered Qty",
        "จำนวนที่สั่ง":                 "Ordered Qty",
        "จำนวนสั่งซื้อ":                "Ordered Qty",
        "Ordered Qty":                    "Ordered Qty",
        # Received Qty
        "จำนวนเข้าคลังสำเร็จ":          "Received Qty",
        "จำนวนที่รับแล้ว":              "Received Qty",
        "จำนวนรับเข้าคลัง":             "Received Qty",
        "Received Qty":                   "Received Qty",
        # Dates
        "วันที่เรือออกจริง":             "Departure Date",
        "วันที่ส่งออก":                  "Departure Date",
        "วันที่เรือถึงจริง":             "Arrival Date",
        "วันที่มาถึง":                   "Arrival Date",
        "วันที่บรรจุสินค้าเข้าตู้":     "Stuffing Date",
        "วันที่ stuffing":               "Stuffing Date",
    }
    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    # ── ถ้ายัง rename ไม่ได้ ให้ match แบบ fuzzy (strip + lower) ─────────────
    remaining = {
        "PO Number":      ["po","order number","เลขpo"],
        "Ordered Qty":    ["order qty","qty order","สั่ง","ordered"],
        "Received Qty":   ["receive","received","รับ","เข้าคลัง"],
        "Product SKU":    ["sku","รหัส"],
        "Order Date":     ["order date","วันสั่ง"],
        "Arrival Date":   ["arrival","ถึง","arrive"],
        "Departure Date": ["depart","ออก"],
        "Stuffing Date":  ["stuff","บรรจุ"],
        "Product Name":   ["product name","ชื่อ"],
    }
    for target, keywords in remaining.items():
        if target not in df.columns:
            for col in df.columns:
                col_low = str(col).strip().lower()
                if any(k in col_low for k in keywords):
                    print(f"   ✅ fuzzy match: '{col}' → '{target}'")
                    df = df.rename(columns={col: target})
                    break

    # ── แจ้งเตือนถ้า column สำคัญยังหาไม่เจอ ────────────────────────────────
    must_have = ["Ordered Qty", "Received Qty", "Product SKU"]
    missing   = [c for c in must_have if c not in df.columns]
    if missing:
        print(f"\n❌ ยังหาคอลัมน์เหล่านี้ไม่เจอ: {missing}")
        print("   → ลองแก้ชื่อใน col_map ให้ตรงกับชื่อจริงในไฟล์ข้างบน\n")
        raise KeyError(f"Column(s) not found: {missing}")

    # ── Ensure required cols exist with defaults ───────────────────────────────
    for col in ["Ordered Qty", "Received Qty"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    for col in ["PO Number"]:
        if col in df.columns:
            df[col] = (pd.to_numeric(df[col], errors="coerce")
                       .fillna(0).astype(int).astype(str))
            df.loc[df[col] == "0", col] = "Unknown"
        else:
            df[col] = "Unknown"

    for col in ["Product SKU","Product Name"]:
        if col not in df.columns:
            df[col] = "—"

    for col in ["Order Date","Departure Date","Arrival Date","Stuffing Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        else:
            df[col] = pd.NaT

    if "Order Date" in df.columns and "Arrival Date" in df.columns:
        df["LT_Total"] = (df["Arrival Date"] - df["Order Date"]).dt.days
    else:
        df["LT_Total"] = np.nan

    def derive_status(row):
        rcv, ord_q = row.get("Received Qty", 0), row.get("Ordered Qty", 0)
        if rcv > 0: return "Fully Received" if rcv >= ord_q else "Partially Received"
        if pd.notna(row.get("Arrival Date")):   return "Arrived"
        if pd.notna(row.get("Departure Date")): return "In Transit"
        if pd.notna(row.get("Stuffing Date")):  return "Stuffed"
        return "Ordered"

    df["Status"]   = df.apply(derive_status, axis=1)
    df["Variance"] = df["Received Qty"] - df["Ordered Qty"]
    return df

# =============================================================================
# SUMMARIES
# =============================================================================
def get_summaries(df):
    po_sum = df.groupby("PO Number").agg(
        Ordered_Qty  =("Ordered Qty",  "sum"),
        Received_Qty =("Received Qty", "sum"),
        Status       =("Status",       "first"),
        Order_Date   =("Order Date",   "min"),
        Arrival_Date =("Arrival Date", "max"),
        LT_Total     =("LT_Total",     "mean"),
    ).reset_index()
    po_sum.columns = ["PO Number","Ordered Qty","Received Qty",
                      "Status","Order Date","Arrival Date","LT_Total"]
    po_sum["Fill Rate"] = (
        po_sum["Received Qty"] / po_sum["Ordered Qty"] * 100
    ).fillna(0).clip(upper=100).round(1)
    po_sum["Missing"] = (po_sum["Ordered Qty"] - po_sum["Received Qty"]).clip(lower=0)

    shortages     = df[df["Variance"] < 0].copy()
    shortages["Missing"] = shortages["Variance"].abs()
    top_shortages = shortages.sort_values("Missing", ascending=False).head(15)

    return po_sum, top_shortages

# =============================================================================
# CHARTS  — scalable design
# =============================================================================
def build_charts(df, po_sum, top_shortages):
    charts = {}

    # ── 1. Status Donut ────────────────────────────────────────────────────────
    status_count = df.groupby("Status")["Product SKU"].count().reset_index()
    fig1 = px.pie(status_count, names="Status", values="Product SKU",
                  color="Status", color_discrete_map=STATUS_COLORS,
                  hole=0.55, title="Inventory Pipeline Status")
    fig1.update_traces(textposition="outside", textinfo="percent+label")
    fig1.update_layout(margin=dict(t=50,b=10,l=10,r=10), height=360,
                       showlegend=False)
    charts["pipeline"] = pio.to_json(fig1)

    # ── 2. Fill Rate Bubble (replaces bar — scales well) ─────────────────────
    #    x=Order Date, y=Fill Rate, size=Ordered Qty, color=Status
    po_plot = po_sum.copy()
    po_plot["Order Date"] = po_plot["Order Date"].dt.strftime("%Y-%m-%d")
    po_plot["size_norm"]  = po_plot["Ordered Qty"].clip(lower=1)

    fig2 = px.scatter(
        po_plot,
        x="Order Date", y="Fill Rate",
        size="size_norm", color="Status",
        color_discrete_map=STATUS_COLORS,
        hover_name="PO Number",
        hover_data={"Ordered Qty":":,", "Received Qty":":,",
                    "Missing":":,", "size_norm":False},
        text="PO Number",
        title="Fill Rate by PO  (size = order qty, click to filter table)",
        height=400,
    )
    fig2.update_traces(textposition="top center", textfont_size=10,
                       marker=dict(line=dict(width=1, color="#fff")))
    fig2.update_yaxes(range=[-5, 110], ticksuffix="%")
    fig2.update_layout(margin=dict(t=50,b=40,l=10,r=10),
                       clickmode="event+select")
    charts["fillrate"] = pio.to_json(fig2)

    # ── 3. Top Shortages — horizontal bar (unchanged, still useful) ───────────
    if not top_shortages.empty:
        fig3 = px.bar(top_shortages, x="Missing", y="Product SKU",
                      orientation="h",
                      title="Top 15 Shortage Risks (Units Missing)",
                      color="Missing",
                      color_continuous_scale=["#fca5a5","#dc2626"],
                      text="Missing")
        fig3.update_layout(margin=dict(t=50,b=10,l=10,r=10), height=420,
                           yaxis={"categoryorder":"total ascending"},
                           coloraxis_showscale=False)
        fig3.update_traces(textposition="outside")
    else:
        fig3 = go.Figure()
        fig3.update_layout(
            title="Top 15 Shortage Risks",
            height=420,
            annotations=[dict(text="No shortages ✅", xref="paper", yref="paper",
                              x=0.5, y=0.5, showarrow=False, font_size=18)]
        )
    charts["shortages"] = pio.to_json(fig3)

    # ── 4. Lead Time trend (scatter — works for any number of POs) ───────────
    lt_data = po_sum.dropna(subset=["LT_Total","Order Date"]).sort_values("Order Date")
    if not lt_data.empty:
        fig4 = px.scatter(lt_data, x="Order Date", y="LT_Total",
                          color="Status", color_discrete_map=STATUS_COLORS,
                          hover_name="PO Number",
                          title="Lead Time Trend (Order → Arrival days)",
                          height=320)
        lt_sorted = lt_data.sort_values("Order Date").dropna(subset=["LT_Total"])
        if len(lt_sorted) >= 3:
            lt_sorted = lt_sorted.copy()
            lt_sorted["trend"] = lt_sorted["LT_Total"].rolling(3, min_periods=1).mean()
            fig4.add_scatter(x=lt_sorted["Order Date"], y=lt_sorted["trend"],
                             mode="lines", name="3-PO avg",
                             line=dict(color="#6366f1", width=2, dash="dot"))
        fig4.update_layout(margin=dict(t=50,b=30,l=10,r=10))
    else:
        fig4 = go.Figure()
        fig4.update_layout(
            title="Lead Time Trend",
            height=320,
            annotations=[dict(text="ยังไม่มีข้อมูลวันที่ครบถ้วน",
                              xref="paper", yref="paper",
                              x=0.5, y=0.5, showarrow=False, font_size=16)]
        )
    charts["leadtime"] = pio.to_json(fig4)

    return charts

# =============================================================================
# MAIN
# =============================================================================
print(f"📊 Generating dashboard from {EXCEL_PATH}...")
df                    = load_data(EXCEL_PATH)
po_sum, top_shortages = get_summaries(df)
charts                = build_charts(df, po_sum, top_shortages)

# Table rows
table_data = (
    df.groupby(["PO Number","Product SKU"])
    .agg(Product_Name=("Product Name","first"), Status=("Status","first"),
         Ordered_Qty=("Ordered Qty","sum"), Received_Qty=("Received Qty","sum"),
         Variance=("Variance","sum"), Order_Date=("Order Date","first"))
    .reset_index()
)

rows_html = ""
for _, r in table_data.iterrows():
    vc   = "danger" if r["Variance"] < 0 else ("success" if r["Variance"] > 0 else "neutral")
    sbg  = STATUS_COLORS.get(r["Status"], "#ccc")
    odt  = r["Order_Date"].strftime("%Y-%m-%d") if pd.notna(r["Order_Date"]) else "—"
    fill = round(r["Received_Qty"]/r["Ordered_Qty"]*100) if r["Ordered_Qty"] else 0
    bar  = f'<div style="width:100%;background:#e5e7eb;border-radius:4px;height:6px"><div style="width:{min(fill,100)}%;background:{"#22c55e" if fill>=100 else "#3b82f6" if fill>0 else "#e5e7eb"};height:6px;border-radius:4px"></div></div><span style="font-size:10px;color:#6b7280">{fill}%</span>'
    rows_html += (
        f'<tr data-po="{r["PO Number"]}" data-status="{r["Status"]}" '
        f'data-shortage="{"yes" if r["Variance"]<0 else "no"}">'
        f'<td><b>{r["PO Number"]}</b></td>'
        f'<td style="font-family:monospace;font-size:12px">{r["Product SKU"]}</td>'
        f'<td><span class="badge" style="background:{sbg}">{r["Status"]}</span></td>'
        f'<td class="num">{int(r["Ordered_Qty"]):,}</td>'
        f'<td class="num">{int(r["Received_Qty"]):,}</td>'
        f'<td class="num {"tdanger" if r["Variance"]<0 else "tsuccess"}">{int(r["Variance"]):+,}</td>'
        f'<td>{bar}</td>'
        f'<td>{odt}</td>'
        f'</tr>'
    )

total_ord = df["Ordered Qty"].sum()
total_rec = df["Received Qty"].sum()
stats = {
    "fill_rate":      f"{total_rec/total_ord*100:.1f}%" if total_ord else "0%",
    "shortage_count": int((df["Variance"] < 0).sum()),
    "in_transit":     int((df["Status"] == "In Transit").sum()),
    "total_pos":      df["PO Number"].nunique(),
    "total_ordered":  f"{int(total_ord):,}",
    "total_received": f"{int(total_rec):,}",
}
status_opts = "".join(f'<option value="{s}">{s}</option>'
                      for s in sorted(df["Status"].unique()))
po_opts     = "".join(f'<option value="{p}">{p}</option>'
                      for p in sorted(df["PO Number"].unique()))

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Inventory Executive Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{{--primary:#1e40af;--danger:#dc2626;--success:#16a34a;--bg:#f1f5f9}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',sans-serif;background:var(--bg);color:#1f2937;font-size:14px}}
header{{background:#1e40af;color:#fff;padding:14px 32px;display:flex;justify-content:space-between;align-items:center}}
header h1{{font-size:18px;font-weight:700}}
.gen{{font-size:11px;opacity:.7}}
.container{{max-width:1500px;margin:24px auto;padding:0 20px}}
/* KPI */
.kpi-row{{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:20px}}
.kpi{{background:#fff;border-radius:10px;padding:14px 16px;border-left:4px solid var(--primary);box-shadow:0 1px 3px rgba(0,0,0,.07)}}
.kpi.red{{border-left-color:var(--danger)}} .kpi.blue{{border-left-color:#3b82f6}}
.kpi.green{{border-left-color:var(--success)}}
.kpi .lbl{{font-size:10px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:.05em}}
.kpi .val{{font-size:22px;font-weight:800;margin-top:4px}}
/* Layout */
.row2{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
.row3{{display:grid;grid-template-columns:1.4fr 1fr;gap:16px;margin-bottom:16px}}
.card{{background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.07)}}
.card h3{{font-size:13px;font-weight:600;color:#374151;margin-bottom:10px}}
/* Filters */
.filters{{background:#fff;border-radius:10px;padding:14px 18px;margin-bottom:16px;
          display:flex;gap:14px;align-items:center;flex-wrap:wrap;box-shadow:0 1px 3px rgba(0,0,0,.07)}}
.filters label{{font-size:12px;font-weight:600;color:#374151}}
.filters select,.filters input{{padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;
                                font-size:12px;font-family:Inter;outline:none}}
.filters select:focus,.filters input:focus{{border-color:#2563eb}}
.btn{{background:#1e40af;color:#fff;border:none;padding:7px 16px;border-radius:6px;
      font-size:12px;font-weight:600;cursor:pointer}}
.btn:hover{{background:#1d4ed8}}
.btn.ghost{{background:#fff;color:#374151;border:1px solid #d1d5db}}
.btn.ghost:hover{{background:#f9fafb}}
/* Table */
.tbl-wrap{{overflow:auto;max-height:460px;border-radius:6px;border:1px solid #e5e7eb}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{position:sticky;top:0;background:#f8fafc;padding:9px 10px;text-align:left;
    font-weight:600;font-size:11px;color:#6b7280;border-bottom:1px solid #e5e7eb;
    white-space:nowrap;z-index:1}}
td{{padding:8px 10px;border-bottom:1px solid #f3f4f6;vertical-align:middle}}
tr:hover td{{background:#f8fafc}}
.num{{text-align:right;font-variant-numeric:tabular-nums}}
.tdanger{{color:var(--danger);font-weight:700;text-align:right}}
.tsuccess{{color:var(--success);font-weight:700;text-align:right}}
.badge{{display:inline-block;padding:2px 8px;border-radius:99px;color:#fff;
        font-size:10px;font-weight:700;white-space:nowrap}}
.empty{{text-align:center;padding:30px;color:#9ca3af}}
/* Highlight selected PO */
tr.selected td{{background:#eff6ff!important}}
@media(max-width:900px){{.row2,.row3{{grid-template-columns:1fr}}.kpi-row{{grid-template-columns:repeat(3,1fr)}}}}
</style>
</head>
<body>

<header>
  <h1>🏢 Inventory Executive Dashboard</h1>
  <span class="gen">Generated: {pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")} &nbsp;|&nbsp; <button class="btn" style="padding:5px 12px;font-size:11px" onclick="location.reload()">Refresh</button></span>
</header>

<div class="container">

<!-- KPI Row -->
<div class="kpi-row">
  <div class="kpi green"><div class="lbl">Fill Rate</div><div class="val" id="k_fill" style="color:var(--success)">{stats['fill_rate']}</div></div>
  <div class="kpi"><div class="lbl">Total POs</div><div class="val" id="k_pos">{stats['total_pos']}</div></div>
  <div class="kpi"><div class="lbl">Ordered (units)</div><div class="val">{stats['total_ordered']}</div></div>
  <div class="kpi green"><div class="lbl">Received (units)</div><div class="val" style="color:var(--success)">{stats['total_received']}</div></div>
  <div class="kpi red"><div class="lbl">Shortage Items</div><div class="val" id="k_short" style="color:var(--danger)">{stats['shortage_count']}</div></div>
  <div class="kpi blue"><div class="lbl">In Transit SKUs</div><div class="val" id="k_transit" style="color:#3b82f6">{stats['in_transit']}</div></div>
</div>

<!-- Filters -->
<div class="filters">
  <div><label>PO#</label><br><select id="f_po" onchange="applyFilters()"><option value="">All POs</option>{po_opts}</select></div>
  <div><label>Status</label><br><select id="f_status" onchange="applyFilters()"><option value="">All</option>{status_opts}</select></div>
  <div><label>Focus</label><br><select id="f_focus" onchange="applyFilters()">
    <option value="all">All Items</option>
    <option value="shortage">⚠️ Shortages Only</option>
  </select></div>
  <div><label>Search SKU</label><br><input type="text" id="f_sku" placeholder="type to search..." oninput="applyFilters()"></div>
  <div style="margin-top:18px"><button class="btn ghost" onclick="resetFilters()">Reset</button></div>
  <div style="margin-top:18px;margin-left:auto;font-size:11px;color:#6b7280" id="rowCount"></div>
</div>

<!-- Row 1: Pipeline + Fill Rate bubble -->
<div class="row2">
  <div class="card"><h3>Inventory Pipeline</h3><div id="c_pipeline"></div></div>
  <div class="card"><h3>Fill Rate by PO <span style="font-size:10px;font-weight:400;color:#9ca3af">(คลิกจุดเพื่อ filter)</span></h3><div id="c_fillrate"></div></div>
</div>

<!-- Row 2: Shortages + Lead Time -->
<div class="row3">
  <div class="card"><h3>Top Shortage Risks</h3><div id="c_shortages"></div></div>
  <div class="card"><h3>Lead Time Trend</h3><div id="c_leadtime"></div></div>
</div>

<!-- Table -->
<div class="card">
  <h3 style="margin-bottom:12px">Tracking Detail <span id="rowCount2" style="font-size:11px;font-weight:400;color:#9ca3af"></span></h3>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>PO#</th><th>SKU</th><th>Status</th>
        <th class="num">Ordered</th><th class="num">Received</th>
        <th class="num">Variance</th><th style="min-width:100px">Fill</th>
        <th>Order Date</th>
      </tr></thead>
      <tbody id="tBody">{rows_html}</tbody>
    </table>
  </div>
</div>

</div><!-- /container -->

<script>
const CHARTS = {json.dumps(charts)};
const plotOpts = {{ paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{{family:'Inter',size:12}} }};

function render(id, key) {{
  const spec = JSON.parse(CHARTS[key]);
  Plotly.newPlot(id, spec.data, {{...spec.layout, ...plotOpts}}, {{responsive:true, displayModeBar:false}});
}}
render('c_pipeline',  'pipeline');
render('c_fillrate',  'fillrate');
render('c_shortages', 'shortages');
render('c_leadtime',  'leadtime');

// Click bubble → filter table
document.getElementById('c_fillrate').on('plotly_click', function(data) {{
  const po = data.points[0]?.hovertext;
  if (!po) return;
  document.getElementById('f_po').value = po;
  applyFilters();
}});

function applyFilters() {{
  const fPO     = document.getElementById('f_po').value;
  const fStatus = document.getElementById('f_status').value;
  const fFocus  = document.getElementById('f_focus').value;
  const fSku    = document.getElementById('f_sku').value.toLowerCase();
  const rows    = document.querySelectorAll('#tBody tr');

  let shown = 0, vOrd = 0, vRec = 0, vShort = 0, vPos = new Set();

  rows.forEach(r => {{
    const rPo     = r.getAttribute('data-po');
    const rStatus = r.getAttribute('data-status');
    const isShort = r.getAttribute('data-shortage') === 'yes';
    const txt     = r.textContent.toLowerCase();

    let show = true;
    if (fPO     && rPo !== fPO)         show = false;
    if (fStatus && rStatus !== fStatus) show = false;
    if (fFocus === 'shortage' && !isShort) show = false;
    if (fSku    && !txt.includes(fSku)) show = false;

    r.style.display = show ? '' : 'none';
    r.classList.toggle('selected', show && !!fPO);

    if (show) {{
      shown++;
      vOrd  += parseFloat(r.cells[3].innerText.replace(/,/g,'')) || 0;
      vRec  += parseFloat(r.cells[4].innerText.replace(/,/g,'')) || 0;
      if (isShort) vShort++;
      vPos.add(rPo);
    }}
  }});

  document.getElementById('k_fill').innerText  = (vOrd > 0 ? (vRec/vOrd*100).toFixed(1) : '0.0') + '%';
  document.getElementById('k_short').innerText = vShort;
  document.getElementById('k_pos').innerText   = vPos.size;
  const msg = `แสดง ${{shown.toLocaleString()}} แถว`;
  document.getElementById('rowCount').innerText  = msg;
  document.getElementById('rowCount2').innerText = msg;
}}

function resetFilters() {{
  ['f_po','f_status','f_focus'].forEach(id => document.getElementById(id).value = id === 'f_focus' ? 'all' : '');
  document.getElementById('f_sku').value = '';
  applyFilters();
}}

// Init count
applyFilters();
</script>
</body>
</html>"""

with open(OUT_PATH, "w", encoding="utf-8") as f:
    f.write(HTML)
print(f"🚀 Dashboard saved → {OUT_PATH}")
"""
po_dashboard_gen.py
────────────────────
สร้าง HTML dashboard จากไฟล์ Excel PO_Items
Usage:  python po_dashboard_gen.py [path/to/PO_Items.xlsx]
Output: po_dashboard.html  (เปิดด้วย browser ได้เลย)
"""

import sys, json, math
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from pathlib import Path

# ── 0. Config ──────────────────────────────────────────────────────────────
EXCEL_PATH = sys.argv[1] if len(sys.argv) > 1 else "PO_Items.xlsx"
OUT_PATH   = Path(EXCEL_PATH).stem + "_dashboard.html"

STATUS_COLOR = {
    "Delivered":   "#22c55e",
    "In transit":  "#3b82f6",
}
STATUS_BG = {
    "Delivered":   "#dcfce7",
    "In transit":  "#dbeafe",
}

# ── 1. Load & prep ─────────────────────────────────────────────────────────
print(f"📂  Loading  {EXCEL_PATH} ...")
df = pd.read_excel(EXCEL_PATH)

DATE_COLS = ['เวลาที่สั่ง','วันที่บรรจุสินค้าเข้าตู้','วันที่เรือออกจริง','วันที่เรือถึงจริง']
for c in DATE_COLS:
    if c in df.columns:
        df[c] = pd.to_datetime(df[c], errors='coerce')

# derive model code
df['model'] = df['ชื่อย่อสินค้า'].fillna(df['รหัสสินค้า'])

def safe_date(ts):
    try:
        if pd.isna(ts): return ''
        return ts.replace(tzinfo=None).strftime('%d/%m/%Y')
    except: return ''

def safe_date_iso(ts):
    try:
        if pd.isna(ts): return ''
        return ts.replace(tzinfo=None).strftime('%Y-%m-%d')
    except: return ''

# ── 2. Aggregations ────────────────────────────────────────────────────────
po_df = df.groupby('รหัสใบจัดซื้อสินค้า').agg(
    po_status =('สถานะใบจัดซื้อสินค้า','first'),
    ship_status=('สถานะ','first'),
    supplier  =('ชื่อผู้จําหน่าย','first'),
    order_date=('เวลาที่สั่ง','min'),
    pack_date =('วันที่บรรจุสินค้าเข้าตู้','first'),
    depart_date=('วันที่เรือออกจริง','first'),
    arrive_date=('วันที่เรือถึงจริง','first'),
    total_ordered =('จำนวนการจัดซื้อสินค้า','sum'),
    total_received=('จำนวนเข้าคลังสำเร็จ','sum'),
    model_count=('model','nunique'),
    sku_count  =('รหัสสินค้า','nunique'),
).reset_index().sort_values('arrive_date')

# Model summary + PO detail string
po_detail = df.groupby(['model','รหัสใบจัดซื้อสินค้า']).agg(
    arrive_date=('วันที่เรือถึงจริง','first'),
    qty=('จำนวนการจัดซื้อสินค้า','sum'),
).reset_index()

def po_str(grp):
    parts = [f"PO{int(r['รหัสใบจัดซื้อสินค้า'])}({safe_date(r['arrive_date'])})"
             for _,r in grp.iterrows()]
    return ' | '.join(parts)

po_str_df = po_detail.groupby('model').apply(po_str).reset_index()
po_str_df.columns = ['model','po_str']

model_df = df.groupby('model').agg(
    total_ordered =('จำนวนการจัดซื้อสินค้า','sum'),
    total_received=('จำนวนเข้าคลังสำเร็จ','sum'),
    po_count=('รหัสใบจัดซื้อสินค้า','nunique'),
).reset_index().merge(po_str_df, on='model').sort_values('total_ordered', ascending=False)

# SKU summary
sku_df = df.groupby(['รหัสสินค้า','model']).agg(
    total_ordered =('จำนวนการจัดซื้อสินค้า','sum'),
    total_received=('จำนวนเข้าคลังสำเร็จ','sum'),
    po_count=('รหัสใบจัดซื้อสินค้า','nunique'),
).reset_index().merge(po_str_df, on='model', how='left').sort_values('total_ordered', ascending=False)

# ── 3. KPIs ────────────────────────────────────────────────────────────────
today = pd.Timestamp.now()
total_po       = len(po_df)
in_transit     = int((po_df['ship_status']=='In transit').sum())
delivered      = int((po_df['ship_status']=='Delivered').sum())
total_ordered  = int(po_df['total_ordered'].sum())
total_received = int(po_df['total_received'].sum())
pending        = total_ordered - total_received
pct_recv       = round(total_received/total_ordered*100, 1) if total_ordered else 0

def clean_dt(ts):
    try:
        if pd.isna(ts): return pd.NaT
        return ts.replace(tzinfo=None)
    except: return pd.NaT

arrive_clean = po_df['arrive_date'].apply(clean_dt)
arriving_7d  = int(((arrive_clean - today).dt.days.between(0,7)).sum())
upcoming     = po_df[arrive_clean >= today].copy()
upcoming['_arrive'] = arrive_clean[arrive_clean >= today]
upcoming = upcoming.sort_values('_arrive')
next_po      = str(int(upcoming.iloc[0]['รหัสใบจัดซื้อสินค้า'])) if len(upcoming) else '—'
next_arrive  = safe_date(upcoming.iloc[0]['arrive_date']) if len(upcoming) else '—'

# ── 4. Charts ──────────────────────────────────────────────────────────────
def fig_to_json(fig):
    return json.dumps(json.loads(pio.to_json(fig)))

# 4a. Donut: ship status
ship_counts = po_df.groupby('ship_status').size().reset_index(name='count')
fig_donut = go.Figure(go.Pie(
    labels=ship_counts['ship_status'], values=ship_counts['count'],
    hole=0.55,
    marker_colors=[STATUS_COLOR.get(s,'#94a3b8') for s in ship_counts['ship_status']],
    textinfo='label+value', textfont_size=13,
))
fig_donut.update_layout(
    margin=dict(t=0,b=0,l=0,r=0), height=240,
    showlegend=False, paper_bgcolor='rgba(0,0,0,0)'
)
chart_donut = fig_to_json(fig_donut)

# 4b. Grouped bar: Ordered vs Received per PO
po_bar = po_df.copy()
po_bar['PO#'] = po_bar['รหัสใบจัดซื้อสินค้า'].astype(str)
fig_bar = go.Figure()
fig_bar.add_trace(go.Bar(name='สั่ง', x=po_bar['PO#'], y=po_bar['total_ordered'],
                         marker_color='#3b82f6', text=po_bar['total_ordered'].apply(lambda v: f'{v:,}'),
                         textposition='outside', textfont_size=10))
fig_bar.add_trace(go.Bar(name='เข้าคลัง', x=po_bar['PO#'], y=po_bar['total_received'],
                         marker_color='#22c55e', text=po_bar['total_received'].apply(lambda v: f'{v:,}'),
                         textposition='outside', textfont_size=10))
fig_bar.update_layout(
    barmode='group', height=280, margin=dict(t=10,b=40,l=40,r=10),
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
    legend=dict(orientation='h', y=1.1), xaxis_title='PO#',
    font=dict(family='Tahoma, sans-serif', size=11),
    yaxis=dict(gridcolor='#e2e8f0'),
)
chart_bar = fig_to_json(fig_bar)

# 4c. Horizontal bar: Top 15 models
top15 = model_df.head(15).copy()
fig_hbar = go.Figure()
fig_hbar.add_trace(go.Bar(
    name='สั่ง', y=top15['model'], x=top15['total_ordered'],
    orientation='h', marker_color='#93c5fd',
    text=top15['total_ordered'].apply(lambda v: f'{v:,}'),
    textposition='outside', textfont_size=10,
))
fig_hbar.add_trace(go.Bar(
    name='เข้าคลัง', y=top15['model'], x=top15['total_received'],
    orientation='h', marker_color='#22c55e',
    text=top15['total_received'].apply(lambda v: f'{v:,}'),
    textposition='outside', textfont_size=10,
))
fig_hbar.update_layout(
    barmode='overlay', height=460, margin=dict(t=10,b=10,l=160,r=60),
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
    legend=dict(orientation='h', y=1.05),
    font=dict(family='Tahoma, sans-serif', size=11),
    xaxis=dict(gridcolor='#e2e8f0'),
)
chart_hbar = fig_to_json(fig_hbar)

# ── 5. Table rows ──────────────────────────────────────────────────────────
def pct_bar(pct, color):
    """mini inline progress bar html"""
    w = min(pct, 100)
    return (f'<div class="pbar-wrap"><div class="pbar-fill" style="width:{w}%;background:{color}"></div>'
            f'<span class="pbar-label">{pct:.1f}%</span></div>')

# ── PO table rows ──
po_rows = ''
for _, r in po_df.iterrows():
    pct  = r['total_received']/r['total_ordered']*100 if r['total_ordered'] else 0
    ship = r['ship_status']
    sc   = STATUS_COLOR.get(ship, '#94a3b8')
    sbg  = STATUS_BG.get(ship, '#f8fafc')
    pct_color = '#22c55e' if pct >= 100 else ('#f59e0b' if pct > 0 else '#ef4444')
    po_rows += f"""
    <tr class="trow" data-po="{r['รหัสใบจัดซื้อสินค้า']}" data-ship="{ship}"
        data-arrive="{safe_date_iso(r['arrive_date'])}" style="background:{sbg}">
      <td class="tc bold">{int(r['รหัสใบจัดซื้อสินค้า'])}</td>
      <td>{r['supplier'] or '—'}</td>
      <td class="tc"><span class="badge" style="background:{sc}">{ship}</span></td>
      <td class="tc">{r['model_count']} model<br><small>{r['sku_count']} SKU</small></td>
      <td class="tc">{safe_date(r['order_date'])}</td>
      <td class="tc">{safe_date(r['pack_date'])}</td>
      <td class="tc">{safe_date(r['depart_date'])}</td>
      <td class="tc bold">{safe_date(r['arrive_date'])}</td>
      <td class="tr">{int(r['total_ordered']):,}</td>
      <td class="tr">{int(r['total_received']):,}</td>
      <td class="tr">{int(r['total_ordered']-r['total_received']):,}</td>
      <td style="min-width:120px">{pct_bar(pct, pct_color)}</td>
    </tr>"""

# ── Model table rows ──
model_rows = ''
for _, r in model_df.iterrows():
    pct = r['total_received']/r['total_ordered']*100 if r['total_ordered'] else 0
    pct_color = '#22c55e' if pct >= 100 else ('#f59e0b' if pct > 0 else '#ef4444')
    model_rows += f"""
    <tr class="trow" data-model="{r['model']}">
      <td class="bold">{r['model']}</td>
      <td class="tr">{int(r['total_ordered']):,}</td>
      <td class="tr">{int(r['total_received']):,}</td>
      <td class="tr">{int(r['total_ordered']-r['total_received']):,}</td>
      <td style="min-width:120px">{pct_bar(pct, pct_color)}</td>
      <td class="tc">{int(r['po_count'])}</td>
      <td class="small">{r['po_str']}</td>
    </tr>"""

# ── SKU table rows ──
sku_rows = ''
for _, r in sku_df.iterrows():
    pct = r['total_received']/r['total_ordered']*100 if r['total_ordered'] else 0
    pct_color = '#22c55e' if pct >= 100 else ('#f59e0b' if pct > 0 else '#ef4444')
    sku_rows += f"""
    <tr class="trow" data-sku="{r['รหัสสินค้า']}" data-model="{r['model']}">
      <td class="bold small">{r['รหัสสินค้า']}</td>
      <td>{r['model']}</td>
      <td class="tr">{int(r['total_ordered']):,}</td>
      <td class="tr">{int(r['total_received']):,}</td>
      <td class="tr">{int(r['total_ordered']-r['total_received']):,}</td>
      <td style="min-width:120px">{pct_bar(pct, pct_color)}</td>
      <td class="tc">{int(r['po_count'])}</td>
      <td class="small">{r['po_str'] if isinstance(r['po_str'],str) else '—'}</td>
    </tr>"""

# ── filter options ──
po_options  = ''.join(f'<option value="{int(p)}">{int(p)}</option>' for p in sorted(po_df['รหัสใบจัดซื้อสินค้า']))
ship_options= ''.join(f'<option value="{s}">{s}</option>' for s in sorted(po_df['ship_status'].unique()))
model_options=''.join(f'<option value="{m}">{m}</option>' for m in sorted(model_df['model']))

# ── 6. HTML ────────────────────────────────────────────────────────────────
HTML = f"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PO Tracker Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
:root{{
  --navy:#1F3864; --teal:#1D6E56; --blue:#2D6FB5; --green:#16a34a;
  --red:#dc2626; --orange:#ea580c; --amber:#d97706;
  --bg:#f0f4f8; --card:#ffffff; --border:#e2e8f0;
  --text:#1e293b; --muted:#64748b; --font:'Tahoma',sans-serif;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:var(--font);background:var(--bg);color:var(--text);font-size:13px}}

/* ── Layout ── */
.wrap{{max-width:1400px;margin:0 auto;padding:16px}}
.top-bar{{background:var(--navy);color:#fff;padding:14px 20px;border-radius:12px;
          display:flex;align-items:center;gap:12px;margin-bottom:16px}}
.top-bar h1{{font-size:20px;font-weight:700;flex:1}}
.top-bar .sub{{font-size:11px;opacity:.7}}
.card{{background:var(--card);border-radius:10px;padding:16px;
       box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:14px}}
.section-title{{font-size:13px;font-weight:700;color:var(--navy);
                margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid var(--border)}}

/* ── KPI grid ── */
.kpi-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:14px}}
.kpi{{background:var(--card);border-radius:10px;padding:12px 14px;
      box-shadow:0 1px 3px rgba(0,0,0,.08);border-top:3px solid var(--navy)}}
.kpi-label{{font-size:10px;color:var(--muted);font-weight:600;text-transform:uppercase;margin-bottom:4px}}
.kpi-val{{font-size:20px;font-weight:700;line-height:1.1}}
.kpi-sub{{font-size:10px;color:var(--muted);margin-top:3px}}

/* ── Charts row ── */
.chart-row{{display:grid;grid-template-columns:280px 1fr;gap:14px;margin-bottom:14px}}
.chart-full{{margin-bottom:14px}}

/* ── Tabs ── */
.tabs{{display:flex;gap:6px;margin-bottom:12px;flex-wrap:wrap}}
.tab{{padding:7px 16px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;
      border:1.5px solid var(--border);background:var(--card);color:var(--muted)}}
.tab.active{{background:var(--navy);color:#fff;border-color:var(--navy)}}

/* ── Filter bar ── */
.filter-bar{{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;
             background:var(--card);border-radius:10px;padding:12px 14px;
             box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:12px}}
.filter-group{{display:flex;flex-direction:column;gap:4px}}
.filter-group label{{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase}}
.filter-group select,.filter-group input{{
  padding:6px 10px;border:1.5px solid var(--border);border-radius:6px;
  font-size:12px;color:var(--text);background:#fff;min-width:130px}}
.btn{{padding:7px 18px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:700;border:none}}
.btn-apply{{background:var(--blue);color:#fff}}
.btn-reset{{background:var(--border);color:var(--muted)}}

/* ── Tables ── */
.tbl-wrap{{overflow-x:auto;max-height:480px;border-radius:8px;border:1px solid var(--border)}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
thead th{{position:sticky;top:0;z-index:2;background:var(--navy);color:#fff;
          padding:9px 10px;font-size:11px;font-weight:600;white-space:nowrap;
          border-right:1px solid rgba(255,255,255,.12)}}
thead th:last-child{{border-right:none}}
tbody tr.trow:hover td{{filter:brightness(.96)}}
tbody td{{padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:middle;white-space:nowrap}}
.tc{{text-align:center}}
.tr{{text-align:right}}
.bold{{font-weight:600}}
.small{{font-size:11px;white-space:normal!important;min-width:180px}}

/* ── Badge ── */
.badge{{display:inline-block;padding:2px 9px;border-radius:20px;
        color:#fff;font-size:10px;font-weight:700;white-space:nowrap}}

/* ── Progress bar ── */
.pbar-wrap{{display:flex;align-items:center;gap:6px}}
.pbar-fill{{height:7px;border-radius:4px;min-width:2px}}
.pbar-label{{font-size:11px;font-weight:600;white-space:nowrap;min-width:36px}}

/* hidden tab panels */
.tab-panel{{display:none}}.tab-panel.active{{display:block}}

/* no-result */
.no-result{{text-align:center;padding:30px;color:var(--muted)}}
</style>
</head>
<body>
<div class="wrap">

<!-- ── Title ── -->
<div class="top-bar">
  <span style="font-size:24px">🚢</span>
  <div>
    <h1>PO & Product Tracker Dashboard</h1>
    <div class="sub">อัปเดต: {today.strftime('%d/%m/%Y %H:%M')}  |  ข้อมูลจาก: {Path(EXCEL_PATH).name}</div>
  </div>
</div>

<!-- ── KPI Cards ── -->
<div class="kpi-grid">
  <div class="kpi">
    <div class="kpi-label">ใบจัดซื้อทั้งหมด</div>
    <div class="kpi-val" style="color:var(--navy)">{total_po}</div>
    <div class="kpi-sub">In Transit {in_transit}  ·  Delivered {delivered}</div>
  </div>
  <div class="kpi" style="border-top-color:#8b5cf6">
    <div class="kpi-label">ถึงภายใน 7 วัน</div>
    <div class="kpi-val" style="color:#8b5cf6">{arriving_7d}</div>
    <div class="kpi-sub">ถัดไป: PO {next_po} ({next_arrive})</div>
  </div>
  <div class="kpi" style="border-top-color:var(--blue)">
    <div class="kpi-label">สั่งซื้อรวม (ชิ้น)</div>
    <div class="kpi-val" style="color:var(--blue)">{total_ordered:,}</div>
    <div class="kpi-sub">ทุก PO รวมกัน</div>
  </div>
  <div class="kpi" style="border-top-color:var(--green)">
    <div class="kpi-label">เข้าคลังแล้ว</div>
    <div class="kpi-val" style="color:var(--green)">{total_received:,}</div>
    <div class="kpi-sub">{pct_recv}% ของที่สั่ง</div>
  </div>
  <div class="kpi" style="border-top-color:var(--orange)">
    <div class="kpi-label">รอเข้าคลัง</div>
    <div class="kpi-val" style="color:var(--orange)">{pending:,}</div>
    <div class="kpi-sub">ชิ้นที่ยังไม่รับเข้า</div>
  </div>
  <div class="kpi" style="border-top-color:var(--teal)">
    <div class="kpi-label">% รับเข้าคลัง</div>
    <div class="kpi-val" style="color:var(--teal)">{pct_recv}%</div>
    <div class="kpi-sub">Fulfillment rate</div>
  </div>
</div>

<!-- ── Charts ── -->
<div class="chart-row">
  <div class="card" style="display:flex;flex-direction:column;align-items:center;justify-content:center">
    <div class="section-title" style="width:100%">สถานะการขนส่ง</div>
    <div id="chart_donut" style="width:100%"></div>
  </div>
  <div class="card">
    <div class="section-title">สั่ง vs เข้าคลัง แยกตาม PO</div>
    <div id="chart_bar"></div>
  </div>
</div>
<div class="card chart-full">
  <div class="section-title">Top 15 สินค้า — สั่ง vs เข้าคลัง</div>
  <div id="chart_hbar"></div>
</div>

<!-- ── Tabs ── -->
<div class="tabs">
  <div class="tab active" onclick="switchTab('po',this)">📋 PO Overview</div>
  <div class="tab" onclick="switchTab('model',this)">📈 รายสินค้า (Model)</div>
  <div class="tab" onclick="switchTab('sku',this)">📦 Product Tracker (SKU)</div>
</div>

<!-- ════════════ TAB: PO ════════════ -->
<div id="tab-po" class="tab-panel active">
  <div class="filter-bar">
    <div class="filter-group"><label>PO#</label>
      <select id="f_po"><option value="">ทั้งหมด</option>{po_options}</select></div>
    <div class="filter-group"><label>สถานะเรือ</label>
      <select id="f_ship"><option value="">ทั้งหมด</option>{ship_options}</select></div>
    <div class="filter-group"><label>เรือถึง (ตั้งแต่)</label>
      <input type="date" id="f_arrive_from"></div>
    <div class="filter-group"><label>เรือถึง (ถึง)</label>
      <input type="date" id="f_arrive_to"></div>
    <button class="btn btn-apply" onclick="filterPO()">🔍 Apply</button>
    <button class="btn btn-reset" onclick="resetPO()">↺ Reset</button>
    <span id="po_count" style="font-size:12px;color:var(--muted);align-self:center"></span>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>PO#</th><th>Supplier</th><th>สถานะ</th><th>Model/SKU</th>
        <th>วันสั่ง</th><th>บรรจุตู้</th><th>เรือออก</th><th>เรือถึง</th>
        <th>สั่ง (ชิ้น)</th><th>เข้าคลัง</th><th>รอเข้า</th><th>% เข้าคลัง</th>
      </tr></thead>
      <tbody id="tb_po">{po_rows}</tbody>
    </table>
  </div>
</div>

<!-- ════════════ TAB: MODEL ════════════ -->
<div id="tab-model" class="tab-panel">
  <div class="filter-bar">
    <div class="filter-group"><label>ค้นหา Model</label>
      <input type="text" id="f_model_search" placeholder="พิมพ์ชื่อ model..." style="min-width:200px"
             oninput="filterModel()"></div>
    <div class="filter-group"><label>% เข้าคลัง</label>
      <select id="f_model_pct" onchange="filterModel()">
        <option value="">ทั้งหมด</option>
        <option value="0">ยังไม่ได้รับ (0%)</option>
        <option value="partial">รับบางส่วน (&gt;0% &lt;100%)</option>
        <option value="full">รับครบ (100%)</option>
      </select></div>
    <span id="model_count" style="font-size:12px;color:var(--muted);align-self:center"></span>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>รหัสสินค้า (Model)</th><th>สั่งรวม</th><th>เข้าคลัง</th>
        <th>รอเข้า</th><th>% เข้าคลัง</th><th>PO</th><th>PO ที่เกี่ยวข้อง (วันเรือถึง)</th>
      </tr></thead>
      <tbody id="tb_model">{model_rows}</tbody>
    </table>
  </div>
</div>

<!-- ════════════ TAB: SKU ════════════ -->
<div id="tab-sku" class="tab-panel">
  <div class="filter-bar">
    <div class="filter-group"><label>ค้นหา SKU / Model</label>
      <input type="text" id="f_sku_search" placeholder="พิมพ์ SKU หรือ model..." style="min-width:220px"
             oninput="filterSKU()"></div>
    <div class="filter-group"><label>% เข้าคลัง</label>
      <select id="f_sku_pct" onchange="filterSKU()">
        <option value="">ทั้งหมด</option>
        <option value="0">ยังไม่ได้รับ (0%)</option>
        <option value="partial">รับบางส่วน</option>
        <option value="full">รับครบ (100%)</option>
      </select></div>
    <span id="sku_count" style="font-size:12px;color:var(--muted);align-self:center"></span>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>รหัสสินค้า (SKU)</th><th>Model</th><th>สั่งรวม</th><th>เข้าคลัง</th>
        <th>รอเข้า</th><th>% เข้าคลัง</th><th>PO</th><th>PO ที่เกี่ยวข้อง (วันเรือถึง)</th>
      </tr></thead>
      <tbody id="tb_sku">{sku_rows}</tbody>
    </table>
  </div>
</div>

</div><!-- /wrap -->

<!-- ── Scripts ── -->
<script>
const CHARTS = {{
  donut: {chart_donut},
  bar:   {chart_bar},
  hbar:  {chart_hbar},
}};

function renderChart(id, obj) {{
  Plotly.newPlot(id, obj.data, {{...obj.layout, responsive:true}});
}}
renderChart('chart_donut', CHARTS.donut);
renderChart('chart_bar',   CHARTS.bar);
renderChart('chart_hbar',  CHARTS.hbar);

// ── Tab switch ──
function switchTab(name, el) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  el.classList.add('active');
}}

// ── PO filter ──
function filterPO() {{
  const po    = document.getElementById('f_po').value;
  const ship  = document.getElementById('f_ship').value;
  const from_ = document.getElementById('f_arrive_from').value;
  const to_   = document.getElementById('f_arrive_to').value;
  const rows  = document.querySelectorAll('#tb_po .trow');
  let vis = 0;
  rows.forEach(r => {{
    const rpo   = r.getAttribute('data-po');
    const rship = r.getAttribute('data-ship');
    const rarr  = r.getAttribute('data-arrive');
    let show = true;
    if (po   && rpo   !== po)   show = false;
    if (ship && rship !== ship) show = false;
    if (from_ && rarr && rarr < from_) show = false;
    if (to_   && rarr && rarr > to_)   show = false;
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  document.getElementById('po_count').textContent = `แสดง ${{vis}} / ${{rows.length}} รายการ`;
}}
function resetPO() {{
  ['f_po','f_ship'].forEach(id => document.getElementById(id).value='');
  ['f_arrive_from','f_arrive_to'].forEach(id => document.getElementById(id).value='');
  filterPO();
}}
filterPO();

// ── Model filter ──
function filterModel() {{
  const q   = document.getElementById('f_model_search').value.toLowerCase();
  const pct = document.getElementById('f_model_pct').value;
  const rows = document.querySelectorAll('#tb_model .trow');
  let vis = 0;
  rows.forEach(r => {{
    const name = r.getAttribute('data-model').toLowerCase();
    const pbar = r.querySelector('.pbar-label');
    const pval = pbar ? parseFloat(pbar.textContent) : 0;
    let show = name.includes(q);
    if (pct === '0'       && pval > 0)             show = false;
    if (pct === 'partial' && (pval === 0 || pval >= 100)) show = false;
    if (pct === 'full'    && pval < 100)            show = false;
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  document.getElementById('model_count').textContent = `แสดง ${{vis}} / ${{rows.length}} รายการ`;
}}
filterModel();

// ── SKU filter ──
function filterSKU() {{
  const q   = document.getElementById('f_sku_search').value.toLowerCase();
  const pct = document.getElementById('f_sku_pct').value;
  const rows = document.querySelectorAll('#tb_sku .trow');
  let vis = 0;
  rows.forEach(r => {{
    const sku   = r.getAttribute('data-sku').toLowerCase();
    const model = r.getAttribute('data-model').toLowerCase();
    const pbar  = r.querySelector('.pbar-label');
    const pval  = pbar ? parseFloat(pbar.textContent) : 0;
    let show = (sku.includes(q) || model.includes(q));
    if (pct === '0'       && pval > 0)             show = false;
    if (pct === 'partial' && (pval === 0 || pval >= 100)) show = false;
    if (pct === 'full'    && pval < 100)            show = false;
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  document.getElementById('sku_count').textContent = `แสดง ${{vis}} / ${{rows.length}} รายการ`;
}}
filterSKU();
</script>
</body>
</html>"""

# ── 7. Write output ────────────────────────────────────────────────────────
with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"✅  Dashboard saved → {OUT_PATH}")
print(f"📌  เปิดไฟล์ด้วย browser ได้เลย")
print(f"📌  ครั้งหน้าอัปเดตข้อมูลใน {EXCEL_PATH} แล้วรันสคริปต์ใหม่ dashboard จะ refresh อัตโนมัติ")
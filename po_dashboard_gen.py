"""
po_dashboard_gen.py  (v3 — full SKU tracking edition)
Usage:  python po_dashboard_gen.py [path/to/Data_for_track.xlsx]
Output: PO_Items_dashboard.html
"""

import sys, json
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from pathlib import Path

EXCEL_PATH = sys.argv[1] if len(sys.argv) > 1 else "Data_for_track.xlsx"
OUT_PATH   = Path(EXCEL_PATH).stem + "_dashboard.html"

STATUS_COLOR = {"Delivered": "#22c55e", "In transit": "#3b82f6"}
STATUS_BG    = {"Delivered": "#dcfce7",  "In transit": "#dbeafe"}

# ── 1. Load ─────────────────────────────────────────────────────────────────
print(f"📂  Loading {EXCEL_PATH} ...")
df = pd.read_excel(EXCEL_PATH)

DATE_COLS = ['เวลาที่สั่ง','วันที่บรรจุสินค้าเข้าตู้','วันที่เรือออกจริง','วันที่เรือถึงจริง']
for c in DATE_COLS:
    if c in df.columns:
        df[c] = pd.to_datetime(df[c], errors='coerce')

for col in ['คำสั่งซื้อ','ใบPL','จำนวนเข้าคลังสำเร็จ']:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

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

# ── 2. Aggregations ──────────────────────────────────────────────────────────
po_df = df.groupby('รหัสใบจัดซื้อสินค้า').agg(
    ship_status =('สถานะ','first'),
    supplier    =('ชื่อผู้จําหน่าย','first'),
    container   =('หมายเลขตู้','first'),
    order_date  =('เวลาที่สั่ง','min'),
    pack_date   =('วันที่บรรจุสินค้าเข้าตู้','first'),
    depart_date =('วันที่เรือออกจริง','first'),
    arrive_date =('วันที่เรือถึงจริง','first'),
    total_ordered  =('คำสั่งซื้อ','sum'),
    total_pl       =('ใบPL','sum'),
    total_received =('จำนวนเข้าคลังสำเร็จ','sum'),
    model_count =('model','nunique'),
    sku_count   =('รหัสสินค้า','nunique'),
).reset_index().sort_values('arrive_date')

today = pd.Timestamp.now()

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
next_po     = str(int(upcoming.iloc[0]['รหัสใบจัดซื้อสินค้า'])) if len(upcoming) else '—'
next_arrive = safe_date(upcoming.iloc[0]['arrive_date']) if len(upcoming) else '—'

# KPIs
total_po       = len(po_df)
in_transit     = int((po_df['ship_status']=='In transit').sum())
delivered      = int((po_df['ship_status']=='Delivered').sum())
total_ordered  = int(po_df['total_ordered'].sum())
total_pl_sum   = int(po_df['total_pl'].sum())
total_received = int(po_df['total_received'].sum())
pending        = total_ordered - total_received
pct_recv       = round(total_received/total_ordered*100,1) if total_ordered else 0
pct_pl         = round(total_pl_sum/total_ordered*100,1)  if total_ordered else 0

# ── 3. Charts ────────────────────────────────────────────────────────────────
def fig_to_json(fig):
    return json.dumps(json.loads(pio.to_json(fig)))

# Donut
ship_counts = po_df.groupby('ship_status').size().reset_index(name='count')
fig_donut = go.Figure(go.Pie(
    labels=ship_counts['ship_status'], values=ship_counts['count'],
    hole=0.55,
    marker_colors=[STATUS_COLOR.get(s,'#94a3b8') for s in ship_counts['ship_status']],
    textinfo='label+value', textfont_size=13,
))
fig_donut.update_layout(margin=dict(t=10,b=10,l=10,r=10), height=240,
    showlegend=True, legend=dict(orientation='h',y=-0.1,x=0.5,xanchor='center'),
    paper_bgcolor='rgba(0,0,0,0)', autosize=True)
chart_donut = fig_to_json(fig_donut)

# Grouped bar: Order vs PL vs Received per PO
po_bar = po_df.copy()
po_bar['PO#'] = po_bar['รหัสใบจัดซื้อสินค้า'].astype(str)
max_val = po_bar['total_ordered'].max()
y_max   = max_val * 1.28

fig_bar = go.Figure()
fig_bar.add_trace(go.Bar(name='Ordered', x=po_bar['PO#'], y=po_bar['total_ordered'],
    marker_color='#3b82f6', text=po_bar['total_ordered'].apply(lambda v:f'{int(v):,}'),
    textposition='outside', textfont=dict(size=10), constraintext='none'))
fig_bar.add_trace(go.Bar(name='PL', x=po_bar['PO#'], y=po_bar['total_pl'],
    marker_color='#f59e0b', text=po_bar['total_pl'].apply(lambda v:f'{int(v):,}'),
    textposition='outside', textfont=dict(size=10), constraintext='none'))
fig_bar.add_trace(go.Bar(name='Received', x=po_bar['PO#'], y=po_bar['total_received'],
    marker_color='#22c55e', text=po_bar['total_received'].apply(lambda v:f'{int(v):,}'),
    textposition='outside', textfont=dict(size=10), constraintext='none'))
fig_bar.update_layout(
    barmode='group', height=320, margin=dict(t=30,b=50,l=60,r=20),
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
    legend=dict(orientation='h',y=1.12,x=0),
    xaxis=dict(title='PO#', tickangle=-30 if len(po_bar)>6 else 0),
    yaxis=dict(gridcolor='#e2e8f0', range=[0,y_max]),
    font=dict(family='Tahoma,sans-serif',size=11),
    bargap=0.2, bargroupgap=0.05, autosize=True)
chart_bar = fig_to_json(fig_bar)

# Fulfillment funnel per PO (stacked bar showing gaps)
fig_funnel = go.Figure()
po_bar2 = po_bar.copy()
po_bar2['gap_pl']   = (po_bar2['total_ordered'] - po_bar2['total_pl']).clip(lower=0)
po_bar2['gap_recv'] = (po_bar2['total_pl']      - po_bar2['total_received']).clip(lower=0)

fig_funnel.add_trace(go.Bar(name='Received', x=po_bar2['PO#'], y=po_bar2['total_received'],
    marker_color='#22c55e', hovertemplate='PO %{x}<br>Received: %{y:,}<extra></extra>'))
fig_funnel.add_trace(go.Bar(name='PL not received', x=po_bar2['PO#'], y=po_bar2['gap_recv'],
    marker_color='#f59e0b', hovertemplate='PO %{x}<br>In PL, not received: %{y:,}<extra></extra>'))
fig_funnel.add_trace(go.Bar(name='Not in PL', x=po_bar2['PO#'], y=po_bar2['gap_pl'],
    marker_color='#ef4444', hovertemplate='PO %{x}<br>Ordered, not in PL: %{y:,}<extra></extra>'))
fig_funnel.update_layout(
    barmode='stack', height=280, margin=dict(t=20,b=50,l=60,r=20),
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
    legend=dict(orientation='h',y=1.12,x=0),
    xaxis=dict(title='PO#', tickangle=-30 if len(po_bar2)>6 else 0),
    yaxis=dict(gridcolor='#e2e8f0'),
    font=dict(family='Tahoma,sans-serif',size=11), autosize=True)
chart_funnel = fig_to_json(fig_funnel)

# ── 4. Progress bar helper ───────────────────────────────────────────────────
def pct_bar(pct, color):
    w = min(pct, 100)
    return (f'<div class="pbar-wrap"><div class="pbar-fill" style="width:{w}%;background:{color}"></div>'
            f'<span class="pbar-label">{pct:.1f}%</span></div>')

def status_dot(ordered, pl, received):
    """3-stage traffic light"""
    if received >= ordered:    return '<span class="dot dot-green" title="Fully received">●</span>'
    elif received > 0:         return '<span class="dot dot-yellow" title="Partially received">●</span>'
    elif pl >= ordered:        return '<span class="dot dot-blue" title="PL complete, not received">●</span>'
    elif pl > 0:               return '<span class="dot dot-orange" title="Partial PL">●</span>'
    else:                      return '<span class="dot dot-red" title="Not in PL yet">●</span>'

# ── 5. PO table rows ─────────────────────────────────────────────────────────
po_rows = ''
for _, r in po_df.iterrows():
    pct_p = r['total_pl']/r['total_ordered']*100      if r['total_ordered'] else 0
    pct_r = r['total_received']/r['total_ordered']*100 if r['total_ordered'] else 0
    ship  = r['ship_status']
    sc    = STATUS_COLOR.get(ship,'#94a3b8')
    sbg   = STATUS_BG.get(ship,'#f8fafc')
    gap_pl   = int(r['total_ordered'] - r['total_pl'])
    gap_recv = int(r['total_pl']      - r['total_received'])
    po_rows += f"""
    <tr class="trow" data-po="{r['รหัสใบจัดซื้อสินค้า']}" data-ship="{ship}"
        data-arrive="{safe_date_iso(r['arrive_date'])}" style="background:{sbg}">
      <td class="tc bold">{int(r['รหัสใบจัดซื้อสินค้า'])}</td>
      <td>{r['supplier'] or '—'}</td>
      <td class="tc"><span class="badge" style="background:{sc}">{ship}</span></td>
      <td class="tc small">{r['container'] or '—'}</td>
      <td class="tc">{r['model_count']} model<br><small>{r['sku_count']} SKU</small></td>
      <td class="tc">{safe_date(r['order_date'])}</td>
      <td class="tc">{safe_date(r['pack_date'])}</td>
      <td class="tc">{safe_date(r['depart_date'])}</td>
      <td class="tc bold">{safe_date(r['arrive_date'])}</td>
      <td class="tr">{int(r['total_ordered']):,}</td>
      <td class="tr">{int(r['total_pl']):,}</td>
      <td class="tr {'warn' if gap_pl>0 else ''}">{gap_pl:,}</td>
      <td class="tr">{int(r['total_received']):,}</td>
      <td class="tr {'warn' if gap_recv>0 else ''}">{gap_recv:,}</td>
      <td style="min-width:130px">{pct_bar(pct_r,'#22c55e' if pct_r>=100 else '#f59e0b' if pct_r>0 else '#ef4444')}</td>
    </tr>"""

# ── 6. SKU table rows ─────────────────────────────────────────────────────────
# Full SKU detail — every row in the Excel
sku_rows = ''
sku_df = df.sort_values(['รหัสใบจัดซื้อสินค้า','รหัสสินค้า'])
for _, r in sku_df.iterrows():
    ordered  = int(r['คำสั่งซื้อ'])
    pl_qty   = int(r['ใบPL'])
    received = int(r['จำนวนเข้าคลังสำเร็จ'])
    gap_pl   = ordered - pl_qty
    gap_recv = pl_qty  - received
    ship = r['สถานะ']
    sc   = STATUS_COLOR.get(ship,'#94a3b8')
    sbg  = STATUS_BG.get(ship,'#f8fafc')
    pct_r = received/ordered*100 if ordered else 0
    pct_p = pl_qty/ordered*100   if ordered else 0
    dot   = status_dot(ordered, pl_qty, received)

    sku_rows += f"""
    <tr class="trow"
        data-po="{r['รหัสใบจัดซื้อสินค้า']}"
        data-ship="{ship}"
        data-model="{r['model']}"
        data-sku="{r['รหัสสินค้า']}"
        data-arrive="{safe_date_iso(r['วันที่เรือถึงจริง'])}"
        style="background:{sbg}">
      <td class="tc bold">{int(r['รหัสใบจัดซื้อสินค้า'])}</td>
      <td class="tc"><span class="badge" style="background:{sc}">{ship}</span></td>
      <td class="tc">{safe_date(r['วันที่เรือถึงจริง'])}</td>
      <td class="bold small">{r['รหัสสินค้า']}</td>
      <td>{r['model']}</td>
      <td class="small muted">{r['รูปแบบ'] or '—'}</td>
      <td class="tr">{ordered:,}</td>
      <td class="tr {'warn' if gap_pl>0 else 'ok' if pl_qty>=ordered else ''}">{pl_qty:,}</td>
      <td class="tr {'warn' if gap_pl>0 else ''}">{gap_pl:,}</td>
      <td class="tr {'ok' if received>=ordered else 'warn' if gap_recv>0 else ''}">{received:,}</td>
      <td class="tr {'warn' if gap_recv>0 else ''}">{gap_recv:,}</td>
      <td>{dot}</td>
      <td style="min-width:110px">{pct_bar(pct_r,'#22c55e' if pct_r>=100 else '#f59e0b' if pct_r>0 else '#ef4444')}</td>
    </tr>"""

# ── 7. Model summary rows ─────────────────────────────────────────────────────
model_grp = df.groupby('model').agg(
    total_ordered  =('คำสั่งซื้อ','sum'),
    total_pl       =('ใบPL','sum'),
    total_received =('จำนวนเข้าคลังสำเร็จ','sum'),
    sku_count      =('รหัสสินค้า','nunique'),
    po_count       =('รหัสใบจัดซื้อสินค้า','nunique'),
).reset_index().sort_values('total_ordered',ascending=False)

model_rows = ''
for _, r in model_grp.iterrows():
    pct_r = r['total_received']/r['total_ordered']*100 if r['total_ordered'] else 0
    pct_p = r['total_pl']/r['total_ordered']*100       if r['total_ordered'] else 0
    gap_pl   = int(r['total_ordered'] - r['total_pl'])
    gap_recv = int(r['total_pl']      - r['total_received'])
    dot = status_dot(r['total_ordered'], r['total_pl'], r['total_received'])
    model_rows += f"""
    <tr class="trow" data-model="{r['model']}">
      <td class="bold">{r['model']}</td>
      <td class="tc">{int(r['sku_count'])}</td>
      <td class="tc">{int(r['po_count'])}</td>
      <td class="tr">{int(r['total_ordered']):,}</td>
      <td class="tr {'warn' if gap_pl>0 else ''}">{int(r['total_pl']):,}</td>
      <td class="tr {'warn' if gap_pl>0 else ''}">{gap_pl:,}</td>
      <td class="tr">{int(r['total_received']):,}</td>
      <td class="tr {'warn' if gap_recv>0 else ''}">{gap_recv:,}</td>
      <td>{dot}</td>
      <td style="min-width:110px">{pct_bar(pct_r,'#22c55e' if pct_r>=100 else '#f59e0b' if pct_r>0 else '#ef4444')}</td>
    </tr>"""

# ── 8. Filter options ─────────────────────────────────────────────────────────
po_options   = ''.join(f'<option value="{int(p)}">{int(p)}</option>' for p in sorted(po_df['รหัสใบจัดซื้อสินค้า']))
ship_options = ''.join(f'<option value="{s}">{s}</option>' for s in sorted(df['สถานะ'].dropna().unique()))
model_options= ''.join(f'<option value="{m}">{m}</option>' for m in sorted(model_grp['model']))

# ── 9. HTML ────────────────────────────────────────────────────────────────────
HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PO Tracker Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
<script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
<style>
:root{{
  --navy:#1F3864;--teal:#1D6E56;--blue:#2D6FB5;--green:#16a34a;
  --red:#dc2626;--orange:#ea580c;--amber:#d97706;
  --bg:#f0f4f8;--card:#ffffff;--border:#e2e8f0;
  --text:#1e293b;--muted:#64748b;--font:'Tahoma',sans-serif;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:var(--font);background:var(--bg);color:var(--text);font-size:13px;-webkit-text-size-adjust:100%}}
.wrap{{max-width:1500px;margin:0 auto;padding:12px}}

/* Top bar */
.top-bar{{background:var(--navy);color:#fff;padding:12px 18px;border-radius:10px;
          display:flex;align-items:center;gap:12px;margin-bottom:14px;flex-wrap:wrap}}
.top-bar h1{{font-size:18px;font-weight:700;flex:1;min-width:200px}}
.top-bar .sub{{font-size:11px;opacity:.7}}

/* Cards */
.card{{background:var(--card);border-radius:10px;padding:14px;
       box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:12px}}
.section-title{{font-size:13px;font-weight:700;color:var(--navy);
                margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid var(--border)}}

/* KPI */
.kpi-grid{{display:grid;grid-template-columns:repeat(7,1fr);gap:8px;margin-bottom:12px}}
@media(max-width:1200px){{.kpi-grid{{grid-template-columns:repeat(4,1fr)}}}}
@media(max-width:640px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
.kpi{{background:var(--card);border-radius:10px;padding:10px 12px;
      box-shadow:0 1px 3px rgba(0,0,0,.08);border-top:3px solid var(--navy)}}
.kpi-label{{font-size:10px;color:var(--muted);font-weight:600;text-transform:uppercase;margin-bottom:3px}}
.kpi-val{{font-size:clamp(15px,2vw,22px);font-weight:700;line-height:1.1}}
.kpi-sub{{font-size:10px;color:var(--muted);margin-top:3px}}

/* Charts */
.chart-row{{display:grid;grid-template-columns:260px 1fr 1fr;gap:12px;margin-bottom:12px}}
@media(max-width:900px){{.chart-row{{grid-template-columns:1fr}}}}

/* Tabs */
.tabs{{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap}}
.tab{{padding:7px 16px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;
      border:1.5px solid var(--border);background:var(--card);color:var(--muted)}}
.tab.active{{background:var(--navy);color:#fff;border-color:var(--navy)}}
.tab-panel{{display:none}}.tab-panel.active{{display:block}}

/* Filters */
.filter-bar{{display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end;
             background:var(--card);border-radius:10px;padding:10px 12px;
             box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:10px}}
.filter-group{{display:flex;flex-direction:column;gap:3px}}
.filter-group label{{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase}}
.filter-group select,.filter-group input{{
  padding:5px 8px;border:1.5px solid var(--border);border-radius:6px;
  font-size:12px;color:var(--text);background:#fff;min-width:110px}}
.btn{{padding:6px 14px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:700;border:none}}
.btn-apply{{background:var(--blue);color:#fff}}
.btn-reset{{background:var(--border);color:var(--muted)}}

/* Table */
.tbl-wrap{{overflow-x:auto;max-height:520px;border-radius:8px;border:1px solid var(--border);
           -webkit-overflow-scrolling:touch}}
.tbl-wrap::-webkit-scrollbar{{height:6px;width:6px}}
.tbl-wrap::-webkit-scrollbar-thumb{{background:#cbd5e1;border-radius:3px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
thead th{{position:sticky;top:0;z-index:2;background:var(--navy);color:#fff;
          padding:8px 10px;font-size:11px;font-weight:600;white-space:nowrap;
          border-right:1px solid rgba(255,255,255,.12)}}
thead th:last-child{{border-right:none}}
tbody tr.trow:hover td{{filter:brightness(.96)}}
tbody td{{padding:7px 10px;border-bottom:1px solid var(--border);vertical-align:middle;white-space:nowrap}}
.tc{{text-align:center}}.tr{{text-align:right}}
.bold{{font-weight:600}}.small{{font-size:11px}}.muted{{color:var(--muted)}}
.warn{{color:#dc2626;font-weight:600}}
.ok{{color:#16a34a;font-weight:600}}

/* Badge */
.badge{{display:inline-block;padding:2px 8px;border-radius:20px;
        color:#fff;font-size:10px;font-weight:700;white-space:nowrap}}

/* Progress bar */
.pbar-wrap{{display:flex;align-items:center;gap:5px}}
.pbar-fill{{height:7px;border-radius:4px;min-width:2px}}
.pbar-label{{font-size:11px;font-weight:600;white-space:nowrap;min-width:38px}}

/* Status dots */
.dot{{font-size:14px;cursor:help}}
.dot-green{{color:#22c55e}}.dot-yellow{{color:#f59e0b}}
.dot-blue{{color:#3b82f6}}.dot-orange{{color:#ea580c}}.dot-red{{color:#ef4444}}

/* Legend */
.dot-legend{{display:flex;flex-wrap:wrap;gap:12px;font-size:11px;color:var(--muted);margin-bottom:8px}}
.dot-legend span{{display:flex;align-items:center;gap:4px}}
</style>
</head>
<body>
<div class="wrap">

<!-- Title -->
<div class="top-bar">
  <span style="font-size:22px">🚢</span>
  <div>
    <h1>PO & Product Tracker Dashboard</h1>
    <div class="sub">Updated: {today.strftime('%d/%m/%Y %H:%M')}  |  Source: {Path(EXCEL_PATH).name}</div>
  </div>
</div>

<!-- KPIs -->
<div class="kpi-grid">
  <div class="kpi">
    <div class="kpi-label">Total POs</div>
    <div class="kpi-val" style="color:var(--navy)">{total_po}</div>
    <div class="kpi-sub">In Transit {in_transit} · Delivered {delivered}</div>
  </div>
  <div class="kpi" style="border-top-color:#8b5cf6">
    <div class="kpi-label">Arriving in 7 Days</div>
    <div class="kpi-val" style="color:#8b5cf6">{arriving_7d}</div>
    <div class="kpi-sub">Next: PO {next_po} ({next_arrive})</div>
  </div>
  <div class="kpi" style="border-top-color:var(--blue)">
    <div class="kpi-label">Total Ordered (pcs)</div>
    <div class="kpi-val" style="color:var(--blue)">{total_ordered:,}</div>
    <div class="kpi-sub">Across all POs</div>
  </div>
  <div class="kpi" style="border-top-color:var(--amber)">
    <div class="kpi-label">Total PL (pcs)</div>
    <div class="kpi-val" style="color:var(--amber)">{total_pl_sum:,}</div>
    <div class="kpi-sub">{pct_pl}% of ordered</div>
  </div>
  <div class="kpi" style="border-top-color:var(--green)">
    <div class="kpi-label">Received</div>
    <div class="kpi-val" style="color:var(--green)">{total_received:,}</div>
    <div class="kpi-sub">{pct_recv}% of ordered</div>
  </div>
  <div class="kpi" style="border-top-color:var(--orange)">
    <div class="kpi-label">Pending Receipt</div>
    <div class="kpi-val" style="color:var(--orange)">{pending:,}</div>
    <div class="kpi-sub">Units not yet received</div>
  </div>
  <div class="kpi" style="border-top-color:var(--teal)">
    <div class="kpi-label">Fulfillment Rate</div>
    <div class="kpi-val" style="color:var(--teal)">{pct_recv}%</div>
    <div class="kpi-sub">Received / Ordered</div>
  </div>
</div>

<!-- Charts -->
<div class="chart-row">
  <div class="card" style="display:flex;flex-direction:column;align-items:center;justify-content:center">
    <div class="section-title" style="width:100%">Shipment Status</div>
    <div id="chart_donut" style="width:100%"></div>
  </div>
  <div class="card">
    <div class="section-title">Ordered vs PL vs Received by PO</div>
    <div id="chart_bar"></div>
  </div>
  <div class="card">
    <div class="section-title">Fulfillment Gap by PO</div>
    <div id="chart_funnel"></div>
  </div>
</div>

<!-- Tabs -->
<div class="tabs">
  <div class="tab active" onclick="switchTab('sku',this)">🔍 SKU Tracker</div>
  <div class="tab" onclick="switchTab('po',this)">📋 PO Overview</div>
  <div class="tab" onclick="switchTab('model',this)">📈 By Model</div>
</div>

<!-- ════ TAB: SKU TRACKER ════ -->
<div id="tab-sku" class="tab-panel active">
  <div class="filter-bar">
    <div class="filter-group"><label>PO#</label>
      <select id="s_po" onchange="filterSKU()"><option value="">All</option>{po_options}</select></div>
    <div class="filter-group"><label>Ship Status</label>
      <select id="s_ship" onchange="filterSKU()"><option value="">All</option>{ship_options}</select></div>
    <div class="filter-group"><label>Model</label>
      <select id="s_model" onchange="filterSKU()"><option value="">All</option>{model_options}</select></div>
    <div class="filter-group"><label>SKU / Code</label>
      <input type="text" id="s_sku_q" placeholder="Search SKU..." oninput="filterSKU()"></div>
    <div class="filter-group"><label>Receipt Status</label>
      <select id="s_recv_status" onchange="filterSKU()">
        <option value="">All</option>
        <option value="full">✅ Fully Received</option>
        <option value="partial">🟡 Partially Received</option>
        <option value="pl_only">🔵 In PL, Not Received</option>
        <option value="none">🔴 Not in PL Yet</option>
      </select></div>
    <div class="filter-group"><label>Arrival From</label>
      <input type="text" id="s_arrive_from" placeholder="DD/MM/YYYY" style="min-width:120px" readonly></div>
    <div class="filter-group"><label>Arrival To</label>
      <input type="text" id="s_arrive_to" placeholder="DD/MM/YYYY" style="min-width:120px" readonly></div>
    <button class="btn btn-reset" onclick="resetSKU()">↺ Reset</button>
    <span id="sku_count" style="font-size:12px;color:var(--muted);align-self:center"></span>
  </div>
  <div class="dot-legend">
    <span><span class="dot dot-green">●</span> Fully received</span>
    <span><span class="dot dot-yellow">●</span> Partially received</span>
    <span><span class="dot dot-blue">●</span> PL complete, not received</span>
    <span><span class="dot dot-orange">●</span> Partial PL</span>
    <span><span class="dot dot-red">●</span> Not in PL yet</span>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>PO#</th><th>Status</th><th>Arrival</th>
        <th>SKU Code</th><th>Model</th><th>Variant</th>
        <th>Ordered</th><th>PL Qty</th><th>Not in PL</th>
        <th>Received</th><th>PL not recv</th>
        <th>Stage</th><th>% Received</th>
      </tr></thead>
      <tbody id="tb_sku">{sku_rows}</tbody>
    </table>
  </div>
</div>

<!-- ════ TAB: PO OVERVIEW ════ -->
<div id="tab-po" class="tab-panel">
  <div class="filter-bar">
    <div class="filter-group"><label>PO#</label>
      <select id="f_po" onchange="filterPO()"><option value="">All</option>{po_options}</select></div>
    <div class="filter-group"><label>Ship Status</label>
      <select id="f_ship" onchange="filterPO()"><option value="">All</option>{ship_options}</select></div>
    <div class="filter-group"><label>Arrival From</label>
      <input type="text" id="f_arrive_from" placeholder="DD/MM/YYYY" style="min-width:120px" readonly></div>
    <div class="filter-group"><label>Arrival To</label>
      <input type="text" id="f_arrive_to" placeholder="DD/MM/YYYY" style="min-width:120px" readonly></div>
    <button class="btn btn-reset" onclick="resetPO()">↺ Reset</button>
    <span id="po_count" style="font-size:12px;color:var(--muted);align-self:center"></span>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>PO#</th><th>Supplier</th><th>Status</th><th>Container</th><th>Model/SKU</th>
        <th>Order Date</th><th>Packed</th><th>Departed</th><th>Arrived</th>
        <th>Ordered</th><th>PL Qty</th><th>Not in PL</th>
        <th>Received</th><th>PL not recv</th><th>% Received</th>
      </tr></thead>
      <tbody id="tb_po">{po_rows}</tbody>
    </table>
  </div>
</div>

<!-- ════ TAB: MODEL ════ -->
<div id="tab-model" class="tab-panel">
  <div class="filter-bar">
    <div class="filter-group"><label>Search Model</label>
      <input type="text" id="f_model_q" placeholder="Type model name..." style="min-width:200px" oninput="filterModel()"></div>
    <div class="filter-group"><label>Receipt Status</label>
      <select id="f_model_pct" onchange="filterModel()">
        <option value="">All</option>
        <option value="full">Fully Received (100%)</option>
        <option value="partial">Partial (&gt;0% &lt;100%)</option>
        <option value="none">Not Received (0%)</option>
      </select></div>
    <span id="model_count" style="font-size:12px;color:var(--muted);align-self:center"></span>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>Model</th><th>SKUs</th><th>POs</th>
        <th>Ordered</th><th>PL Qty</th><th>Not in PL</th>
        <th>Received</th><th>PL not recv</th>
        <th>Stage</th><th>% Received</th>
      </tr></thead>
      <tbody id="tb_model">{model_rows}</tbody>
    </table>
  </div>
</div>

</div><!-- /wrap -->

<script>
const CHARTS = {{
  donut:  {chart_donut},
  bar:    {chart_bar},
  funnel: {chart_funnel},
}};
const CFG = {{responsive:true, displayModeBar:false}};

function renderChart(id, obj) {{
  Plotly.newPlot(id, obj.data, {{...obj.layout, responsive:true}}, CFG);
}}
renderChart('chart_donut',  CHARTS.donut);
renderChart('chart_bar',    CHARTS.bar);
renderChart('chart_funnel', CHARTS.funnel);

let resizeT;
window.addEventListener('resize', () => {{
  clearTimeout(resizeT);
  resizeT = setTimeout(() => {{
    ['chart_donut','chart_bar','chart_funnel'].forEach(id => Plotly.Plots.resize(id));
  }}, 200);
}});

// ── Tab switch ──
function switchTab(name, el) {{
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  el.classList.add('active');
  setTimeout(()=>['chart_donut','chart_bar','chart_funnel'].forEach(id=>Plotly.Plots.resize(id)),50);
}}

// ── Flatpickr date pickers ──
function makePicker(id, onClose) {{
  return flatpickr('#'+id, {{dateFormat:'d/m/Y', allowInput:false,
    locale:{{firstDayOfWeek:1}}, onClose}});
}}
const pickerSFrom = makePicker('s_arrive_from', filterSKU);
const pickerSTo   = makePicker('s_arrive_to',   filterSKU);
const pickerPFrom = makePicker('f_arrive_from', filterPO);
const pickerPTo   = makePicker('f_arrive_to',   filterPO);

function parseDate(str) {{
  if (!str) return null;
  const [d,m,y] = str.split('/');
  return new Date(y, m-1, d);
}}

// ── SKU filter ──
function filterSKU() {{
  const po     = document.getElementById('s_po').value;
  const ship   = document.getElementById('s_ship').value;
  const model  = document.getElementById('s_model').value;
  const q      = document.getElementById('s_sku_q').value.toLowerCase();
  const recv   = document.getElementById('s_recv_status').value;
  const fromDt = parseDate(document.getElementById('s_arrive_from').value);
  const toDt   = parseDate(document.getElementById('s_arrive_to').value);
  const rows   = document.querySelectorAll('#tb_sku .trow');
  let vis = 0;
  rows.forEach(r => {{
    let show = true;
    if (po    && r.getAttribute('data-po')    !== po)    show=false;
    if (ship  && r.getAttribute('data-ship')  !== ship)  show=false;
    if (model && r.getAttribute('data-model') !== model) show=false;
    if (q     && !r.getAttribute('data-sku').toLowerCase().includes(q)) show=false;
    if (fromDt || toDt) {{
      const arr = r.getAttribute('data-arrive');
      if (arr) {{
        const d = new Date(arr);
        if (fromDt && d < fromDt) show=false;
        if (toDt   && d > toDt)   show=false;
      }}
    }}
    if (recv) {{
      const dot = r.querySelector('.dot');
      const title = dot ? dot.getAttribute('title') : '';
      if (recv==='full'     && title !== 'Fully received')           show=false;
      if (recv==='partial'  && title !== 'Partially received')       show=false;
      if (recv==='pl_only'  && title !== 'PL complete, not received') show=false;
      if (recv==='none'     && title !== 'Not in PL yet' && title !== 'Partial PL') show=false;
    }}
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  document.getElementById('sku_count').textContent = `Showing ${{vis}} / ${{rows.length}} items`;
}}
function resetSKU() {{
  ['s_po','s_ship','s_model','s_recv_status'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('s_sku_q').value='';
  pickerSFrom.clear(); pickerSTo.clear();
  filterSKU();
}}
filterSKU();

// ── PO filter ──
function filterPO() {{
  const po   = document.getElementById('f_po').value;
  const ship = document.getElementById('f_ship').value;
  const fromDt = parseDate(document.getElementById('f_arrive_from').value);
  const toDt   = parseDate(document.getElementById('f_arrive_to').value);
  const rows = document.querySelectorAll('#tb_po .trow');
  let vis = 0;
  rows.forEach(r => {{
    let show = true;
    if (po   && r.getAttribute('data-po')   !== po)   show=false;
    if (ship && r.getAttribute('data-ship') !== ship) show=false;
    if (fromDt || toDt) {{
      const arr = r.getAttribute('data-arrive');
      if (arr) {{
        const d = new Date(arr);
        if (fromDt && d < fromDt) show=false;
        if (toDt   && d > toDt)   show=false;
      }}
    }}
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  document.getElementById('po_count').textContent = `Showing ${{vis}} / ${{rows.length}} items`;
}}
function resetPO() {{
  ['f_po','f_ship'].forEach(id=>document.getElementById(id).value='');
  pickerPFrom.clear(); pickerPTo.clear();
  filterPO();
}}
filterPO();

// ── Model filter ──
function filterModel() {{
  const q   = document.getElementById('f_model_q').value.toLowerCase();
  const pct = document.getElementById('f_model_pct').value;
  const rows = document.querySelectorAll('#tb_model .trow');
  let vis = 0;
  rows.forEach(r => {{
    const name = r.getAttribute('data-model').toLowerCase();
    const pbar = r.querySelector('.pbar-label');
    const pval = pbar ? parseFloat(pbar.textContent) : 0;
    let show = name.includes(q);
    if (pct==='full'    && pval < 100)              show=false;
    if (pct==='partial' && (pval===0||pval>=100))   show=false;
    if (pct==='none'    && pval > 0)                show=false;
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  document.getElementById('model_count').textContent = `Showing ${{vis}} / ${{rows.length}} items`;
}}
filterModel();
</script>
</body>
</html>"""

with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"✅  Saved → {OUT_PATH}")
print(f"📌  Open in browser — {Path(EXCEL_PATH).name} → {OUT_PATH}")
"""
Honeywell Binstock Program Scorecard — Dashboard Generator (Urbana)
-------------------------------------------------------------------
GitLab Portability Mode: paths resolve relative to this script's location.

Run:    python generate_dashboard_urbana.py
Output: index.html  (saved to same folder as this script)

Requirements:
    pip install pandas openpyxl
"""

import pandas as pd
import os
import sys
from datetime import datetime

# ── Resolve paths relative to this script's location ────────────────────────
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
# ────────────────────────────────────────────────────────────────────────────

# ── CONFIG ──────────────────────────────────────────────────────────────────
XLSX_PATH   = os.path.join(script_dir, "Honeywell Urbana_042026_binstrat.xlsx")
SHEET_NAME  = "Bin Map Rpt_Urbana"
OUTPUT_FILE = os.path.join(script_dir, "index.html")
# ────────────────────────────────────────────────────────────────────────────

print("Honeywell Dashboard Refresh (GitLab Portability Mode)")
print("=" * 52)
print()
print("Loading Excel file...")

if not os.path.exists(XLSX_PATH):
    print(f"ERROR: Excel file not found in the project folder.")
    print(f"Expected filename: {os.path.basename(XLSX_PATH)}")
    sys.exit(1)


def load_and_calculate(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)

    total    = len(df)
    active   = len(df[df['Bin Activity Status'] == 'Active'])
    inactive = len(df[df['Bin Activity Status'] == 'Inactive'])

    stockout_total  = len(df[df['Stockout Status'] == 'STOCKOUT'])
    stockout_active = len(df[(df['Stockout Status'] == 'STOCKOUT') & (df['Bin Activity Status'] == 'Active')])

    fill_total  = round((total  - stockout_total)  / total  * 100, 2)
    fill_active = round((active - stockout_active) / active * 100, 2)

    past_due_total  = len(df[df['Past Due?'] == 'Yes'])
    past_due_active = len(df[(df['Past Due?'] == 'Yes') & (df['Bin Activity Status'] == 'Active')])
    pd_pct_total    = round(past_due_total  / total  * 100, 2)
    pd_pct_active   = round(past_due_active / active * 100, 2)
    pd_risk_delta   = round((pd_pct_active - pd_pct_total) / pd_pct_total * 100, 0) if pd_pct_total else 0

    on_priced    = len(df[df['Contract Status'] == 'On-Contract : Priced'])
    off_contract = len(df[df['Contract Status'] == 'Off-Contract'])
    unpriced     = len(df[df['Contract Status'] == 'On-Contract : Unpriced'])

    df_active            = df[df['Bin Activity Status'] == 'Active']
    on_priced_active     = len(df_active[df_active['Contract Status'] == 'On-Contract : Priced'])
    off_contract_active  = len(df_active[df_active['Contract Status'] == 'Off-Contract'])

    on_contract_pct         = round(on_priced           / total  * 100, 1)
    off_contract_pct        = round(off_contract        / total  * 100, 2)
    unpriced_pct            = round(unpriced            / total  * 100, 2)
    on_contract_active_pct  = round(on_priced_active    / active * 100, 1)
    off_contract_active_pct = round(off_contract_active / active * 100, 2)

    flag_delete = len(df[df['Action'] == 'DELETE'])
    flag_review = len(df[df['Action'] == 'Move to PO/BOM Review Required'])

    active_pct   = round(active   / total * 100, 1)
    inactive_pct = round(inactive / total * 100, 1)

    stockout_pct_total  = round(stockout_total  / total  * 100, 2)
    stockout_pct_active = round(stockout_active / active * 100, 2)
    stocked_total       = total  - stockout_total
    stocked_active      = active - stockout_active

    C = 263.9
    def dash(pct):   return f"{round(C * pct / 100, 1)} {round(C - C * pct / 100, 1)}"
    def offset(pct): return f"-{round(C * pct / 100, 1)}"

    return dict(
        total=f"{total:,}", active=f"{active:,}", inactive=f"{inactive:,}",
        active_pct=active_pct, inactive_pct=inactive_pct,
        stockout_total=stockout_total, stockout_active=stockout_active,
        stockout_pct_total=stockout_pct_total, stockout_pct_active=stockout_pct_active,
        stocked_total=f"{stocked_total:,}", stocked_active=f"{stocked_active:,}",
        fill_total=fill_total, fill_active=fill_active,
        past_due_total=past_due_total, past_due_active=past_due_active,
        pd_pct_total=pd_pct_total, pd_pct_active=pd_pct_active,
        pd_risk_delta=int(pd_risk_delta),
        on_priced=f"{on_priced:,}", off_contract=off_contract, unpriced=unpriced,
        on_contract_pct=on_contract_pct, off_contract_pct=off_contract_pct, unpriced_pct=unpriced_pct,
        on_priced_active=f"{on_priced_active:,}", off_contract_active=off_contract_active,
        on_contract_active_pct=on_contract_active_pct, off_contract_active_pct=off_contract_active_pct,
        flag_delete=f"{flag_delete:,}", flag_review=f"{flag_review:,}",
        arc_priced_total=dash(on_contract_pct),
        arc_off_total=dash(off_contract_pct),
        off_offset_total=offset(on_contract_pct),
        arc_unpriced_total=dash(unpriced_pct),
        unpriced_offset_total=offset(on_contract_pct + off_contract_pct),
        arc_priced_active=dash(on_contract_active_pct),
        arc_off_active=dash(off_contract_active_pct),
        off_offset_active=offset(on_contract_active_pct),
        report_date=datetime.now().strftime("%B %Y"),
        file_name=os.path.basename(path),
    )


def build_html(d):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Honeywell Binstock Program Scorecard</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@400;600;700;800&display=swap');

  :root {{
    --bg:      #eef1f6;
    --surface: #ffffff;
    --border:  #c8d1de;
    --accent:  #025f99;
    --green:   #15803d;
    --yellow:  #b45309;
    --red:     #b91c1c;
    --purple:  #6d28d9;
    --muted:   #374151;
    --text:    #0f172a;
    --subtext: #1e293b;
  }}

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    background: var(--bg);
    color: var(--text);
    font-family: 'Inter', sans-serif;
    font-weight: 600;
    min-height: 100vh;
    padding: 0;
    overflow-x: hidden;
  }}

  .header {{
    background: linear-gradient(135deg, #f0f4f8 0%, #e1eaf5 100%);
    border-bottom: 1px solid var(--border);
    padding: 28px 40px 24px;
    display: flex; justify-content: space-between; align-items: flex-end;
    position: relative; overflow: hidden;
  }}
  .header::before {{
    content: ''; position: absolute; top: -60px; right: -60px;
    width: 260px; height: 260px;
    background: radial-gradient(circle, rgba(3,105,161,0.07) 0%, transparent 70%);
  }}
  .header-left {{ display: flex; flex-direction: column; gap: 4px; }}
  .header-eyebrow {{
    font-size: 11px; font-weight: 700; letter-spacing: 0.12em;
    text-transform: uppercase; color: var(--accent); opacity: 0.8;
  }}
  .header-title {{
    font-family: 'Syne', sans-serif;
    font-size: 26px; font-weight: 800;
    color: var(--text); line-height: 1.1;
  }}
  .header-sub {{
    font-size: 13px; font-weight: 500; color: #475569; margin-top: 2px;
  }}
  .header-right {{ text-align: right; }}
  .header-date {{
    font-size: 13px; font-weight: 700; color: var(--accent);
    letter-spacing: 0.04em;
  }}
  .header-source {{
    font-size: 11px; font-weight: 500; color: #64748b; margin-top: 3px;
  }}

  .main {{ padding: 28px 40px; display: flex; flex-direction: column; gap: 28px; }}

  .section-header {{
    display: flex; align-items: center; gap: 12px; margin-bottom: 0;
  }}
  .section-title {{
    font-family: 'Syne', sans-serif;
    font-size: 13px; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase; color: var(--accent); white-space: nowrap;
  }}
  .section-line {{
    flex: 1; height: 1px; background: var(--border);
  }}

  .row {{ display: grid; gap: 16px; margin-top: 12px; }}
  .row-4 {{ grid-template-columns: repeat(4, 1fr); }}
  .row-3 {{ grid-template-columns: repeat(3, 1fr); }}
  .row-2 {{ grid-template-columns: repeat(2, 1fr); }}

  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 18px 20px;
    position: relative;
  }}
  .card-label {{
    font-size: 11px; font-weight: 700; letter-spacing: 0.09em;
    text-transform: uppercase; color: #64748b; margin-bottom: 10px;
  }}

  .kpi-val {{
    font-family: 'Syne', sans-serif;
    font-size: 36px; font-weight: 800; line-height: 1;
  }}
  .kpi-sub {{
    font-size: 12px; font-weight: 500; color: #64748b; margin-top: 4px;
  }}
  .kpi-delta {{
    display: inline-flex; align-items: center; gap: 4px;
    font-size: 11px; font-weight: 700; margin-top: 6px;
    padding: 2px 7px; border-radius: 20px;
  }}
  .delta-up   {{ background: #dcfce7; color: #15803d; }}
  .delta-down {{ background: #fee2e2; color: #b91c1c; }}
  .delta-warn {{ background: #fef9c3; color: #b45309; }}

  .c-green  {{ color: var(--green); }}
  .c-red    {{ color: var(--red); }}
  .c-yellow {{ color: var(--yellow); }}
  .c-accent {{ color: var(--accent); }}
  .c-muted  {{ color: var(--muted); }}

  .fill-bar-wrap {{ margin-top: 10px; }}
  .fill-bar-label {{
    display: flex; justify-content: space-between;
    font-size: 11px; font-weight: 600; color: #64748b; margin-bottom: 4px;
  }}
  .fill-bar-bg {{
    height: 8px; background: #e2e8f0; border-radius: 99px; overflow: hidden;
  }}
  .fill-bar-fill {{
    height: 100%; border-radius: 99px;
    transition: width 0.8s cubic-bezier(0.4,0,0.2,1);
  }}
  .fill-green  {{ background: var(--green); }}
  .fill-yellow {{ background: var(--yellow); }}
  .fill-red    {{ background: var(--red); }}
  .fill-accent {{ background: var(--accent); }}

  .activity-visual {{
    display: flex; flex-wrap: wrap; gap: 3px; margin: 10px 0;
  }}
  .dot-bin {{
    width: 10px; height: 10px; border-radius: 2px;
  }}
  .dot-active   {{ background: var(--green); opacity: 0.85; }}
  .dot-inactive {{ background: #cbd5e1; }}

  .activity-legend {{ display: flex; flex-direction: column; gap: 8px; margin-top: 4px; }}
  .act-item  {{ display: flex; align-items: flex-start; gap: 8px; }}
  .act-dot   {{ width: 10px; height: 10px; border-radius: 2px; margin-top: 3px; flex-shrink: 0; }}
  .bg-muted  {{ background: #cbd5e1; }}
  .act-num   {{ font-family: 'Syne', sans-serif; font-size: 20px; font-weight: 800; }}
  .act-pct   {{ font-size: 12px; font-weight: 600; color: #64748b; }}
  .act-label {{ font-size: 11px; font-weight: 500; color: #64748b; margin-top: 1px; }}

  .contract-layout {{
    display: flex; align-items: center; gap: 16px;
  }}
  .donut-wrap {{ position: relative; flex-shrink: 0; }}
  .donut-center {{
    position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
    text-align: center;
  }}
  .donut-center-val {{
    font-family: 'Syne', sans-serif; font-size: 15px; font-weight: 800;
    color: var(--text); line-height: 1;
  }}
  .donut-center-lbl {{
    font-size: 9px; font-weight: 600; color: #64748b;
    text-transform: uppercase; letter-spacing: 0.05em;
  }}
  .contract-legend {{ display: flex; flex-direction: column; gap: 7px; flex: 1; }}
  .legend-row {{
    display: flex; align-items: center; gap: 7px;
    font-size: 12px; font-weight: 600;
  }}
  .legend-dot {{ width: 9px; height: 9px; border-radius: 50%; flex-shrink: 0; }}
  .bg-green  {{ background: var(--green); }}
  .bg-red    {{ background: var(--red); }}
  .bg-yellow {{ background: var(--yellow); }}
  .legend-name {{ flex: 1; color: var(--subtext); font-size: 11px; }}
  .legend-count {{ font-family: 'Syne', sans-serif; font-size: 14px; font-weight: 800; min-width: 38px; text-align: right; }}
  .legend-pct   {{ font-size: 11px; font-weight: 600; color: #64748b; min-width: 36px; text-align: right; }}

  .risk-row {{
    display: flex; align-items: center; gap: 8px; font-size: 12px;
  }}
  .risk-label  {{ width: 160px; font-weight: 600; color: var(--subtext); flex-shrink: 0; font-size: 11px; }}
  .risk-bar-wrap {{ flex: 1; }}
  .risk-bar-bg {{
    height: 7px; background: #e2e8f0; border-radius: 99px; overflow: hidden;
  }}
  .risk-bar-inner {{
    height: 100%; border-radius: 99px;
    transition: width 0.8s cubic-bezier(0.4,0,0.2,1);
  }}
  .risk-value {{ min-width: 38px; text-align: right; font-weight: 700; font-size: 12px; }}
  .risk-count {{ min-width: 52px; text-align: right; font-size: 11px; font-weight: 500; }}

  .disclosure {{
    margin: 0 40px 0; padding: 12px 18px;
    background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
    display: flex; align-items: flex-start; gap: 10px;
    font-size: 11px; font-weight: 500; color: #475569; line-height: 1.5;
  }}
  .disc-icon  {{ font-size: 15px; color: var(--accent); flex-shrink: 0; margin-top: 1px; }}
  .disc-prop  {{ font-weight: 700; color: var(--text); margin-right: 6px; }}
  .disc-divider {{ width: 1px; background: #cbd5e1; align-self: stretch; margin: 0 4px; }}

  .footer {{
    margin: 12px 40px 28px;
    display: flex; justify-content: space-between;
    font-size: 11px; font-weight: 500; color: #94a3b8;
  }}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <div class="header-eyebrow">Boeing Distribution Services · Program Management</div>
    <div class="header-title">Honeywell Binstock Program Scorecard</div>
    <div class="header-sub">Urbana · Bin Map Analysis &amp; Risk Summary</div>
  </div>
  <div class="header-right">
    <div class="header-date">{d['report_date']}</div>
    <div class="header-source">Source: {d['file_name']}</div>
  </div>
</div>

<div class="main">

  <div>
    <div class="section-header"><span class="section-title">Bin Map Overview</span><div class="section-line"></div></div>
    <div class="row row-4">
      <div class="card">
        <div class="card-label">Total Bins</div>
        <div class="kpi-val c-accent">{d['total']}</div>
        <div class="kpi-sub">Full bin map scope</div>
      </div>
      <div class="card">
        <div class="card-label">Active Bins</div>
        <div class="kpi-val c-green">{d['active']}</div>
        <div class="kpi-sub">{d['active_pct']}% of total · scanned 2023–2026</div>
      </div>
      <div class="card">
        <div class="card-label">Inactive Bins</div>
        <div class="kpi-val" style="color:var(--subtext);">{d['inactive']}</div>
        <div class="kpi-sub">{d['inactive_pct']}% of total · zero scans 3 yrs</div>
      </div>
      <div class="card">
        <div class="card-label">Past Due (Active)</div>
        <div class="kpi-val c-red">{d['past_due_active']}</div>
        <div class="kpi-sub">{d['pd_pct_active']}% of active bins
          <span class="kpi-delta delta-{'down' if d['pd_risk_delta'] > 0 else 'up'}">{'+' if d['pd_risk_delta'] > 0 else ''}{d['pd_risk_delta']}% vs total</span>
        </div>
      </div>
    </div>
  </div>

  <div>
    <div class="section-header"><span class="section-title">Fill Rate &amp; Stockout</span><div class="section-line"></div></div>
    <div class="row row-3">
      <div class="card">
        <div class="card-label">Fill Rate — Total Bin Map</div>
        <div class="kpi-val c-green">{d['fill_total']}%</div>
        <div class="kpi-sub">{d['stocked_total']} stocked of {d['total']} total bins</div>
        <div class="fill-bar-wrap">
          <div class="fill-bar-label"><span>Fill</span><span>{d['fill_total']}%</span></div>
          <div class="fill-bar-bg"><div class="fill-bar-fill fill-green" style="width:{d['fill_total']}%;"></div></div>
        </div>
      </div>
      <div class="card">
        <div class="card-label">Fill Rate — Active Bins Only</div>
        <div class="kpi-val c-green">{d['fill_active']}%</div>
        <div class="kpi-sub">{d['stocked_active']} stocked of {d['active']} active bins</div>
        <div class="fill-bar-wrap">
          <div class="fill-bar-label"><span>Fill</span><span>{d['fill_active']}%</span></div>
          <div class="fill-bar-bg"><div class="fill-bar-fill fill-green" style="width:{d['fill_active']}%;"></div></div>
        </div>
      </div>
      <div class="card">
        <div class="card-label">Stockout Exposure</div>
        <div style="display:flex; gap:20px; margin-top:4px;">
          <div>
            <div class="kpi-val c-red" style="font-size:28px;">{d['stockout_total']}</div>
            <div class="kpi-sub">Total map · {d['stockout_pct_total']}%</div>
          </div>
          <div>
            <div class="kpi-val c-red" style="font-size:28px;">{d['stockout_active']}</div>
            <div class="kpi-sub">Active only · {d['stockout_pct_active']}%</div>
          </div>
        </div>
        <div class="fill-bar-wrap" style="margin-top:14px;">
          <div class="fill-bar-label"><span>Stockout (Active)</span><span>{d['stockout_pct_active']}%</span></div>
          <div class="fill-bar-bg"><div class="fill-bar-fill fill-red" style="width:{d['stockout_pct_active']}%;"></div></div>
        </div>
      </div>
    </div>
  </div>

  <div>
    <div class="section-header"><span class="section-title">Bin Activity</span><div class="section-line"></div></div>
    <div class="row row-2">
      <div class="card">
        <div class="card-label">Past Due Breakdown</div>
        <div style="display:flex; gap:24px; margin-top:6px;">
          <div>
            <div class="kpi-val c-red" style="font-size:28px;">{d['past_due_total']}</div>
            <div class="kpi-sub">Total map · {d['pd_pct_total']}%</div>
          </div>
          <div>
            <div class="kpi-val c-red" style="font-size:28px;">{d['past_due_active']}</div>
            <div class="kpi-sub">Active only · {d['pd_pct_active']}%</div>
          </div>
        </div>
        <div class="fill-bar-wrap" style="margin-top:14px;">
          <div class="fill-bar-label"><span>Past Due — Total</span><span>{d['pd_pct_total']}%</span></div>
          <div class="fill-bar-bg"><div class="fill-bar-fill fill-yellow" style="width:{d['pd_pct_total']}%;"></div></div>
        </div>
        <div class="fill-bar-wrap" style="margin-top:8px;">
          <div class="fill-bar-label"><span>Past Due — Active</span><span>{d['pd_pct_active']}%</span></div>
          <div class="fill-bar-bg"><div class="fill-bar-fill fill-red" style="width:{d['pd_pct_active']}%;"></div></div>
        </div>
      </div>
      <div class="card">
        <div class="card-label">Bin Activity Breakdown</div>
        <div class="activity-visual" id="activityDots"></div>
        <div class="activity-legend">
          <div class="act-item">
            <div class="act-dot" style="background:var(--green);"></div>
            <div>
              <div style="display:flex; align-items:baseline; gap:5px;">
                <span class="act-num c-green">{d['active']}</span>
                <span class="act-pct">({d['active_pct']}%)</span>
              </div>
              <div class="act-label">Active — scanned in last 3 yrs</div>
            </div>
          </div>
          <div class="act-item">
            <div class="act-dot bg-muted"></div>
            <div>
              <div style="display:flex; align-items:baseline; gap:5px;">
                <span class="act-num" style="color:var(--subtext);">{d['inactive']}</span>
                <span class="act-pct">({d['inactive_pct']}%)</span>
              </div>
              <div class="act-label">Inactive — zero scans in 3 yrs</div>
            </div>
          </div>
        </div>
        <div style="margin-top:10px; font-size:13px; font-weight:600; color:var(--muted); line-height:1.6; border-top:1px solid var(--border); padding-top:10px;">
          <span style="color:var(--yellow);">{d['flag_delete']} bins flagged DELETE</span> · {d['flag_review']} flagged for PO/BOM Review — inactive population represents significant footprint recapture opportunity.
        </div>
      </div>
    </div>
  </div>

  <div>
    <div class="section-header"><span class="section-title">Contract Status</span><div class="section-line"></div></div>
    <div style="margin-top:12px;" class="row row-3">

      <div class="card">
        <div class="card-label">Contract Status — Total Bin Map</div>
        <div class="contract-layout" style="margin-top:10px;">
          <div class="donut-wrap">
            <svg width="110" height="110" viewBox="0 0 110 110">
              <circle cx="55" cy="55" r="42" fill="none" stroke="rgba(0,0,0,0.07)" stroke-width="14"/>
              <circle cx="55" cy="55" r="42" fill="none" stroke="#15803d" stroke-width="14" stroke-dasharray="{d['arc_priced_total']}" stroke-dashoffset="0" transform="rotate(-90 55 55)" opacity="0.85"/>
              <circle cx="55" cy="55" r="42" fill="none" stroke="#b91c1c" stroke-width="14" stroke-dasharray="{d['arc_off_total']}" stroke-dashoffset="{d['off_offset_total']}" transform="rotate(-90 55 55)" opacity="0.9"/>
              <circle cx="55" cy="55" r="42" fill="none" stroke="#b45309" stroke-width="14" stroke-dasharray="{d['arc_unpriced_total']}" stroke-dashoffset="{d['unpriced_offset_total']}" transform="rotate(-90 55 55)" opacity="0.9"/>
            </svg>
            <div class="donut-center">
              <div class="donut-center-val">{d['on_contract_pct']}%</div>
              <div class="donut-center-lbl">On-Contract</div>
            </div>
          </div>
          <div class="contract-legend">
            <div class="legend-row"><div class="legend-dot bg-green"></div><div class="legend-name">On-Contract · Priced</div><span class="legend-count c-green">{d['on_priced']}</span><span class="legend-pct">{d['on_contract_pct']}%</span></div>
            <div class="legend-row"><div class="legend-dot bg-red"></div><div class="legend-name">Off-Contract</div><span class="legend-count c-red">{d['off_contract']}</span><span class="legend-pct">{d['off_contract_pct']}%</span></div>
            <div class="legend-row"><div class="legend-dot bg-yellow"></div><div class="legend-name">On-Contract · Unpriced</div><span class="legend-count c-yellow">{d['unpriced']}</span><span class="legend-pct">{d['unpriced_pct']}%</span></div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-label">Contract Status — Active Bins Only</div>
        <div class="contract-layout" style="margin-top:10px;">
          <div class="donut-wrap">
            <svg width="110" height="110" viewBox="0 0 110 110">
              <circle cx="55" cy="55" r="42" fill="none" stroke="rgba(0,0,0,0.07)" stroke-width="14"/>
              <circle cx="55" cy="55" r="42" fill="none" stroke="#15803d" stroke-width="14" stroke-dasharray="{d['arc_priced_active']}" stroke-dashoffset="0" transform="rotate(-90 55 55)" opacity="0.85"/>
              <circle cx="55" cy="55" r="42" fill="none" stroke="#b91c1c" stroke-width="14" stroke-dasharray="{d['arc_off_active']}" stroke-dashoffset="{d['off_offset_active']}" transform="rotate(-90 55 55)" opacity="0.9"/>
            </svg>
            <div class="donut-center">
              <div class="donut-center-val">{d['on_contract_active_pct']}%</div>
              <div class="donut-center-lbl">On-Contract</div>
            </div>
          </div>
          <div class="contract-legend">
            <div class="legend-row"><div class="legend-dot bg-green"></div><div class="legend-name">On-Contract · Priced</div><span class="legend-count c-green">{d['on_priced_active']}</span><span class="legend-pct">{d['on_contract_active_pct']}%</span></div>
            <div class="legend-row"><div class="legend-dot bg-red"></div><div class="legend-name">Off-Contract</div><span class="legend-count c-red">{d['off_contract_active']}</span><span class="legend-pct">{d['off_contract_active_pct']}%</span></div>
            <div class="legend-row" style="opacity:0.4;"><div class="legend-dot" style="background:var(--muted);"></div><div class="legend-name">Unpriced (in inactive)</div><span class="legend-count">—</span></div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-label">Contract Risk Summary</div>
        <div style="margin-top:12px; display:flex; flex-direction:column; gap:10px;">
          <div class="risk-row"><div class="risk-label">Off-Contract (Total Map)</div><div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-red" style="width:{d['off_contract_pct']}%;"></div></div></div><div class="risk-value c-red">{d['off_contract_pct']}%</div><div class="risk-count c-muted">{d['off_contract']} bins</div></div>
          <div class="risk-row"><div class="risk-label">Off-Contract (Active Only)</div><div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-red" style="width:{d['off_contract_active_pct']}%;"></div></div></div><div class="risk-value c-red">{d['off_contract_active_pct']}%</div><div class="risk-count c-muted">{d['off_contract_active']} bins</div></div>
          <div class="risk-row"><div class="risk-label">Unpriced (Total Map)</div><div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-yellow" style="width:{d['unpriced_pct']}%;"></div></div></div><div class="risk-value c-yellow">{d['unpriced_pct']}%</div><div class="risk-count c-muted">{d['unpriced']} bins</div></div>
          <div class="risk-row"><div class="risk-label">Past Due (Active Lens)</div><div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-red" style="width:{d['pd_pct_active']}%;"></div></div></div><div class="risk-value c-red">{d['pd_pct_active']}%</div><div class="risk-count c-muted">{d['past_due_active']} bins</div></div>
          <div class="risk-row"><div class="risk-label">Stockout (Active Lens)</div><div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-red" style="width:{d['stockout_pct_active']}%;"></div></div></div><div class="risk-value c-red">{d['stockout_pct_active']}%</div><div class="risk-count c-muted">{d['stockout_active']} bins</div></div>
        </div>
      </div>

    </div>
  </div>

</div>

<div class="disclosure">
  <div class="disc-icon">&#9432;</div>
  <div class="disc-text">
    <span class="disc-prop">Proprietary</span>
    The information contained herein is proprietary to The Boeing Company and shall not be reproduced or disclosed in whole or in part except when such user possesses direct written authorization from The Boeing Company.
  </div>
  <div class="disc-divider"></div>
  <div class="disc-text">
    The statements contained herein are based on good faith assumptions and are to be used for general information purposes only. These statements do not constitute an offer, promise, warranty or guarantee of performance.
  </div>
</div>

<div class="footer">
  <span>Boeing Distribution Services · Program Management · Honeywell Aerospace Account</span>
  <span>Data: {{d['file_name']}} · {{d['report_date']}} · Active = any scan 2023–2026 · Inactive = 0 scans same period</span>
</div>

<script>
  const container = document.getElementById('activityDots');
  const activeDots = Math.round(({d['active_pct']} / 100) * 150);
  for (let i = 0; i < 150; i++) {{
    const dot = document.createElement('div');
    dot.className = 'dot-bin ' + (i < activeDots ? 'dot-active' : 'dot-inactive');
    container.appendChild(dot);
  }}
  window.addEventListener('load', () => {{
    document.querySelectorAll('.fill-bar-fill').forEach(el => {{
      const w = el.style.width;
      el.style.width = '0%';
      requestAnimationFrame(() => {{ setTimeout(() => {{ el.style.width = w; }}, 100); }});
    }});
  }});
</script>

</body>
</html>"""


if __name__ == "__main__":
    print(f"Reading: {XLSX_PATH}")
    data = load_and_calculate(XLSX_PATH, SHEET_NAME)
    html = build_html(data)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Dashboard generated: {OUTPUT_FILE}")

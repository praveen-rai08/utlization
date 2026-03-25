import openpyxl
from collections import defaultdict

# Working days per month 2026
working_days = {
    "Jan'26": 21, "Feb'26": 20, "Mar'26": 22,
    "Apr'26": 22, "May'26": 21, "Jun'26": 22,
}
months_order = list(working_days.keys())

# Read source data
wb_src = openpyxl.load_workbook(
    r'C:\Users\202294\Downloads\QEA-UHG-Leave-Forecast-2026.xlsx',
    data_only=True
)

sheet_months = [
    ('2026-Jan-Feb-Mar', [("Jan'26", 14, 15), ("Feb'26", 18, 19), ("Mar'26", 22, 23)]),
    ('2026-Apr-May-Jun', [("Apr'26", 14, 15), ("May'26", 18, 19), ("Jun'26", 22, 23)]),
]

all_records = {}

for sheet_name, month_configs in sheet_months:
    ws = wb_src[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    for row in rows[2:]:
        if not any(row):
            continue
        assoc_id = row[1]
        if not assoc_id:
            continue
        if assoc_id not in all_records:
            all_records[assoc_id] = {
                'Associate ID': assoc_id,
                'Associate Name': row[2] or '',
                'Grade': row[3] or '',
                'Account': row[7] or '',
                'Project': row[5] or '',
                'Billability': row[10] or '',
                'Country': row[11] or '',
                'Onsite/Offshore': row[12] or '',
                'City': row[13] or '',
                'months': {}
            }
        for month_name, fc_idx, act_idx in month_configs:
            forecast = row[fc_idx]
            actual   = row[act_idx]
            wd       = working_days[month_name]
            try:
                leave = float(actual) if actual is not None else (float(forecast) if forecast is not None else 0.0)
            except:
                leave = 0.0
            try:
                forecast_val = float(forecast) if forecast is not None else 0.0
            except:
                forecast_val = 0.0
            try:
                actual_val = float(actual) if actual is not None else None
            except:
                actual_val = None
            available = wd - leave
            util_pct  = round((available / wd) * 100, 1) if wd > 0 else 0.0
            all_records[assoc_id]['months'][month_name] = {
                'wd': wd,
                'forecast': forecast_val,
                'actual': actual_val,
                'leave': leave,
                'available': available,
                'util': util_pct,
            }

records = list(all_records.values())

def avg_util(rec):
    vals = [rec['months'][m]['util'] for m in months_order if m in rec['months']]
    return round(sum(vals)/len(vals), 1) if vals else 0

# Aggregate stats
total_assoc    = len(records)
avg_utils      = [avg_util(r) for r in records]
overall_avg    = round(sum(avg_utils)/len(avg_utils), 1) if avg_utils else 0
low_count      = sum(1 for u in avg_utils if u < 80)
med_count      = sum(1 for u in avg_utils if 80 <= u < 90)
high_count     = sum(1 for u in avg_utils if u >= 90)

total_forecast = sum(r['months'].get(m, {}).get('forecast', 0) for r in records for m in months_order)
total_actual   = sum(r['months'].get(m, {}).get('actual', 0) or 0 for r in records for m in months_order)

# Monthly data
monthly_stats = []
for m in months_order:
    wd  = working_days[m]
    mf  = round(sum(r['months'].get(m, {}).get('forecast', 0) for r in records), 1)
    ma  = round(sum(r['months'].get(m, {}).get('actual', 0) or 0 for r in records), 1)
    ml  = round(sum(r['months'].get(m, {}).get('leave', 0) for r in records), 1)
    mu  = round(sum(r['months'].get(m, {}).get('util', 0) for r in records) / len(records), 1)
    monthly_stats.append({'month': m, 'wd': wd, 'forecast': mf, 'actual': ma, 'leave': ml, 'util': mu})

# Account rollup
acct_data = defaultdict(list)
for r in records:
    acct_data[r['Account']].append(r)

acct_rollup = []
for acct, recs in sorted(acct_data.items(), key=lambda x: -len(x[1])):
    n     = len(recs)
    avg_u = round(sum(avg_util(r) for r in recs)/n, 1)
    month_utils = {m: round(sum(r['months'].get(m, {}).get('util', 0) for r in recs)/n, 1) for m in months_order}
    acct_rollup.append({'account': acct, 'count': n, 'avg_util': avg_u, 'months': month_utils})

# Grade rollup
grade_data = defaultdict(list)
for r in records:
    grade_data[r['Grade']].append(r)

grade_order = ['A','SA','M','SM','AD','PA','PAT','Cont']
all_grades  = grade_order + [g for g in grade_data if g not in grade_order]
grade_rollup = []
for g in [x for x in all_grades if x in grade_data]:
    recs  = grade_data[g]
    n     = len(recs)
    avg_u = round(sum(avg_util(r) for r in recs)/n, 1)
    grade_rollup.append({'grade': g, 'count': n, 'avg_util': avg_u})

# Country rollup
country_data = defaultdict(list)
for r in records:
    country_data[r['Country']].append(r)

country_rollup = []
for country, recs in sorted(country_data.items(), key=lambda x: -len(x[1])):
    n     = len(recs)
    avg_u = round(sum(avg_util(r) for r in recs)/n, 1)
    country_rollup.append({'country': country, 'count': n, 'avg_util': avg_u})

# Low util employees
low_util_employees = sorted(
    [r for r in records if avg_util(r) < 90],
    key=lambda r: avg_util(r)
)[:20]

# Top leave takers (Jan+Feb actuals available)
top_leave = sorted(
    records,
    key=lambda r: sum(r['months'].get(m, {}).get('leave', 0) for m in ["Jan'26", "Feb'26"]),
    reverse=True
)[:10]

# JS data
monthly_labels  = [m['month'] for m in monthly_stats]
monthly_util    = [m['util'] for m in monthly_stats]
monthly_forecast= [m['forecast'] for m in monthly_stats]
monthly_actual  = [m['actual'] for m in monthly_stats]

acct_labels = [a['account'][:25] for a in acct_rollup[:10]]
acct_utils  = [a['avg_util'] for a in acct_rollup[:10]]
acct_counts = [a['count'] for a in acct_rollup[:10]]

grade_labels = [g['grade'] for g in grade_rollup]
grade_counts = [g['count'] for g in grade_rollup]
grade_utils  = [g['avg_util'] for g in grade_rollup]

country_labels = [c['country'] for c in country_rollup]
country_counts = [c['count'] for c in country_rollup]

def util_badge(u):
    if u >= 90:
        return f'<span class="badge badge-green">{u}%</span>'
    elif u >= 80:
        return f'<span class="badge badge-yellow">{u}%</span>'
    else:
        return f'<span class="badge badge-red">{u}%</span>'

def util_bar(u):
    color = "#22c55e" if u >= 90 else "#f59e0b" if u >= 80 else "#ef4444"
    return f'''<div class="util-bar-bg"><div class="util-bar-fill" style="width:{min(u,100)}%;background:{color};">{u}%</div></div>'''

# Build account table rows
acct_rows = ""
for i, a in enumerate(acct_rollup):
    month_cells = ""
    for m in months_order:
        u = a['months'][m]
        color = "#dcfce7" if u >= 90 else "#fef9c3" if u >= 80 else "#fee2e2"
        month_cells += f'<td style="background:{color};text-align:center;font-weight:600;">{u}%</td>'
    alert_cls = "badge-green" if a['avg_util'] >= 90 else "badge-yellow" if a['avg_util'] >= 80 else "badge-red"
    acct_rows += f'''
    <tr>
      <td>{i+1}</td>
      <td style="font-weight:600;">{a['account']}</td>
      <td style="text-align:center;">{a['count']}</td>
      <td style="text-align:center;">{util_badge(a['avg_util'])}</td>
      {month_cells}
    </tr>'''

# Low util employee rows
low_rows = ""
for i, r in enumerate(low_util_employees):
    avg_u = avg_util(r)
    month_cells = ""
    for m in months_order:
        u = r['months'].get(m, {}).get('util', 0)
        color = "#dcfce7" if u >= 90 else "#fef9c3" if u >= 80 else "#fee2e2"
        month_cells += f'<td style="background:{color};text-align:center;font-weight:600;">{u}%</td>'
    low_rows += f'''
    <tr>
      <td>{i+1}</td>
      <td style="font-weight:600;">{r['Associate Name']}</td>
      <td style="text-align:center;">{r['Grade']}</td>
      <td>{r['Account'][:30]}</td>
      <td style="text-align:center;">{r['Country']}</td>
      <td style="text-align:center;">{util_badge(avg_u)}</td>
      {month_cells}
    </tr>'''

# Monthly summary rows
monthly_rows = ""
for ms in monthly_stats:
    color = "#dcfce7" if ms['util'] >= 90 else "#fef9c3" if ms['util'] >= 80 else "#fee2e2"
    monthly_rows += f'''
    <tr>
      <td style="font-weight:700;">{ms['month']}</td>
      <td style="text-align:center;">{ms['wd']}</td>
      <td style="text-align:center;">{ms['forecast']}</td>
      <td style="text-align:center;">{ms['actual']}</td>
      <td style="text-align:center;">{ms['leave']}</td>
      <td style="background:{color};text-align:center;font-weight:700;">{ms['util']}%</td>
    </tr>'''

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>QEA UHG Leave &amp; Utilization Dashboard 2026</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', sans-serif; background: #f0f4f8; color: #1e293b; }}

  /* Header */
  .header {{
    background: linear-gradient(135deg, #1e3a5f 0%, #2e75b6 100%);
    color: white; padding: 28px 36px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
  }}
  .header h1 {{ font-size: 1.6rem; font-weight: 700; letter-spacing: 0.3px; }}
  .header p  {{ font-size: 0.85rem; opacity: 0.85; margin-top: 4px; }}
  .header-meta {{ text-align: right; font-size: 0.8rem; opacity: 0.8; }}

  /* Layout */
  .container {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}
  .section-title {{
    font-size: 1.05rem; font-weight: 700; color: #1e3a5f;
    margin: 28px 0 14px; padding-left: 12px;
    border-left: 4px solid #2e75b6;
  }}

  /* KPI Cards */
  .kpi-grid {{ display: grid; grid-template-columns: repeat(7, 1fr); gap: 14px; margin-bottom: 8px; }}
  .kpi-card {{
    background: white; border-radius: 12px; padding: 18px 14px;
    text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    border-top: 4px solid #2e75b6; transition: transform 0.2s;
  }}
  .kpi-card:hover {{ transform: translateY(-3px); }}
  .kpi-card .val {{ font-size: 1.9rem; font-weight: 800; color: #1e3a5f; }}
  .kpi-card .lbl {{ font-size: 0.72rem; color: #64748b; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.5px; }}
  .kpi-card.red   {{ border-top-color: #ef4444; }} .kpi-card.red   .val {{ color: #ef4444; }}
  .kpi-card.yellow{{ border-top-color: #f59e0b; }} .kpi-card.yellow .val {{ color: #d97706; }}
  .kpi-card.green {{ border-top-color: #22c55e; }} .kpi-card.green .val {{ color: #16a34a; }}

  /* Charts grid */
  .charts-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 8px; }}
  .charts-grid-3 {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; margin-bottom: 8px; }}
  .chart-card {{
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
  }}
  .chart-card h3 {{ font-size: 0.9rem; font-weight: 700; color: #1e3a5f; margin-bottom: 14px; }}
  .chart-wrap {{ position: relative; height: 240px; }}

  /* Tables */
  .table-card {{
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07); margin-bottom: 20px; overflow-x: auto;
  }}
  .table-card h3 {{ font-size: 0.95rem; font-weight: 700; color: #1e3a5f; margin-bottom: 14px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.82rem; }}
  th {{
    background: #1e3a5f; color: white; padding: 10px 10px;
    text-align: left; font-size: 0.78rem; white-space: nowrap;
  }}
  td {{ padding: 8px 10px; border-bottom: 1px solid #e2e8f0; white-space: nowrap; }}
  tr:hover td {{ background: #f8fafc; }}
  tr:last-child td {{ border-bottom: none; }}

  /* Badges */
  .badge {{ display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; }}
  .badge-green  {{ background: #dcfce7; color: #16a34a; }}
  .badge-yellow {{ background: #fef9c3; color: #b45309; }}
  .badge-red    {{ background: #fee2e2; color: #dc2626; }}

  /* Util bar */
  .util-bar-bg {{ background: #e2e8f0; border-radius: 20px; height: 18px; width: 100%; min-width: 80px; }}
  .util-bar-fill {{ border-radius: 20px; height: 18px; font-size: 0.7rem; color: white; display: flex; align-items: center; justify-content: center; font-weight: 700; min-width: 32px; transition: width 0.5s; }}

  /* Legend */
  .legend {{ display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }}
  .legend-item {{ display: flex; align-items: center; gap: 6px; font-size: 0.78rem; }}
  .legend-dot {{ width: 12px; height: 12px; border-radius: 50%; }}

  /* Footer */
  .footer {{ text-align: center; padding: 20px; color: #94a3b8; font-size: 0.78rem; }}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>QEA &ndash; UHG Leave &amp; Utilization Dashboard 2026</h1>
    <p>H1 2026 &nbsp;|&nbsp; Jan &ndash; Jun &nbsp;|&nbsp; QEA HC &amp; QEA NFT Pools</p>
  </div>
  <div class="header-meta">
    <div>Generated: March 2026</div>
    <div style="margin-top:4px;">Source: QEA-UHG-Leave-Forecast-2026.xlsx</div>
    <div style="margin-top:4px;">Utilization = (Working Days &minus; Leave) &divide; Working Days &times; 100</div>
  </div>
</div>

<div class="container">

  <!-- KPI CARDS -->
  <div class="section-title">Key Performance Indicators</div>
  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="val">{total_assoc}</div>
      <div class="lbl">Total Associates</div>
    </div>
    <div class="kpi-card">
      <div class="val">{int(total_forecast)}</div>
      <div class="lbl">H1 Forecast Leave (Days)</div>
    </div>
    <div class="kpi-card">
      <div class="val">{int(total_actual)}</div>
      <div class="lbl">H1 Actual Leave (Days)</div>
    </div>
    <div class="kpi-card" style="border-top-color:#2e75b6;">
      <div class="val" style="color:#2e75b6;">{overall_avg}%</div>
      <div class="lbl">Avg H1 Utilization</div>
    </div>
    <div class="kpi-card green">
      <div class="val">{high_count}</div>
      <div class="lbl">High Util (&ge;90%)</div>
    </div>
    <div class="kpi-card yellow">
      <div class="val">{med_count}</div>
      <div class="lbl">Medium Util (80&ndash;89%)</div>
    </div>
    <div class="kpi-card red">
      <div class="val">{low_count}</div>
      <div class="lbl">Low Util (&lt;80%)</div>
    </div>
  </div>

  <!-- CHARTS ROW 1 -->
  <div class="section-title">Monthly Trends</div>
  <div class="charts-grid">
    <div class="chart-card">
      <h3>Monthly Avg Utilization % &mdash; H1 2026</h3>
      <div class="chart-wrap"><canvas id="chartUtil"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>Leave Forecast vs Actuals by Month</h3>
      <div class="chart-wrap"><canvas id="chartLeave"></canvas></div>
    </div>
  </div>

  <!-- CHARTS ROW 2 -->
  <div class="section-title">Workforce Breakdown</div>
  <div class="charts-grid-3">
    <div class="chart-card">
      <h3>Top 10 Accounts &mdash; Avg Utilization %</h3>
      <div class="chart-wrap"><canvas id="chartAcct"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>Associates by Grade</h3>
      <div class="chart-wrap"><canvas id="chartGrade"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>Associates by Country</h3>
      <div class="chart-wrap"><canvas id="chartCountry"></canvas></div>
    </div>
  </div>

  <!-- MONTHLY SUMMARY TABLE -->
  <div class="section-title">Monthly Leave &amp; Utilization Summary</div>
  <div class="table-card">
    <table>
      <thead>
        <tr>
          <th>Month</th>
          <th>Working Days</th>
          <th>Total Forecast (Days)</th>
          <th>Total Actuals (Days)</th>
          <th>Total Leave Used (Days)</th>
          <th>Avg Utilization %</th>
        </tr>
      </thead>
      <tbody>{monthly_rows}</tbody>
    </table>
  </div>

  <!-- ACCOUNT ROLLUP TABLE -->
  <div class="section-title">Account-Level Utilization Rollup</div>
  <div class="table-card">
    <table>
      <thead>
        <tr>
          <th>#</th><th>Account</th><th>Associates</th><th>Avg H1 Util%</th>
          {"".join(f"<th>{m}</th>" for m in months_order)}
        </tr>
      </thead>
      <tbody>{acct_rows}</tbody>
    </table>
  </div>

  <!-- LOW/MED UTIL ALERT TABLE -->
  <div class="section-title">Utilization Attention List (Top 20 Lowest)</div>
  <div class="table-card">
    <div class="legend">
      <div class="legend-item"><div class="legend-dot" style="background:#22c55e;"></div> High (&ge;90%)</div>
      <div class="legend-item"><div class="legend-dot" style="background:#f59e0b;"></div> Medium (80&ndash;89%)</div>
      <div class="legend-item"><div class="legend-dot" style="background:#ef4444;"></div> Low (&lt;80%)</div>
    </div>
    <table>
      <thead>
        <tr>
          <th>#</th><th>Associate Name</th><th>Grade</th><th>Account</th><th>Country</th><th>Avg Util%</th>
          {"".join(f"<th>{m}</th>" for m in months_order)}
        </tr>
      </thead>
      <tbody>{low_rows}</tbody>
    </table>
  </div>

</div>

<div class="footer">QEA &ndash; UHG Leave &amp; Utilization Report 2026 &nbsp;&bull;&nbsp; Generated by Claude Code</div>

<script>
const MONTHS = {monthly_labels};
const UTIL   = {monthly_util};
const FORECAST = {monthly_forecast};
const ACTUAL   = {monthly_actual};
const ACCT_LBL = {acct_labels};
const ACCT_UTL = {acct_utils};
const ACCT_CNT = {acct_counts};
const GRADE_LBL= {grade_labels};
const GRADE_CNT= {grade_counts};
const GRADE_UTL= {grade_utils};
const CTR_LBL  = {country_labels};
const CTR_CNT  = {country_counts};

const defOpts = (title) => ({{
  responsive: true, maintainAspectRatio: false,
  plugins: {{ legend: {{ display: false }}, tooltip: {{ mode: 'index', intersect: false }} }},
  scales: {{ x: {{ grid: {{ display: false }} }}, y: {{ grid: {{ color: '#f1f5f9' }} }} }}
}});

// Chart 1 – Monthly Util Line
new Chart(document.getElementById('chartUtil'), {{
  type: 'line',
  data: {{
    labels: MONTHS,
    datasets: [{{
      label: 'Avg Utilization %',
      data: UTIL,
      borderColor: '#2e75b6', backgroundColor: 'rgba(46,117,182,0.12)',
      fill: true, tension: 0.4, pointBackgroundColor: '#2e75b6',
      pointRadius: 5, borderWidth: 2.5
    }}]
  }},
  options: {{
    ...defOpts(),
    plugins: {{ legend: {{ display: true }}, tooltip: {{ callbacks: {{ label: ctx => ctx.parsed.y + '%' }} }} }},
    scales: {{
      x: {{ grid: {{ display: false }} }},
      y: {{ min: 85, max: 100, grid: {{ color: '#f1f5f9' }},
            ticks: {{ callback: v => v + '%' }} }}
    }}
  }}
}});

// Chart 2 – Forecast vs Actuals bar
new Chart(document.getElementById('chartLeave'), {{
  type: 'bar',
  data: {{
    labels: MONTHS,
    datasets: [
      {{ label: 'Forecast', data: FORECAST, backgroundColor: 'rgba(46,117,182,0.75)', borderRadius: 4 }},
      {{ label: 'Actuals',  data: ACTUAL,   backgroundColor: 'rgba(34,197,94,0.75)',  borderRadius: 4 }}
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: true, position: 'top' }} }},
    scales: {{ x: {{ grid: {{ display: false }} }}, y: {{ grid: {{ color: '#f1f5f9' }} }} }}
  }}
}});

// Chart 3 – Account util horizontal bar
new Chart(document.getElementById('chartAcct'), {{
  type: 'bar',
  data: {{
    labels: ACCT_LBL,
    datasets: [{{
      label: 'Avg Util %',
      data: ACCT_UTL,
      backgroundColor: ACCT_UTL.map(u => u >= 90 ? 'rgba(34,197,94,0.8)' : u >= 80 ? 'rgba(245,158,11,0.8)' : 'rgba(239,68,68,0.8)'),
      borderRadius: 4
    }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => ctx.parsed.x + '%' }} }} }},
    scales: {{
      x: {{ min: 80, max: 100, ticks: {{ callback: v => v + '%' }}, grid: {{ color: '#f1f5f9' }} }},
      y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }} }} }}
    }}
  }}
}});

// Chart 4 – Grade doughnut
new Chart(document.getElementById('chartGrade'), {{
  type: 'doughnut',
  data: {{
    labels: GRADE_LBL,
    datasets: [{{
      data: GRADE_CNT,
      backgroundColor: ['#2e75b6','#70ad47','#ffc000','#ed7d31','#4472c4','#a5a5a5','#5b9bd5','#c00000'],
      borderWidth: 2, borderColor: '#fff'
    }}]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ position: 'right', labels: {{ font: {{ size: 10 }}, padding: 8 }} }},
      tooltip: {{ callbacks: {{ label: ctx => ctx.label + ': ' + ctx.parsed + ' associates' }} }}
    }}
  }}
}});

// Chart 5 – Country doughnut
new Chart(document.getElementById('chartCountry'), {{
  type: 'doughnut',
  data: {{
    labels: CTR_LBL,
    datasets: [{{
      data: CTR_CNT,
      backgroundColor: ['#2e75b6','#ed7d31','#70ad47','#ffc000'],
      borderWidth: 2, borderColor: '#fff'
    }}]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ position: 'right', labels: {{ font: {{ size: 10 }}, padding: 8 }} }},
      tooltip: {{ callbacks: {{ label: ctx => ctx.label + ': ' + ctx.parsed + ' associates' }} }}
    }}
  }}
}});
</script>
</body>
</html>"""

out_path = r'C:\Users\202294\Downloads\QEA-UHG-Utilization-Dashboard-2026.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"Saved: {out_path}")

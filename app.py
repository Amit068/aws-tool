import streamlit as st
import pdfplumber
import re
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# Page Config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="AWS Invoice Dashboard",
    page_icon="https://upload.wikimedia.org/wikipedia/commons/9/93/Amazon_Web_Services_Logo.svg",
    layout="wide",
)

# ─────────────────────────────────────────────
# Custom CSS — Dark AWS Theme
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Global background ── */
    .stApp { background-color: #0A1628 !important; }
    .main  { background-color: #0A1628 !important; }
    section[data-testid="stSidebar"] { background-color: #0D1F3C !important; }
    .block-container { padding-top: 0 !important; max-width: 1400px; }

    /* ── Hero header bar ── */
    .aws-hero {
        background: linear-gradient(135deg, #0D1F3C 0%, #0F2A50 60%, #132E54 100%);
        border-bottom: 3px solid #FF9900;
        padding: 28px 40px 24px 40px;
        margin: -1rem -1rem 2rem -1rem;
        display: flex;
        align-items: center;
        gap: 28px;
    }
    .aws-logo-wrap {
        background: #ffffff;
        border-radius: 14px;
        padding: 10px 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        min-width: 100px;
        box-shadow: 0 4px 20px rgba(255,153,0,0.25);
    }
    .aws-logo-wrap img { height: 44px; width: auto; display: block; }
    .hero-text h1 {
        color: #FFFFFF;
        font-size: 32px;
        font-weight: 700;
        margin: 0 0 4px 0;
        letter-spacing: -0.5px;
    }
    .hero-text p {
        color: #8BAFD4;
        font-size: 14px;
        margin: 0;
    }
    .hero-badge {
        margin-left: auto;
        background: rgba(255,153,0,0.15);
        border: 1px solid rgba(255,153,0,0.4);
        color: #FF9900;
        padding: 6px 16px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
        letter-spacing: 0.3px;
        white-space: nowrap;
    }

    /* ── Cards ── */
    .metric-card {
        background: #0D1F3C;
        border-radius: 14px;
        padding: 20px 22px;
        border: 1px solid #1A3A6B;
        border-top: 3px solid #FF9900;
        transition: transform 0.15s;
    }
    .metric-card:hover { transform: translateY(-2px); }
    .metric-label { font-size: 12px; color: #8BAFD4; font-weight: 600;
                    letter-spacing: 0.8px; text-transform: uppercase; }
    .metric-value { font-size: 26px; color: #FFFFFF; font-weight: 700; margin-top: 6px; }
    .metric-sub   { font-size: 11px; color: #5A7FA8; margin-top: 3px; }

    /* ── Section titles ── */
    .section-title {
        font-size: 16px; font-weight: 700; color: #FF9900;
        border-bottom: 1px solid #1A3A6B;
        padding-bottom: 8px; margin: 28px 0 16px 0;
        letter-spacing: 0.3px;
    }

    /* ── Insight box ── */
    .insight-box {
        background: #0D1F3C;
        border-radius: 10px;
        padding: 14px 18px;
        border-left: 4px solid #FF9900;
        margin: 8px 0;
        color: #C5D8EE;
    }
    .insight-box b { color: #FF9900; }

    /* ── File uploader ── */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #1A3A6B !important;
        border-radius: 14px !important;
        padding: 12px !important;
        background: #0D1F3C !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #FF9900 !important;
    }

    /* ── Dataframe / table ── */
    div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

    /* ── Tab styling ── */
    button[data-baseweb="tab"] {
        color: #8BAFD4 !important;
        font-weight: 600 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #FF9900 !important;
        border-bottom-color: #FF9900 !important;
    }

    /* ── Divider ── */
    hr { border-color: #1A3A6B !important; }

    /* ── General text ── */
    p, span, label, div { color: #C5D8EE; }
    h1, h2, h3 { color: #FFFFFF !important; }

    /* ── Progress bar ── */
    div[data-testid="stProgressBar"] > div > div {
        background: linear-gradient(90deg, #FF9900, #FFB84D) !important;
    }

    /* ── Download button ── */
    div[data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #FF9900, #E68A00) !important;
        color: #0A1628 !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 10px !important;
        font-size: 16px !important;
        padding: 14px !important;
        letter-spacing: 0.3px;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background: linear-gradient(135deg, #FFB84D, #FF9900) !important;
        transform: translateY(-1px);
    }

    /* ── Info box ── */
    div[data-testid="stAlert"] {
        background: #0D1F3C !important;
        border-color: #1A3A6B !important;
        color: #8BAFD4 !important;
        border-radius: 10px !important;
    }

    /* ── Scrollbar ── */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #0A1628; }
    ::-webkit-scrollbar-thumb { background: #1A3A6B; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #FF9900; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# PDF Parsing Logic (same as original)
# ─────────────────────────────────────────────
_SAVINGS_SUB = re.compile(r'^Savings Plan \(Charges covered by Savings Plans\)')
_SUB_STARTS  = ('Charges ', 'Tax ', 'GST ', 'Credits ', 'Discount (')

def _usd(text):
    m = re.search(r'-?USD\s*([\d,]+\.?\d*)', text)
    if not m:
        return 0.0
    val = float(m.group(1).replace(',', ''))
    return -val if re.search(r'(?<!\w)-USD', text) or text.strip().startswith('-') else val

def _is_sub(line):
    if _SAVINGS_SUB.match(line):
        return True
    return any(line.startswith(p) for p in _SUB_STARTS)

@st.cache_data(show_spinner=False)
def parse_pdf(pdf_bytes):
    account_rows = []
    service_data = {}

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    blocks = re.split(r'Summary for Linked Account\s*\n', full_text)

    for block in blocks[1:]:
        lines = [l.strip() for l in block.split('\n') if l.strip()]
        if not lines:
            continue

        hm = re.match(r'^(.+?)\s*\((\d{12})\)\s+USD\s+([\d,]+\.?\d*)', lines[0])
        if not hm:
            continue

        acc_name = hm.group(1).strip()
        acc_id   = hm.group(2).strip()

        charges = savings_plan = private_rate = bundled_dis = 0.0
        edp_dis = credits = gst = total_alloc = 0.0

        in_summary = True
        in_detail  = False
        svc_name   = None
        svc_charges = svc_savings = 0.0
        cur_svcs   = {}

        for line in lines[1:]:
            if 'total allocated for this statement' in line.lower():
                m = re.search(r'USD\s*([\d,]+\.?\d*)', line)
                total_alloc = float(m.group(1).replace(',', '')) if m else 0.0
                in_summary = False
                continue
            if 'Detail for Linked Account' in line:
                in_detail = True; in_summary = False
                continue
            if 'For line item details' in line or 'Account Activity Page' in line:
                break

            if in_summary:
                if re.match(r'^Charges\s+USD', line):
                    charges = abs(_usd(line))
                elif _SAVINGS_SUB.match(line):
                    savings_plan = abs(_usd(line))
                elif 'Private Rate Card' in line:
                    private_rate = abs(_usd(line))
                elif 'Bundled Discount' in line:
                    bundled_dis = abs(_usd(line))
                elif ('Distribution Program Discount' in line or
                      'Enterprise Discount Program' in line or
                      'AWS Distribution Program Discount' in line):
                    edp_dis = abs(_usd(line))
                elif re.match(r'^Credits\s+', line):
                    credits = _usd(line)
                elif re.match(r'^Tax\s+USD', line) or re.match(r'^GST\s+USD', line):
                    gst += abs(_usd(line))
            elif in_detail:
                is_sub    = _is_sub(line)
                svc_match = re.match(r'^(.+?)\s+USD\s+([\d,]+\.?\d*)$', line)
                if svc_match and not is_sub:
                    if svc_name:
                        net = svc_charges - svc_savings
                        cur_svcs[svc_name] = cur_svcs.get(svc_name, 0.0) + net
                    svc_name    = svc_match.group(1).strip()
                    svc_charges = svc_savings = 0.0
                elif svc_name:
                    if re.match(r'^Charges\s+USD', line):
                        svc_charges = abs(_usd(line))
                    elif _SAVINGS_SUB.match(line):
                        svc_savings = abs(_usd(line))

        if svc_name and in_detail:
            cur_svcs[svc_name] = cur_svcs.get(svc_name, 0.0) + (svc_charges - svc_savings)

        account_rows.append({
            'Account ID'                 : acc_id,
            'Account Name'               : acc_name,
            'Charges (USD)'              : round(charges, 2),
            'Savings Plan (USD)'         : round(-savings_plan, 2),
            'Private Rate Card (USD)'    : round(-private_rate, 2),
            'Bundled Discount (USD)'     : round(-bundled_dis, 2),
            'EDP / Dist. Discount (USD)' : round(-edp_dis, 2),
            'Credits (USD)'              : round(credits, 2),
            'GST (USD)'                  : round(gst, 2),
            'Total (USD)'                : round(total_alloc, 2),
        })

        if acc_id not in service_data:
            service_data[acc_id] = {'_name': acc_name}
        for svc, val in cur_svcs.items():
            service_data[acc_id][svc] = round(service_data[acc_id].get(svc, 0.0) + val, 2)

    return account_rows, service_data


# ─────────────────────────────────────────────
# Excel Export
# ─────────────────────────────────────────────
HDR_BG = "1F4E79"; HDR_FG = "FFFFFF"
ALT_BG = "D6E4F0"; NEG_FG = "C00000"
TOTAL_BG = "BDD7EE"; TOTAL_FG = "1F4E79"

def _border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr(cell, value):
    cell.value = value
    cell.font  = Font(bold=True, color=HDR_FG, name="Arial", size=10)
    cell.fill  = PatternFill("solid", start_color=HDR_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = _border()

def _dat(cell, value, row_idx, align="left", neg=False):
    cell.value = value
    bg = ALT_BG if row_idx % 2 == 0 else "FFFFFF"
    cell.fill  = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    color = NEG_FG if neg and isinstance(value, (int, float)) and value < 0 else "000000"
    cell.font = Font(name="Arial", size=10, color=color)
    cell.border = _border()

def _total_cell(cell, value, align="right"):
    cell.value = value
    cell.font  = Font(bold=True, color=TOTAL_FG, name="Arial", size=10)
    cell.fill  = PatternFill("solid", start_color=TOTAL_BG)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = _border()

def _auto_width(ws, min_w=10, max_w=35):
    for col in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(min_w, min(length + 3, max_w))

S1_HDRS = ['Account ID','Account Name','Charges (USD)','Savings Plan (USD)',
           'Private Rate Card (USD)','Bundled Discount (USD)',
           'EDP / Dist. Discount (USD)','Credits (USD)','GST (USD)','Total (USD)']
S1_NEG = {4,5,6,7,8}
S1_NUM = set(range(3, len(S1_HDRS)+1))

def build_excel(all_rows, all_svcs):
    wb  = Workbook()
    ws1 = wb.active
    ws1.title = "Account Details"
    ws1.row_dimensions[1].height = 32
    for ci, h in enumerate(S1_HDRS, 1):
        _hdr(ws1.cell(1, ci), h)
    for ri, row in enumerate(all_rows, 2):
        for ci, h in enumerate(S1_HDRS, 1):
            _dat(ws1.cell(ri, ci), row[h], ri,
                 align="right" if ci in S1_NUM else "left", neg=(ci in S1_NEG))
    tr = len(all_rows)+2
    _total_cell(ws1.cell(tr,1), "TOTAL", align="center")
    _total_cell(ws1.cell(tr,2), "")
    for ci, h in enumerate(S1_HDRS[2:], 3):
        col_vals = [r[h] for r in all_rows if isinstance(r[h], (int,float))]
        _total_cell(ws1.cell(tr,ci), round(sum(col_vals),2))
    _auto_width(ws1)
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions

    ws2 = wb.create_sheet("Service Breakdown")
    ws2.row_dimensions[1].height = 48
    all_services = sorted({s for d in all_svcs.values() for s in d if s != '_name'})
    s2_hdrs = ['Account ID','Account Name'] + all_services
    for ci, h in enumerate(s2_hdrs, 1):
        _hdr(ws2.cell(1, ci), h)
    for ri, (acc_id, acc_data) in enumerate(all_svcs.items(), 2):
        _dat(ws2.cell(ri,1), acc_id, ri)
        _dat(ws2.cell(ri,2), acc_data.get('_name',''), ri)
        for ci, svc in enumerate(all_services, 3):
            val = acc_data.get(svc, None)
            _dat(ws2.cell(ri,ci), round(val,2) if val is not None else None, ri, align="right")
    tr2 = len(all_svcs)+2
    _total_cell(ws2.cell(tr2,1), "TOTAL", align="center")
    _total_cell(ws2.cell(tr2,2), "")
    for ci, svc in enumerate(all_services, 3):
        col_vals = [d.get(svc,0.0) for d in all_svcs.values() if isinstance(d.get(svc,0.0),(int,float))]
        _total_cell(ws2.cell(tr2,ci), round(sum(col_vals),2))
    _auto_width(ws2, min_w=12, max_w=30)
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = ws2.dimensions

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# HERO HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="aws-hero">
    <div class="aws-logo-wrap">
        <img src="https://upload.wikimedia.org/wikipedia/commons/9/93/Amazon_Web_Services_Logo.svg"
             alt="AWS Logo" style="height:44px;width:auto;display:block;"/>
    </div>
    <div class="hero-text">
        <h1 style="color:#FFFFFF;font-size:32px;font-weight:700;margin:0 0 4px 0;">Invoice Dashboard</h1>
        <p style="color:#8BAFD4;font-size:14px;margin:0;">
            Upload PDF &rarr; Extract Data &rarr; View Charts &amp; Insights
        </p>
    </div>
    <div class="hero-badge">&#9679;&nbsp;AWS Cost Analytics</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# REFRESH WARNING BANNER
# ─────────────────────────────────────────────
st.markdown("""
<div style="
    background: rgba(255,153,0,0.08);
    border: 1px solid rgba(255,153,0,0.35);
    border-left: 4px solid #FF9900;
    border-radius: 10px;
    padding: 12px 18px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 12px;
">
    <span style="font-size:20px;">⚠️</span>
    <div>
        <span style="color:#FF9900;font-weight:700;font-size:14px;">Important — Session Data Warning</span><br>
        <span style="color:#C5D8EE;font-size:13px;">
            This dashboard does <b style="color:#FF9900;">not</b> store any data permanently.
            If you <b style="color:#FF9900;">refresh</b> or <b style="color:#FF9900;">close this tab</b>,
            all uploaded data will be lost and you will need to re-upload your PDF files.
            Download the Excel report before leaving.
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📂 Upload AWS Invoice PDFs</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Select one or more AWS Invoice PDF files (max 50)",
    type=["pdf"],
    accept_multiple_files=True,
    help="Upload AWS Consolidated Invoice PDF files here"
)

if not uploaded_files:
    st.info("⬆️ Upload your PDF files above — the dashboard will update automatically.")
    st.stop()

# ─────────────────────────────────────────────
# PARSE ALL PDFs  (cached per file)
# ─────────────────────────────────────────────
all_rows = []
all_svcs = {}
errors   = []

progress_bar = st.progress(0, text="Reading PDF files...")

for i, f in enumerate(uploaded_files):
    try:
        rows, svcs = parse_pdf(f.read())
        all_rows.extend(rows)
        for acc_id, acc_data in svcs.items():
            if acc_id not in all_svcs:
                all_svcs[acc_id] = acc_data.copy()
            else:
                for k, v in acc_data.items():
                    if k == '_name':
                        all_svcs[acc_id]['_name'] = v
                    else:
                        all_svcs[acc_id][k] = round(all_svcs[acc_id].get(k,0.0) + v, 2)
    except Exception as e:
        errors.append(f"❌ {f.name}: {e}")
    progress_bar.progress(int((i+1)/len(uploaded_files)*100),
                          text=f"Processing {i+1}/{len(uploaded_files)}: {f.name}")

progress_bar.empty()

if errors:
    for err in errors:
        st.warning(err)

if not all_rows:
    st.error("No data extracted. Please check that the correct AWS Invoice PDF has been uploaded.")
    st.stop()

# ─────────────────────────────────────────────
# BUILD DATAFRAMES
# ─────────────────────────────────────────────
df_accounts = pd.DataFrame(all_rows)

# Service breakdown dataframe
all_services = sorted({s for d in all_svcs.values() for s in d if s != '_name'})
svc_rows = []
for acc_id, acc_data in all_svcs.items():
    row = {'Account ID': acc_id, 'Account Name': acc_data.get('_name', '')}
    for svc in all_services:
        row[svc] = acc_data.get(svc, 0.0) or 0.0
    svc_rows.append(row)
df_services = pd.DataFrame(svc_rows)

# Aggregated service totals (for charts)
svc_totals = {}
for acc_data in all_svcs.values():
    for k, v in acc_data.items():
        if k != '_name' and isinstance(v, (int, float)):
            svc_totals[k] = round(svc_totals.get(k, 0.0) + v, 2)

df_svc_totals = pd.DataFrame(
    sorted(svc_totals.items(), key=lambda x: x[1], reverse=True),
    columns=['Service', 'Total Cost (USD)']
).query("`Total Cost (USD)` > 0")

# ─────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📊 Summary</div>', unsafe_allow_html=True)

total_charges = df_accounts['Charges (USD)'].sum()
total_bill    = df_accounts['Total (USD)'].sum()
total_gst     = df_accounts['GST (USD)'].sum()
total_discount = (
    df_accounts.get('Savings Plan (USD)', pd.Series([0])).sum() +
    df_accounts.get('Bundled Discount (USD)', pd.Series([0])).sum() +
    df_accounts.get('EDP / Dist. Discount (USD)', pd.Series([0])).sum()
)
top_service   = df_svc_totals.iloc[0]['Service'] if not df_svc_totals.empty else "N/A"
top_svc_cost  = df_svc_totals.iloc[0]['Total Cost (USD)'] if not df_svc_totals.empty else 0

c1, c2, c3, c4, c5 = st.columns(5)

def kpi(col, icon, label, value, sub=""):
    col.markdown(f"""
    <div class="metric-card">
      <div class="metric-label">{icon} {label}</div>
      <div class="metric-value">{value}</div>
      <div class="metric-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

kpi(c1, "🏢", "Total Accounts",  str(len(all_rows)), f"{len(uploaded_files)} PDF(s)")
kpi(c2, "💰", "Total Charges",   f"${total_charges:,.2f}", "Before discounts")
kpi(c3, "🧾", "Total Bill",      f"${total_bill:,.2f}",    "After all discounts")
kpi(c4, "🏷️", "Total Discount",  f"${abs(total_discount):,.2f}", "Savings + EDP")
kpi(c5, "🔝", "Top Service",     top_service[:20]+"…" if len(top_service)>20 else top_service,
              f"${top_svc_cost:,.2f}")

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CHARTS ROW 1 — Pie + Bar
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📈 Service Cost Analysis</div>', unsafe_allow_html=True)

chart_col1, chart_col2 = st.columns([1, 1])

with chart_col1:
    st.markdown("#### 🥧 Service-wise Cost Share")
    # Top 10 services, rest as "Others"
    top_n = 10
    if len(df_svc_totals) > top_n:
        top_df    = df_svc_totals.head(top_n).copy()
        other_sum = df_svc_totals.iloc[top_n:]['Total Cost (USD)'].sum()
        others    = pd.DataFrame([{'Service': 'Others', 'Total Cost (USD)': round(other_sum,2)}])
        pie_df    = pd.concat([top_df, others], ignore_index=True)
    else:
        pie_df = df_svc_totals.copy()

    AWS_COLORS = ['#FF9900','#1A7FBF','#2DB57D','#E6522C','#8B5CF6',
                  '#3B9ED4','#F59E0B','#10B981','#EF4444','#6366F1','#EC4899']
    DARK_LAYOUT = dict(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#C5D8EE', family='Arial'),
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#C5D8EE')),
    )

    fig_pie = px.pie(
        pie_df,
        names='Service',
        values='Total Cost (USD)',
        color_discrete_sequence=AWS_COLORS,
        hole=0.45,
    )
    fig_pie.update_traces(
        textposition='inside',
        textinfo='percent+label',
        textfont=dict(color='white', size=11),
        marker=dict(line=dict(color='#0A1628', width=2)),
        hovertemplate='<b>%{label}</b><br>Cost: $%{value:,.2f}<br>Share: %{percent}<extra></extra>'
    )
    fig_pie.update_layout(
        **DARK_LAYOUT,
        margin=dict(t=10, b=10, l=10, r=10),
        legend=dict(orientation="v", x=1.02, y=0.5, font=dict(color='#C5D8EE')),
        height=380,
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with chart_col2:
    st.markdown("#### 📊 Top Services by Cost (USD)")
    bar_df = df_svc_totals.head(12)
    fig_bar = px.bar(
        bar_df,
        x='Total Cost (USD)',
        y='Service',
        orientation='h',
        color='Total Cost (USD)',
        color_continuous_scale=[[0,'#1A3A6B'],[0.5,'#1A7FBF'],[1,'#FF9900']],
        text='Total Cost (USD)',
    )
    fig_bar.update_traces(
        texttemplate='$%{text:,.0f}',
        textposition='outside',
        textfont=dict(color='#C5D8EE'),
        marker=dict(line=dict(color='#0A1628', width=0.5)),
        hovertemplate='<b>%{y}</b><br>Cost: $%{x:,.2f}<extra></extra>'
    )
    fig_bar.update_layout(
        **DARK_LAYOUT,
        yaxis=dict(autorange="reversed", color='#8BAFD4', gridcolor='#1A3A6B'),
        xaxis=dict(color='#8BAFD4', gridcolor='#1A3A6B'),
        coloraxis_showscale=False,
        margin=dict(t=10, b=10, l=10, r=100),
        height=380,
        xaxis_title="Cost (USD)",
        yaxis_title=""
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# ─────────────────────────────────────────────
# CHARTS ROW 2 — Account Comparison + Donut
# ─────────────────────────────────────────────
if len(all_rows) > 1:
    st.markdown('<div class="section-title">🏢 Account-wise Comparison</div>', unsafe_allow_html=True)
    acct_col1, acct_col2 = st.columns([3, 2])

    with acct_col1:
        st.markdown("#### 💳 Account-wise Total Bill")
        fig_acct = px.bar(
            df_accounts.sort_values('Total (USD)', ascending=False),
            x='Account Name',
            y='Total (USD)',
            color='Total (USD)',
            color_continuous_scale=[[0,'#1A3A6B'],[0.5,'#1A7FBF'],[1,'#FF9900']],
            text='Total (USD)',
        )
        fig_acct.update_traces(
            texttemplate='$%{text:,.0f}',
            textposition='outside',
            textfont=dict(color='#C5D8EE'),
            marker=dict(line=dict(color='#0A1628', width=0.5)),
        )
        fig_acct.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#C5D8EE'),
            coloraxis_showscale=False,
            xaxis=dict(tickangle=-35, color='#8BAFD4', gridcolor='#1A3A6B'),
            yaxis=dict(color='#8BAFD4', gridcolor='#1A3A6B'),
            margin=dict(t=10, b=80),
            height=340,
        )
        st.plotly_chart(fig_acct, use_container_width=True)

    with acct_col2:
        st.markdown("#### 🍩 Account Cost Share")
        fig_acct_pie = px.pie(
            df_accounts,
            names='Account Name',
            values='Total (USD)',
            color_discrete_sequence=['#FF9900','#1A7FBF','#2DB57D','#E6522C','#8B5CF6','#3B9ED4'],
            hole=0.5,
        )
        fig_acct_pie.update_traces(
            textposition='inside',
            textinfo='percent+label',
            textfont=dict(color='white', size=11),
            marker=dict(line=dict(color='#0A1628', width=2)),
        )
        fig_acct_pie.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#C5D8EE'),
            margin=dict(t=10,b=10,l=10,r=10),
            height=340,
            showlegend=False,
        )
        st.plotly_chart(fig_acct_pie, use_container_width=True)

# ─────────────────────────────────────────────
# MARGIN INSIGHT TABLE
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">💡 Margin Opportunity — Which services have the highest margin potential?</div>', unsafe_allow_html=True)

st.markdown("""
<div class="insight-box">
<b>How to use this table:</b> Services with the <b>highest cost share</b> have strong demand —
you can apply a higher margin there. Services used across <b>many accounts</b> are "sticky" —
clients depend on them, making them ideal for a pricing uplift.
</div>
""", unsafe_allow_html=True)

# How many accounts use each service
svc_account_count = {}
for acc_data in all_svcs.values():
    for k, v in acc_data.items():
        if k != '_name' and isinstance(v, (int,float)) and v > 0:
            svc_account_count[k] = svc_account_count.get(k, 0) + 1

margin_df = df_svc_totals.copy()
margin_df['Accounts Using'] = margin_df['Service'].map(lambda s: svc_account_count.get(s, 0))
margin_df['Avg Cost/Account (USD)'] = (margin_df['Total Cost (USD)'] / margin_df['Accounts Using']).round(2)

total_cost_sum = margin_df['Total Cost (USD)'].sum()
margin_df['Cost Share (%)'] = (margin_df['Total Cost (USD)'] / total_cost_sum * 100).round(1)

# Margin suggestion
def margin_suggestion(row):
    if row['Cost Share (%)'] >= 20:
        return "🔴 High — Apply higher margin"
    elif row['Cost Share (%)'] >= 8:
        return "🟠 Medium — Moderate margin"
    elif row['Accounts Using'] >= 3:
        return "🟡 Sticky — High usage, small uplift"
    else:
        return "🟢 Low usage"

margin_df['Margin Suggestion'] = margin_df.apply(margin_suggestion, axis=1)
margin_df = margin_df.rename(columns={'Total Cost (USD)': 'Total Cost (USD) ↓'})

st.dataframe(
    margin_df[['Service','Total Cost (USD) ↓','Cost Share (%)','Accounts Using',
               'Avg Cost/Account (USD)','Margin Suggestion']],
    use_container_width=True,
    hide_index=True,
    column_config={
        'Total Cost (USD) ↓' : st.column_config.NumberColumn(format="$%.2f"),
        'Avg Cost/Account (USD)': st.column_config.NumberColumn(format="$%.2f"),
        'Cost Share (%)'     : st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.1f%%"),
    }
)

# ─────────────────────────────────────────────
# DATA TABLES
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📋 Raw Data Tables</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🏢 Account Details", "🔧 Service Breakdown"])

with tab1:
    st.dataframe(df_accounts, use_container_width=True, hide_index=True,
                 column_config={
                     c: st.column_config.NumberColumn(format="$%.2f")
                     for c in df_accounts.columns if '(USD)' in c
                 })

with tab2:
    st.dataframe(df_services, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# EXCEL DOWNLOAD
# ─────────────────────────────────────────────
st.divider()
st.markdown("### 📥 Download Excel Report")

excel_buf = build_excel(all_rows, all_svcs)
st.download_button(
    label="⬇️  Download AWS_Invoice_Report.xlsx",
    data=excel_buf,
    file_name="AWS_Invoice_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

st.markdown(
    "<div style='text-align:center;color:#5A7FA8;font-size:12px;margin-top:16px;padding:12px;"
    "border-top:1px solid #1A3A6B;'>"
    "☁️ &nbsp;AWS Invoice Dashboard &nbsp;•&nbsp; Built with Streamlit + Plotly &nbsp;•&nbsp; "
    "<span style='color:#FF9900;'>Powered by AWS Analytics</span>"
    "</div>",
    unsafe_allow_html=True
)

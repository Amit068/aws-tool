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
    page_icon="☁️",
    layout="wide",
)

# ─────────────────────────────────────────────
# Custom CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #F0F4F8; }
    .block-container { padding-top: 1.5rem; }
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 5px solid #1F4E79;
    }
    .metric-label { font-size: 13px; color: #666; font-weight: 600; letter-spacing: 0.5px; }
    .metric-value { font-size: 28px; color: #1F4E79; font-weight: 700; margin-top: 4px; }
    .metric-sub   { font-size: 12px; color: #999; margin-top: 2px; }
    .section-title {
        font-size: 18px; font-weight: 700; color: #1F4E79;
        border-bottom: 2px solid #BDD7EE;
        padding-bottom: 6px; margin: 20px 0 14px 0;
    }
    .insight-box {
        background: #EAF3FB; border-radius: 10px;
        padding: 14px 18px; border-left: 4px solid #2E75B6;
        margin: 8px 0;
    }
    .insight-box b { color: #1F4E79; }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #2E75B6;
        border-radius: 12px;
        padding: 10px;
        background: white;
    }
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
# HEADER
# ─────────────────────────────────────────────
col_logo, col_title = st.columns([1, 8])
with col_logo:
    st.markdown("## ☁️")
with col_title:
    st.markdown("# AWS Invoice Dashboard")
    st.markdown("<span style='color:#666;font-size:14px;'>PDF upload karo → data extract hoga → dashboard pe dekho</span>", unsafe_allow_html=True)

st.divider()

# ─────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📂 AWS Invoice PDF Upload karo</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Ek ya zyada AWS Invoice PDF files select karo (max 50)",
    type=["pdf"],
    accept_multiple_files=True,
    help="AWS Consolidated Invoice PDF files yahan upload karo"
)

if not uploaded_files:
    st.info("⬆️ Upar se PDF upload karo — dashboard automatically update ho jaayega.")
    st.stop()

# ─────────────────────────────────────────────
# PARSE ALL PDFs
# ─────────────────────────────────────────────
all_rows = []
all_svcs = {}
errors   = []

progress_bar = st.progress(0, text="PDF parse ho raha hai...")

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
    st.error("Koi data extract nahi hua. Check karo ki sahi AWS Invoice PDF hai.")
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

    fig_pie = px.pie(
        pie_df,
        names='Service',
        values='Total Cost (USD)',
        color_discrete_sequence=px.colors.sequential.Blues_r,
        hole=0.4,
    )
    fig_pie.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate='<b>%{label}</b><br>Cost: $%{value:,.2f}<br>Share: %{percent}<extra></extra>'
    )
    fig_pie.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        legend=dict(orientation="v", x=1.02, y=0.5),
        height=380,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
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
        color_continuous_scale='Blues',
        text='Total Cost (USD)',
    )
    fig_bar.update_traces(
        texttemplate='$%{text:,.0f}',
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Cost: $%{x:,.2f}<extra></extra>'
    )
    fig_bar.update_layout(
        yaxis=dict(autorange="reversed"),
        coloraxis_showscale=False,
        margin=dict(t=10, b=10, l=10, r=80),
        height=380,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
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
            color_continuous_scale='Blues',
            text='Total (USD)',
        )
        fig_acct.update_traces(
            texttemplate='$%{text:,.0f}',
            textposition='outside',
        )
        fig_acct.update_layout(
            coloraxis_showscale=False,
            xaxis_tickangle=-35,
            margin=dict(t=10, b=80),
            height=340,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
        )
        st.plotly_chart(fig_acct, use_container_width=True)

    with acct_col2:
        st.markdown("#### 🍩 Account Cost Share")
        fig_acct_pie = px.pie(
            df_accounts,
            names='Account Name',
            values='Total (USD)',
            color_discrete_sequence=px.colors.sequential.Blues_r,
            hole=0.5,
        )
        fig_acct_pie.update_traces(
            textposition='inside',
            textinfo='percent+label',
        )
        fig_acct_pie.update_layout(
            margin=dict(t=10,b=10,l=10,r=10),
            height=340,
            showlegend=False,
            paper_bgcolor='rgba(0,0,0,0)',
        )
        st.plotly_chart(fig_acct_pie, use_container_width=True)

# ─────────────────────────────────────────────
# MARGIN INSIGHT TABLE
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">💡 Margin Opportunity — Kaun si service pe zyada margin rakho?</div>', unsafe_allow_html=True)

st.markdown("""
<div class="insight-box">
<b>Kaise use karein:</b> Jo service ka <b>cost sabse zyada</b> hai, 
uski demand high hai — wahan aap higher margin rakh sakte ho. 
Jo service <b>bahut zyada accounts</b> mein use ho rahi hai, 
woh "sticky" service hai — uska pricing thoda upar karo.
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
        return "🔴 High — zyada margin rakho"
    elif row['Cost Share (%)'] >= 8:
        return "🟠 Medium — moderate margin"
    elif row['Accounts Using'] >= 3:
        return "🟡 Sticky — usage high, thoda badhao"
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
st.markdown("### 📥 Excel Report Download")

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
    "<div style='text-align:center;color:#aaa;font-size:12px;margin-top:12px;'>"
    "AWS Invoice Dashboard • Built with Streamlit + Plotly"
    "</div>",
    unsafe_allow_html=True
)
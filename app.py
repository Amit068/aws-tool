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
    /* ── Global background — clean AWS blue ── */
    .stApp { background-color: #EBF5FB !important; }
    .main  { background-color: #EBF5FB !important; }
    section[data-testid="stSidebar"] { background-color: #BBDEFB !important; }
    .block-container { padding-top: 0 !important; max-width: 1400px; }

    /* ── Hero header — solid AWS blue, no clip ── */
    .aws-hero {
        background: linear-gradient(135deg, #1565C0 0%, #1976D2 60%, #1E88E5 100%);
        border-bottom: 4px solid #FF9900;
        padding: 24px 40px;
        margin: 0 0 2rem 0;
        display: flex;
        align-items: center;
        gap: 24px;
        border-radius: 0 0 16px 16px;
        box-shadow: 0 4px 20px rgba(21,101,192,0.3);
    }
    /* Logo box — enough padding so image never clips */
    .aws-logo-wrap {
        background: #ffffff;
        border-radius: 12px;
        padding: 10px 18px;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
        width: 160px;
        height: 68px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }
    .aws-logo-wrap img {
        width: 130px;
        height: auto;
        display: block;
        object-fit: contain;
    }
    .hero-text h1 {
        color: #FFFFFF;
        font-size: 30px;
        font-weight: 700;
        margin: 0 0 4px 0;
    }
    .hero-text p { color: #BDD9F2; font-size: 14px; margin: 0; }
    .hero-badge {
        margin-left: auto;
        background: rgba(255,153,0,0.2);
        border: 1px solid rgba(255,153,0,0.6);
        color: #FF9900;
        padding: 6px 18px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 700;
        white-space: nowrap;
    }

    /* ── Metric cards ── */
    .metric-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 18px 20px;
        border: 1px solid #BBDEFB;
        border-top: 4px solid #1976D2;
        transition: box-shadow 0.15s;
        box-shadow: 0 2px 8px rgba(25,118,210,0.08);
    }
    .metric-card:hover { box-shadow: 0 4px 16px rgba(25,118,210,0.2); }
    .metric-label { font-size: 11px; color: #1565C0; font-weight: 700;
                    letter-spacing: 0.8px; text-transform: uppercase; }
    .metric-value { font-size: 26px; color: #0D2B4E; font-weight: 700; margin-top: 6px; }
    .metric-sub   { font-size: 11px; color: #64B5F6; margin-top: 3px; }

    /* ── Section titles ── */
    .section-title {
        font-size: 15px; font-weight: 700; color: #1565C0;
        border-bottom: 2px solid #90CAF9;
        padding-bottom: 8px; margin: 28px 0 16px 0;
        text-transform: uppercase; letter-spacing: 0.5px;
    }

    /* ── Insight / info box ── */
    .insight-box {
        background: #E3F2FD;
        border-radius: 10px;
        padding: 14px 18px;
        border-left: 4px solid #1976D2;
        margin: 8px 0;
        color: #0D2B4E;
    }
    .insight-box b { color: #1565C0; }

    /* ── Warning banner ── */
    .warn-banner {
        background: #FFF8EC;
        border: 1px solid #FFCC70;
        border-left: 4px solid #FF9900;
        border-radius: 10px;
        padding: 12px 18px;
        margin-bottom: 20px;
    }

    /* ── File uploader ── */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #64B5F6 !important;
        border-radius: 12px !important;
        padding: 12px !important;
        background: #ffffff !important;
    }
    div[data-testid="stFileUploader"]:hover { border-color: #FF9900 !important; }

    /* ── Tabs ── */
    button[data-baseweb="tab"] { color: #1976D2 !important; font-weight: 600 !important; }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #1565C0 !important;
        border-bottom-color: #FF9900 !important;
    }

    /* ── Divider ── */
    hr { border-color: #90CAF9 !important; }

    /* ── General text ── */
    p, span, label { color: #0D2B4E; }
    h1, h2, h3, h4 { color: #0D2B4E !important; }

    /* ── Progress bar ── */
    div[data-testid="stProgressBar"] > div > div {
        background: linear-gradient(90deg, #1A6DAF, #FF9900) !important;
    }

    /* ── Download button ── */
    div[data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #FF9900, #E68A00) !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 10px !important;
        font-size: 16px !important;
        padding: 14px !important;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background: linear-gradient(135deg, #FFB84D, #FF9900) !important;
    }

    /* ── Alert / info ── */
    div[data-testid="stAlert"] {
        background: #E3F2FD !important;
        border-color: #64B5F6 !important;
        color: #0D2B4E !important;
        border-radius: 10px !important;
    }

    /* ── Scrollbar ── */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #EBF5FB; }
    ::-webkit-scrollbar-thumb { background: #64B5F6; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #1976D2; }

    /* ── Chart white box ── */
    .chart-box {
        background: #ffffff;
        border-radius: 14px;
        padding: 16px;
        border: 1px solid #BBDEFB;
        box-shadow: 0 2px 8px rgba(25,118,210,0.07);
        margin-bottom: 4px;
    }

    /* ── Payer info card ── */
    .payer-card {
        background: linear-gradient(135deg, #E3F2FD, #BBDEFB);
        border-radius: 12px;
        padding: 16px 20px;
        border-left: 5px solid #1565C0;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 20px;
    }
    .payer-logo {
        background: white;
        border-radius: 10px;
        padding: 8px 14px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        flex-shrink: 0;
    }
    .payer-logo img { width: 80px; height: auto; display: block; }

    /* ── Report selector ── */
    .report-selector {
        background: #ffffff;
        border-radius: 14px;
        padding: 18px 22px;
        border: 2px solid #90CAF9;
        box-shadow: 0 2px 12px rgba(25,118,210,0.1);
        margin-bottom: 20px;
    }
    .report-selector-title {
        font-size: 14px; font-weight: 700; color: #1565C0;
        text-transform: uppercase; letter-spacing: 0.5px;
        margin-bottom: 10px;
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

@st.cache_data(show_spinner=False)
def parse_pdf(pdf_bytes):
    account_rows = []
    service_data = {}

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    # ── Extract payer info from invoice header ──
    payer_account = None
    payer_name    = None
    acct_m = re.search(r'Account number[:\s]*\n?\s*(\d{12})', full_text)
    if acct_m:
        payer_account = acct_m.group(1).strip()
    bill_m = re.search(r'Bill to Address[:\s]*\n\s*(.+)', full_text)
    if bill_m:
        payer_name = bill_m.group(1).strip()

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

    return account_rows, service_data, payer_account, payer_name


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
             alt="AWS Logo" style="width:130px;height:auto;display:block;"/>
    </div>
    <div class="hero-text">
        <h1 style="color:#FFFFFF;font-size:32px;font-weight:700;margin:0 0 4px 0;">Invoice Dashboard</h1>
        <p style="color:#BDD9F2;font-size:14px;margin:0;">
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
        <span style="color:#0D2B4E;font-size:13px;">
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
all_rows  = []
all_svcs  = {}
errors    = []
payer_map = {}   # payer_account_id -> {account_id, name, rows, svcs}

progress_bar = st.progress(0, text="Reading PDF files...")

for i, f in enumerate(uploaded_files):
    try:
        rows, svcs, p_acct, p_name = parse_pdf(f.read())
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

        # Build payer map — key = payer account number (or fallback)
        payer_key = p_acct or f"file_{i}"
        if payer_key not in payer_map:
            payer_map[payer_key] = {
                'account_id': p_acct or "Unknown",
                'name': p_name or f.name.replace('.pdf',''),
                'rows': [],
                'svcs': {},
            }
        payer_map[payer_key]['rows'].extend(rows)
        for acc_id, acc_data in svcs.items():
            pm = payer_map[payer_key]['svcs']
            if acc_id not in pm:
                pm[acc_id] = acc_data.copy()
            else:
                for k, v in acc_data.items():
                    if k == '_name':
                        pm[acc_id]['_name'] = v
                    else:
                        pm[acc_id][k] = round(pm[acc_id].get(k,0.0) + v, 2)

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
# REPORT SELECTOR  (only when multiple payers)
# ─────────────────────────────────────────────
selected_payer_key  = None
selected_payer_info = None

if len(payer_map) > 1:
    st.markdown('<div class="section-title">📑 Select Report View</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-selector">', unsafe_allow_html=True)
    st.markdown('<div class="report-selector-title">🗂️ Choose which payer\'s report to view</div>',
                unsafe_allow_html=True)

    payer_options = {"🌐 All Payers (Combined Report)": None}
    for pk, pv in payer_map.items():
        label = f"🏢  {pv['name']}   |   Account: {pv['account_id']}"
        payer_options[label] = pk

    chosen_label = st.selectbox(
        "Payer",
        options=list(payer_options.keys()),
        label_visibility="collapsed"
    )
    selected_payer_key = payer_options[chosen_label]
    st.markdown('</div>', unsafe_allow_html=True)

    if selected_payer_key is not None:
        selected_payer_info = payer_map[selected_payer_key]
        display_rows = selected_payer_info['rows']
        display_svcs = selected_payer_info['svcs']
    else:
        display_rows = all_rows
        display_svcs = all_svcs
else:
    display_rows = all_rows
    display_svcs = all_svcs
    if payer_map:
        selected_payer_info = list(payer_map.values())[0]

# ── Payer info banner when individual selected ──
if selected_payer_info and len(payer_map) > 1 and selected_payer_key is not None:
    p = selected_payer_info
    total_bill_payer = sum(r.get('Total (USD)', 0) for r in p['rows'])
    st.markdown(f"""
    <div class="payer-card">
        <div class="payer-logo">
            <img src="https://upload.wikimedia.org/wikipedia/commons/9/93/Amazon_Web_Services_Logo.svg"/>
        </div>
        <div>
            <h3 style="margin:0 0 5px 0;color:#1565C0;font-size:17px;">📋 {p['name']}</h3>
            <p style="margin:0;color:#0D2B4E;font-size:13px;">
                🔑 Payer Account: <b>{p['account_id']}</b> &nbsp;|&nbsp;
                🏢 Linked Accounts: <b>{len(p['rows'])}</b> &nbsp;|&nbsp;
                💰 Total Bill: <b>${total_bill_payer:,.2f}</b>
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# BUILD DATAFRAMES
# ─────────────────────────────────────────────
df_accounts = pd.DataFrame(display_rows)

all_services = sorted({s for d in display_svcs.values() for s in d if s != '_name'})
svc_rows = []
for acc_id, acc_data in display_svcs.items():
    row = {'Account ID': acc_id, 'Account Name': acc_data.get('_name', '')}
    for svc in all_services:
        row[svc] = acc_data.get(svc, 0.0) or 0.0
    svc_rows.append(row)
df_services = pd.DataFrame(svc_rows)

svc_totals = {}
for acc_data in display_svcs.values():
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

kpi(c1, "🏢", "Total Accounts",  str(len(display_rows)), f"{len(uploaded_files)} PDF(s)")
kpi(c2, "💰", "Total Charges",   f"${total_charges:,.2f}", "Before discounts")
kpi(c3, "🧾", "Total Bill",      f"${total_bill:,.2f}",    "After all discounts")
kpi(c4, "🏷️", "Total Discount",  f"${abs(total_discount):,.2f}", "Savings + EDP")
kpi(c5, "🔝", "Top Service",     top_service[:20]+"…" if len(top_service)>20 else top_service,
              f"${top_svc_cost:,.2f}")

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CHARTS ROW 1 — Pie + Bar  (white bg, sky blue)
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📈 Service Cost Analysis</div>', unsafe_allow_html=True)

AWS_COLORS = [
    '#FF9900','#29B6F6','#26C6DA','#42A5F5','#7E57C2',
    '#26A69A','#FFA726','#66BB6A','#EF5350','#5C6BC0','#EC407A'
]

chart_col1, chart_col2 = st.columns([1, 1])

with chart_col1:
    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
    st.markdown("#### 🥧 Service-wise Cost Share")
    top_n = 10
    if len(df_svc_totals) > top_n:
        top_df    = df_svc_totals.head(top_n).copy()
        other_sum = df_svc_totals.iloc[top_n:]['Total Cost (USD)'].sum()
        others    = pd.DataFrame([{'Service': 'Others', 'Total Cost (USD)': round(other_sum, 2)}])
        pie_df    = pd.concat([top_df, others], ignore_index=True)
    else:
        pie_df = df_svc_totals.copy()

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
        marker=dict(line=dict(color='#ffffff', width=2)),
        hovertemplate='<b>%{label}</b><br>Cost: $%{value:,.2f}<br>Share: %{percent}<extra></extra>'
    )
    fig_pie.update_layout(
        paper_bgcolor='#ffffff',
        plot_bgcolor='#ffffff',
        font=dict(color='#0D2B4E', family='Arial'),
        margin=dict(t=10, b=10, l=10, r=10),
        legend=dict(orientation="v", x=1.02, y=0.5,
                    bgcolor='rgba(0,0,0,0)', font=dict(color='#0D2B4E')),
        height=380,
    )
    st.plotly_chart(fig_pie, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with chart_col2:
    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
    st.markdown("#### 📊 Top Services by Cost (USD)")
    bar_df = df_svc_totals.head(12)
    fig_bar = px.bar(
        bar_df,
        x='Total Cost (USD)',
        y='Service',
        orientation='h',
        color='Total Cost (USD)',
        color_continuous_scale=[[0,'#87CEEB'],[0.5,'#29B6F6'],[1,'#FF9900']],
        text='Total Cost (USD)',
    )
    fig_bar.update_traces(
        texttemplate='$%{text:,.0f}',
        textposition='outside',
        textfont=dict(color='#0D2B4E', size=11),
        marker=dict(line=dict(color='#ffffff', width=0.5)),
        hovertemplate='<b>%{y}</b><br>Cost: $%{x:,.2f}<extra></extra>'
    )
    fig_bar.update_layout(
        paper_bgcolor='#ffffff',
        plot_bgcolor='#ffffff',
        font=dict(color='#0D2B4E', family='Arial'),
        yaxis=dict(autorange="reversed", color='#1565C0', gridcolor='#E3F2FD'),
        xaxis=dict(color='#1565C0', gridcolor='#E3F2FD'),
        coloraxis_showscale=False,
        margin=dict(t=10, b=10, l=10, r=110),
        height=380,
        xaxis_title="Cost (USD)",
        yaxis_title=""
    )
    st.plotly_chart(fig_bar, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CHARTS ROW 2 — Account Comparison + Donut
# ─────────────────────────────────────────────
if len(display_rows) > 1:
    st.markdown('<div class="section-title">🏢 Account-wise Comparison</div>', unsafe_allow_html=True)
    acct_col1, acct_col2 = st.columns([3, 2])

    with acct_col1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        st.markdown("#### 💳 Account-wise Total Bill")
        fig_acct = px.bar(
            df_accounts.sort_values('Total (USD)', ascending=False),
            x='Account Name',
            y='Total (USD)',
            color='Total (USD)',
            color_continuous_scale=[[0,'#87CEEB'],[0.5,'#29B6F6'],[1,'#FF9900']],
            text='Total (USD)',
        )
        fig_acct.update_traces(
            texttemplate='$%{text:,.0f}',
            textposition='outside',
            textfont=dict(color='#0D2B4E'),
            marker=dict(line=dict(color='#ffffff', width=0.5)),
        )
        fig_acct.update_layout(
            paper_bgcolor='#ffffff', plot_bgcolor='#ffffff',
            font=dict(color='#0D2B4E'),
            coloraxis_showscale=False,
            xaxis=dict(tickangle=-35, color='#1565C0', gridcolor='#E3F2FD'),
            yaxis=dict(color='#1565C0', gridcolor='#E3F2FD'),
            margin=dict(t=10, b=80),
            height=340,
        )
        st.plotly_chart(fig_acct, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with acct_col2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        st.markdown("#### 🍩 Account Cost Share")
        fig_acct_pie = px.pie(
            df_accounts,
            names='Account Name',
            values='Total (USD)',
            color_discrete_sequence=['#FF9900','#29B6F6','#26C6DA','#EF5350','#7E57C2','#42A5F5'],
            hole=0.5,
        )
        fig_acct_pie.update_traces(
            textposition='inside',
            textinfo='percent+label',
            textfont=dict(color='white', size=11),
            marker=dict(line=dict(color='#ffffff', width=2)),
        )
        fig_acct_pie.update_layout(
            paper_bgcolor='#ffffff',
            font=dict(color='#0D2B4E'),
            margin=dict(t=10,b=10,l=10,r=10),
            height=340,
            showlegend=False,
        )
        st.plotly_chart(fig_acct_pie, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

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
for acc_data in display_svcs.values():
    for k, v in acc_data.items():
        if k != '_name' and isinstance(v, (int,float)) and v > 0:
            svc_account_count[k] = svc_account_count.get(k, 0) + 1

margin_df = df_svc_totals.copy()
margin_df['Accounts Using'] = margin_df['Service'].map(lambda s: svc_account_count.get(s, 0))
margin_df['Avg Cost/Account (USD)'] = (margin_df['Total Cost (USD)'] / margin_df['Accounts Using']).round(2)

total_cost_sum = margin_df['Total Cost (USD)'].sum()
margin_df['Cost Share (%)'] = (margin_df['Total Cost (USD)'] / total_cost_sum * 100).round(1)

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

if selected_payer_key and selected_payer_info:
    payer_safe = re.sub(r'[^\w\s-]', '', selected_payer_info['name'])[:30].strip()
    dl_label = f"⬇️  Download Report — {selected_payer_info['name']}.xlsx"
    dl_fname = f"AWS_Report_{selected_payer_info['account_id']}.xlsx"
else:
    dl_label = "⬇️  Download AWS_Invoice_Report.xlsx  (All Payers)"
    dl_fname = "AWS_Invoice_Report_All.xlsx"

excel_buf = build_excel(display_rows, display_svcs)
st.download_button(
    label=dl_label,
    data=excel_buf,
    file_name=dl_fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

st.markdown(
    "<div style='text-align:center;color:#1565C0;font-size:12px;margin-top:16px;padding:12px;"
    "border-top:1px solid #90CAF9;'>"
    "☁️ &nbsp;AWS Invoice Dashboard &nbsp;•&nbsp; Built with Streamlit + Plotly &nbsp;•&nbsp; "
    "<span style='color:#FF9900;'>Powered by AWS Analytics</span>"
    "</div>",
    unsafe_allow_html=True
)

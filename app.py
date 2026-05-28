import streamlit as st
import pdfplumber
import re
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AWS PDF Invoice Extractor",
    layout="wide"
)

st.title("AWS PDF Invoice Extractor")

# ── Styling ──────────────────────────────────────────────────────────────────
HDR_BG    = "1F4E79"
HDR_FG    = "FFFFFF"
ALT_BG    = "D6E4F0"
NEG_FG    = "C00000"
TOTAL_BG  = "BDD7EE"
TOTAL_FG  = "1F4E79"


def _border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def _hdr(cell, value):
    cell.value = value
    cell.font = Font(bold=True, color=HDR_FG, name="Arial", size=10)
    cell.fill = PatternFill("solid", start_color=HDR_BG)
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )
    cell.border = _border()


def _dat(cell, value, row_idx, align="left", neg=False):
    cell.value = value

    bg = ALT_BG if row_idx % 2 == 0 else "FFFFFF"

    cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center")

    color = (
        NEG_FG
        if neg and isinstance(value, (int, float)) and value < 0
        else "000000"
    )

    cell.font = Font(name="Arial", size=10, color=color)
    cell.border = _border()


def _total_cell(cell, value, align="right"):
    cell.value = value
    cell.font = Font(
        bold=True,
        color=TOTAL_FG,
        name="Arial",
        size=10
    )
    cell.fill = PatternFill("solid", start_color=TOTAL_BG)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = _border()


def _auto_width(ws, min_w=10, max_w=35):

    for col in ws.columns:

        length = max(
            (len(str(c.value)) if c.value is not None else 0)
            for c in col
        )

        ws.column_dimensions[
            get_column_letter(col[0].column)
        ].width = max(min_w, min(length + 3, max_w))


def _freeze_filter(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ── PDF Parsing ──────────────────────────────────────────────────────────────

_SAVINGS_SUB = re.compile(
    r'^Savings Plan \(Charges covered by Savings Plans\)'
)

_SUB_STARTS = (
    'Charges ',
    'Tax ',
    'GST ',
    'Credits ',
    'Discount ('
)


def _usd(text):

    m = re.search(r'-?USD\s*([\d,]+\.?\d*)', text)

    if not m:
        return 0.0

    val = float(m.group(1).replace(',', ''))

    return (
        -val
        if re.search(r'(?<!\w)-USD', text)
        or text.strip().startswith('-')
        else val
    )


def _is_sub(line):

    if _SAVINGS_SUB.match(line):
        return True

    return any(line.startswith(p) for p in _SUB_STARTS)


def parse_pdf(uploaded_file):

    account_rows = []
    service_data = {}

    with pdfplumber.open(uploaded_file) as pdf:

        full_text = "\n".join(
            (p.extract_text() or "")
            for p in pdf.pages
        )

    blocks = re.split(
        r'Summary for Linked Account\s*\n',
        full_text
    )

    for block in blocks[1:]:

        lines = [
            l.strip()
            for l in block.split('\n')
            if l.strip()
        ]

        if not lines:
            continue

        hm = re.match(
            r'^(.+?)\s*\((\d{12})\)\s+USD\s+([\d,]+\.?\d*)',
            lines[0]
        )

        if not hm:
            continue

        acc_name = hm.group(1).strip()
        acc_id = hm.group(2).strip()

        charges = 0.0
        savings_plan = 0.0
        private_rate = 0.0
        bundled_dis = 0.0
        edp_dis = 0.0
        credits = 0.0
        gst = 0.0
        total_alloc = 0.0

        in_summary = True
        in_detail = False

        svc_name = None
        svc_charges = 0.0
        svc_savings = 0.0
        cur_svcs = {}

        for line in lines[1:]:

            if 'total allocated for this statement' in line.lower():

                m = re.search(
                    r'USD\s*([\d,]+\.?\d*)',
                    line
                )

                total_alloc = (
                    float(m.group(1).replace(',', ''))
                    if m else 0.0
                )

                in_summary = False
                continue

            if 'Detail for Linked Account' in line:
                in_detail = True
                in_summary = False
                continue

            if (
                'For line item details' in line
                or 'Account Activity Page' in line
            ):
                break

            # ── Summary ────────────────────────────────────────────────────
            if in_summary:

                if re.match(r'^Charges\s+USD', line):
                    charges = abs(_usd(line))

                elif _SAVINGS_SUB.match(line):
                    savings_plan = abs(_usd(line))

                elif 'Private Rate Card' in line:
                    private_rate = abs(_usd(line))

                elif 'Bundled Discount' in line:
                    bundled_dis = abs(_usd(line))

                elif (
                    'Distribution Program Discount' in line
                    or 'Enterprise Discount Program' in line
                    or 'AWS Distribution Program Discount' in line
                ):
                    edp_dis = abs(_usd(line))

                elif re.match(r'^Credits\s+', line):
                    credits = _usd(line)

                elif (
                    re.match(r'^Tax\s+USD', line)
                    or re.match(r'^GST\s+USD', line)
                ):
                    gst += abs(_usd(line))

            # ── Detail ─────────────────────────────────────────────────────
            elif in_detail:

                is_sub = _is_sub(line)

                svc_match = re.match(
                    r'^(.+?)\s+USD\s+([\d,]+\.?\d*)$',
                    line
                )

                if svc_match and not is_sub:

                    if svc_name:

                        net = svc_charges - svc_savings

                        cur_svcs[svc_name] = (
                            cur_svcs.get(svc_name, 0.0) + net
                        )

                    svc_name = svc_match.group(1).strip()
                    svc_charges = 0.0
                    svc_savings = 0.0

                elif svc_name:

                    if re.match(r'^Charges\s+USD', line):
                        svc_charges = abs(_usd(line))

                    elif _SAVINGS_SUB.match(line):
                        svc_savings = abs(_usd(line))

        # Save last service
        if svc_name and in_detail:

            cur_svcs[svc_name] = (
                cur_svcs.get(svc_name, 0.0)
                + (svc_charges - svc_savings)
            )

        account_rows.append({

            'Account ID': acc_id,
            'Account Name': acc_name,
            'Charges (USD)': round(charges, 2),
            'Savings Plan (USD)': round(-savings_plan, 2),
            'Private Rate Card (USD)': round(-private_rate, 2),
            'Bundled Discount (USD)': round(-bundled_dis, 2),
            'EDP / Dist. Discount (USD)': round(-edp_dis, 2),
            'Credits (USD)': round(credits, 2),
            'GST (USD)': round(gst, 2),
            'Total (USD)': round(total_alloc, 2),
        })

        if acc_id not in service_data:
            service_data[acc_id] = {'_name': acc_name}

        for svc, val in cur_svcs.items():

            service_data[acc_id][svc] = round(
                service_data[acc_id].get(svc, 0.0) + val,
                2
            )

    return account_rows, service_data


# ── Excel Writer ─────────────────────────────────────────────────────────────

S1_HDRS = [
    'Account ID',
    'Account Name',
    'Charges (USD)',
    'Savings Plan (USD)',
    'Private Rate Card (USD)',
    'Bundled Discount (USD)',
    'EDP / Dist. Discount (USD)',
    'Credits (USD)',
    'GST (USD)',
    'Total (USD)',
]

S1_NEG = {4, 5, 6, 7, 8}
S1_NUM = set(range(3, len(S1_HDRS) + 1))


def write_excel(all_rows, all_svcs):

    output = BytesIO()

    wb = Workbook()

    # ── Sheet 1 ────────────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Account Details"

    for ci, h in enumerate(S1_HDRS, 1):
        _hdr(ws1.cell(1, ci), h)

    for ri, row in enumerate(all_rows, 2):

        for ci, h in enumerate(S1_HDRS, 1):

            val = row[h]

            _dat(
                ws1.cell(ri, ci),
                val,
                ri,
                align="right" if ci in S1_NUM else "left",
                neg=(ci in S1_NEG)
            )

    total_ri = len(all_rows) + 2

    _total_cell(
        ws1.cell(total_ri, 1),
        "TOTAL",
        align="center"
    )

    for ci, h in enumerate(S1_HDRS[2:], 3):

        col_vals = [
            r[h]
            for r in all_rows
            if isinstance(r[h], (int, float))
        ]

        _total_cell(
            ws1.cell(total_ri, ci),
            round(sum(col_vals), 2)
        )

    _auto_width(ws1)
    _freeze_filter(ws1)

    # ── Sheet 2 ────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Service Breakdown")

    all_services = sorted({
        s
        for d in all_svcs.values()
        for s in d
        if s != '_name'
    })

    s2_hdrs = ['Account ID', 'Account Name'] + all_services

    for ci, h in enumerate(s2_hdrs, 1):
        _hdr(ws2.cell(1, ci), h)

    for ri, (acc_id, acc_data) in enumerate(all_svcs.items(), 2):

        _dat(ws2.cell(ri, 1), acc_id, ri)
        _dat(ws2.cell(ri, 2), acc_data.get('_name', ''), ri)

        for ci, svc in enumerate(all_services, 3):

            val = acc_data.get(svc, None)

            display = (
                round(val, 2)
                if val is not None
                else None
            )

            _dat(
                ws2.cell(ri, ci),
                display,
                ri,
                align="right"
            )

    _auto_width(ws2)
    _freeze_filter(ws2)

    wb.save(output)

    output.seek(0)

    return output


# ── STREAMLIT UI ─────────────────────────────────────────────────────────────

uploaded_files = st.file_uploader(
    "Upload AWS PDF Files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:

    all_rows = []
    all_svcs = {}

    progress = st.progress(0)

    status = st.empty()

    total = len(uploaded_files)

    for i, uploaded_file in enumerate(uploaded_files, 1):

        status.info(
            f"Processing {i}/{total}: {uploaded_file.name}"
        )

        rows, svcs = parse_pdf(uploaded_file)

        all_rows.extend(rows)

        for acc_id, acc_data in svcs.items():

            if acc_id not in all_svcs:

                all_svcs[acc_id] = acc_data.copy()

            else:

                for k, v in acc_data.items():

                    if k == '_name':
                        all_svcs[acc_id]['_name'] = v

                    else:

                        all_svcs[acc_id][k] = round(
                            all_svcs[acc_id].get(k, 0.0) + v,
                            2
                        )

        progress.progress(int(i / total * 100))

    status.success("Processing Complete!")

    st.success(
        f"Total Accounts Extracted: {len(all_rows)}"
    )

    # Preview
    st.subheader("Account Details")

    st.dataframe(
        all_rows,
        use_container_width=True
    )

    # Excel
    excel_file = write_excel(all_rows, all_svcs)

    st.download_button(
        label="Download Excel Report",
        data=excel_file,
        file_name="AWS_Invoice_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
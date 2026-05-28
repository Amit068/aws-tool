import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="AWS PDF Extractor", layout="wide")

st.title("AWS PDF to Excel Extractor")

uploaded_file = st.file_uploader(
    "PDF Upload Karo",
    type=["pdf"]
)

if uploaded_file:

    st.info("PDF read ho rahi hai...")

    data = []
    current_account = None
    current_account_id = None

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:

            text = page.extract_text()

            if not text:
                continue

            lines = text.split('\n')

            for line in lines:

                line = line.strip()

                # Account Name + Account ID
                account_match = re.match(
                    r'^(.+?)\s*\((\d{12})\)\s+USD\s+([\d,]+\.?\d*)',
                    line
                )

                if account_match:

                    current_account = account_match.group(1).strip()
                    current_account_id = account_match.group(2).strip()

                    charges = 0
                    savings = 0
                    bundled_dis = 0
                    enterprise_dis = 0
                    credits = 0
                    gst = 0

                    continue

                # Charges
                if re.match(r'^Charges\s+USD\s+([\d,]+\.?\d*)', line) and current_account:

                    m = re.search(r'USD\s+([\d,]+\.?\d*)', line)

                    if m:
                        charges = float(m.group(1).replace(',', ''))

                # Savings Plan
                elif re.match(r'^Savings Plan.*USD\s*-?([\d,]+\.?\d*)', line) and current_account:

                    m = re.search(r'USD\s*-?([\d,]+\.?\d*)', line)

                    if m:
                        savings = float(m.group(1).replace(',', ''))

                # Bundled Discount
                elif 'Bundled Discount' in line and current_account:

                    m = re.search(r'USD\s*-?([\d,]+\.?\d*)', line)

                    if m:
                        bundled_dis = float(m.group(1).replace(',', ''))

                # Enterprise Discount
                elif 'Distribution Program Discount' in line and current_account:

                    m = re.search(r'USD\s*-?([\d,]+\.?\d*)', line)

                    if m:
                        enterprise_dis = float(m.group(1).replace(',', ''))

                # Credits
                elif re.match(r'^Credits\s+USD', line) and current_account:

                    m = re.search(r'USD\s*([\d,]+\.?\d*)', line)

                    if m:
                        credits = float(m.group(1).replace(',', ''))

                # GST
                elif re.match(r'^GST\s+USD', line) and current_account:

                    m = re.search(r'USD\s*([\d,]+\.?\d*)', line)

                    if m:
                        gst = float(m.group(1).replace(',', ''))

                # Final row save
                elif 'total allocated for this statement' in line and current_account:

                    m = re.search(r'USD\s*([\d,]+\.?\d*)', line)

                    total_final = float(m.group(1).replace(',', '')) if m else 0

                    data.append({
                        'Account ID': current_account_id,
                        'Account Name': current_account,
                        'Charges (USD)': charges,
                        'Savings Plan (USD)': -savings,
                        'Bundled Discount (USD)': -bundled_dis,
                        'Enterprise Discount (USD)': -enterprise_dis,
                        'Credits (USD)': credits,
                        'GST (USD)': gst,
                        'Total (USD)': total_final
                    })

                    current_account = None
                    current_account_id = None

    if data:

        df = pd.DataFrame(data)

        st.success(f"Total accounts extracted: {len(data)}")

        st.dataframe(df, use_container_width=True)

        # Excel in memory
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:

            df.to_excel(writer, sheet_name='Account Data', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Account Data']

            # Auto width
            for col in worksheet.columns:

                max_len = max(len(str(cell.value)) for cell in col if cell.value)

                worksheet.column_dimensions[col[0].column_letter].width = max_len + 5

        output.seek(0)

        st.download_button(
            label="Download Excel",
            data=output,
            file_name="aws_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("Koi data nahi mila!")
import pdfplumber
import pandas as pd
import re
from tkinter import filedialog, Tk

# PDF file select karo
root = Tk()
root.withdraw()
pdf_path = filedialog.askopenfilename(
    title="PDF Select Karo",
    filetypes=[("PDF Files", "*.pdf")]
)

if not pdf_path:
    print("Koi file select nahi ki!")
    exit()

print("PDF read ho rahi hai...")

data = []
current_account = None
current_account_id = None

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if not text:
            continue
        
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            
            # Account Name (AccountId) pattern dhundo
            # Pattern: "AccountName (123456789012) USD X,XXX.XX"
            account_match = re.match(
                r'^(.+?)\s*\((\d{12})\)\s+USD\s+([\d,]+\.?\d*)', line
            )
            
            if account_match:
                current_account = account_match.group(1).strip()
                current_account_id = account_match.group(2).strip()
                total = float(account_match.group(3).replace(',', ''))
                
                # Reset values
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
            
            # Enterprise/Distribution Discount
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
            
            # "Account XXXX total allocated" — row save karo
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

# Excel file banao
if data:
    df = pd.DataFrame(data)
    
    # Output file same folder mein save karo
    output_path = pdf_path.replace('.pdf', '_extracted.xlsx')
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Account Data', index=False)
        
        # Formatting
        workbook = writer.book
        worksheet = writer.sheets['Account Data']
        
        # Column width
        for col in worksheet.columns:
            max_len = max(len(str(cell.value)) for cell in col if cell.value)
            worksheet.column_dimensions[col[0].column_letter].width = max_len + 5
    
    print(f"\nDone! File saved: {output_path}")
    print(f"Total accounts extracted: {len(data)}")
else:
    print("Koi data nahi mila!")

input("\nEnter dabao band karne ke liye...")
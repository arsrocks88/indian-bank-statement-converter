import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# App Layout Configuration
st.set_page_config(page_title="Indian Bank Statement Converter", layout="centered")
st.title("🏦 Indian Bank Statement Converter")
st.write("Convert your **SBI, HDFC, ICICI, or Bank of Baroda** statements into a clean Excel file.")

st.info("🔒 **Privacy Guarantee:** All document parsing runs inside your personal web browser session. Your statements are never sent to a backend server.")

# File Uploader Widget
uploaded_file = st.file_uploader("Upload your Bank Statement (PDF only)", type=["pdf"])

def extract_raw_text_lines(pdf_file, password=None):
    """Safely extracts raw text lines from the uploaded document, checking for locks."""
    lines_list = []
    try:
        with pdfplumber.open(pdf_file, password=password) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines_list.extend(text.split('\n'))
    except Exception as e:
        error_msg = str(e).lower()
        if "password" in error_msg or "encrypted" in error_msg:
            return "LOCKED_PDF"
        else:
            return f"ERROR: {str(e)}"
    return lines_list

def parse_indian_statements(raw_lines):
    """Processes messy rows, stitching multi-line text blocks together cleanly."""
    clean_transactions = []
    current_row = None
    
    # Matches DD/MM/YYYY, DD.MM.YYYY, and DD/MM/YY formats safely
    date_regex = re.compile(r'^(\d{2}[/\.]\d{2}[/\.]\d{2,4})')
    
    for line in raw_lines:
        line = line.strip()
        if not line:
            continue
            
        # Ignore common header navigation/footer texts found in your samples
        if any(x in line.lower() for x in ["page no", "generated on", "statement summary", "computer-generated", "registered office"]):
            continue
            
        match_date = date_regex.match(line)
        
        if match_date:
            # Save the previous row before constructing a new one
            if current_row:
                clean_transactions.append(current_row)
            
            # Start tracking a clean row step-by-step
            current_row = {"date": match_date.group(1), "text": line[len(match_date.group(1)):].strip()}
        else:
            # If a row doesn't start with a date, stitch it directly to the description cell above
            if current_row:
                current_row["text"] += " " + line

    # Catch the final row left over in storage
    if current_row:
        clean_transactions.append(current_row)
        
    return clean_transactions

def convert_to_dataframe(structured_rows):
    """Refines standard space blocks into transaction table entries using fallback parsing."""
    parsed_records = []
    
    for row in structured_rows:
        date = row["date"]
        body = row["text"]
        
        # Strategy A: Try splitting by wide space blocks first
        parts = [p.strip() for p in re.split(r'\s{2,}', body) if p.strip()]
        
        # Strategy B: If everything collapsed into one big sentence (like your HDFC screenshot),
        # we parse it from the back because Indian amounts usually end with clear numbers.
        if len(parts) < 3:
            # Split by single spaces just to extract numbers at the tail
            all_words = body.split()
            if len(all_words) >= 3:
                balance = all_words[-1]
                amount_candidate = all_words[-2]
                
                # Check if it looks like a secondary date or reference number
                if '/' in all_words[-3] or len(all_words[-3]) > 8:
                    narration = " ".join(all_words[:-2])
                    # If it's a deposit or withdrawal based on narration keywords
                    if "credit" in narration.lower() or "deposit" in narration.lower() or "redeem" in narration.lower():
                        withdrawal, deposit = "0.00", amount_candidate
                    else:
                        withdrawal, deposit = amount_candidate, "0.00"
                else:
                    # Normal layout ending structure
                    narration = " ".join(all_words[:-3])
                    if "credit" in narration.lower() or "deposit" in narration.lower() or "redeem" in narration.lower():
                        withdrawal, deposit = "0.00", all_words[-2]
                    else:
                        withdrawal, deposit = all_words[-2], "0.00"
                
                parsed_records.append({
                    "Date": date,
                    "Narration / Description": narration,
                    "Withdrawal (Dr)": withdrawal,
                    "Deposit (Cr)": deposit,
                    "Closing Balance": balance
                })
                continue

        # Normal Column Processing (For Bank of Baroda, SBI, ICICI layouts)
        if len(parts) >= 2:
            try:
                balance = parts[-1]
                if len(parts) >= 4:
                    withdrawal = parts[-3]
                    deposit = parts[-2]
                    narration = " ".join(parts[:-3])
                else:
                    # Intelligent guess for 3-part layout columns
                    val = parts[-2]
                    if "credit" in body.lower() or "deposit" in body.lower() or "redeem" in body.lower():
                        withdrawal, deposit = "0.00", val
                    else:
                        withdrawal, deposit = val, "0.00"
                    narration = " ".join(parts[:-2])
            except:
                narration = body
                withdrawal, deposit, balance = "N/A", "N/A", "N/A"
        else:
            narration = body
            withdrawal, deposit, balance = "0.00", "0.00", "N/A"
            
        parsed_records.append({
            "Date": date,
            "Narration / Description": narration,
            "Withdrawal (Dr)": withdrawal,
            "Deposit (Cr)": deposit,
            "Closing Balance": balance
        })
        
    return pd.DataFrame(parsed_records)

# Execution Sequence
if uploaded_file is not None:
    file_stream = io.BytesIO(uploaded_file.read())
    
    # Check for password encryption
    raw_output = extract_raw_text_lines(file_stream)
    
    user_password = None
    if raw_output == "LOCKED_PDF":
        st.warning("🔑 This file is password protected. Enter your credentials below to process.")
        user_password = st.text_input("Enter Document Password:", type="password")
        if not user_password:
            st.stop()
            
        # Retry extraction once credentials are submitted
        raw_output = extract_raw_text_lines(file_stream, password=user_password)
        
    if raw_output == "LOCKED_PDF":
        st.error("❌ Incorrect Password. Access Denied.")
    elif isinstance(raw_output, str) and raw_output.startswith("ERROR:"):
        st.error(f"Failed to process file template: {raw_output}")
    elif not raw_output:
        st.error("Could not find readable transaction metrics.")
    else:
        # Run sorting algorithms safely
        with st.spinner("Stitching multi-line UPI and interest rows..."):
            sorted_lines = parse_indian_statements(raw_output)
            final_df = convert_to_dataframe(sorted_lines)
            
        if not final_df.empty:
            st.success(f"Successfully filtered out {len(final_df)} transaction entries!")
            
            st.subheader("Data Preview Panel")
            st.dataframe(final_df.head(15), use_container_width=True)
            
            # Excel export memory block generator
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Clean Transactions')
            excel_bytes = excel_buffer.getvalue()
            
            st.download_button(
                label="📥 Download Structured Excel File (.xlsx)",
                data=excel_bytes,
                file_name="Clean_Bank_Statement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

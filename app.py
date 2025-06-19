# app.py
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import os

def extract_value_line_table_fixed_with_skipped(pdf_path):
    doc = fitz.open(pdf_path)
    all_data = []
    skipped_rows = []

    for page_number in range(start_page, end_page + 1):
        page = doc[page_number - 1]
        blocks = page.get_text("blocks", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
        blocks.sort(key=lambda b: (b[1], b[0]))
        lines = [block[4] for block in blocks if block[6] == 0]

        joined_lines = '\n'.join(lines).split('\n')
        records = []
        buffer = ""

        for line in joined_lines:
            line = line.strip()
            if re.match(r'^\d{3,4}\s+\d{3,4}', line):
                if buffer:
                    records.append(buffer.strip())
                buffer = re.sub(r'^\d{3,4}\s+(\d{3,4})\s+', r'\1 ', line)
            elif re.match(r'^\d{3,4}\s', line):
                if buffer:
                    records.append(buffer.strip())
                buffer = line
            else:
                buffer += " " + line
        if buffer:
            records.append(buffer.strip())

        value_with_suffix = r'[\u25C6\u25B2\u25BC]?(?:d)?(?:\d*\.\d+(?:-\d*\.\d+)?|\d+)(?:\([a-zA-Z]+\))?'

        pattern = re.compile(
            rf'^(\d{{3,4}})\s+'
            rf'(.+?)\s+'
            rf'([A-Z.&\'\-]{{1,10}})\s+'
            rf'(\d*\.\d{{2}}[a-zA-Z]?)\s+'
            rf'([\u25C6\u25B2\u25BC\-]?\d|\u2013|\u2014)\s+'  # – = — = en/em dash
            rf'([\u25C6\u25B2\u25BC\-]?\d|\u2013|\u2014)\s+'
            rf'([\u25C6\u25B2\u25BC\-]?\d|\u2013|\u2014)\s+'
            rf'(\d*\.\d+|\u2013|NMF)\s+'
            rf'[\u25C6\u25B2\u25BC]?\s*([\d\-]+\s*[\d\-]+)\s*'
            rf'(\(.*?\))\s*'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s+'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s+'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s+'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s+'
            rf'(\d+)\s+'
            rf'(\d{{1,2}}/\d{{2}})\s+'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s+'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s+'
            rf'(\d{{1,2}}/\d{{2}})\s+'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s+'
            rf'({value_with_suffix}|NIL|NA|NMF|\u2013|\u2014)\s*'
            rf'(YES|NO)?(?:\s+\S+)?\s*$'
        )

        headers = [
            "Page_Number", "Number", "Company", "Ticker", "Price", "Timeliness", "Safety", "Technical",
            "Beta", "Target_Price_Range", "Appreciation_Potential", "Current_PE", "Est_Yield_Pct",
            "Est_Earnings", "Est_Dividend", "Industry_Rank", "Quarter_Ended", "EPS_Latest",
            "EPS_Year_Ago", "Dividend_Qtr_Ended", "Latest_Dividend", "Year_Ago_Dividend", "Options_Traded"
        ]

        for rec in records:
            cleaned_rec = re.sub(r'\s+\(NDQ\)', '', rec)
            cleaned_rec = re.sub(r'(YES|NO)\s+.+$', r'\1', cleaned_rec)
            match = pattern.match(cleaned_rec)
            if match:
                row = [str(page_number)] + list(match.groups())
                all_data.append(row)
            else:
                if cleaned_rec.strip().endswith(("YES", "NO")):
                    skipped_rows.append((str(page_number), cleaned_rec))

    df = pd.DataFrame(all_data, columns=headers).astype(str)
    df['Company'] = df['Company'].str.replace(r'^\d{3,4}\s+', '', regex=True)
    df['Company'] = df['Company'].str.replace(r'\([A-Z]{2,4}\)', '', regex=True).str.strip()
    df['Company'] = df['Company'].str.replace(r'\s+', ' ', regex=True)

    skipped_header_row = ["Header", " ".join(headers)]
    skipped_content_rows = [["Page " + pg, line] for pg, line in skipped_rows]
    skipped_df = pd.DataFrame([skipped_header_row] + skipped_content_rows, columns=["Page", "Skipped_Line"])

    return df, skipped_df


st.title("📄 Value Line PDF Table Parser")

uploaded_file = st.file_uploader("Upload a Value Line PDF report", type="pdf")
if uploaded_file:
    with open("temp.pdf", "wb") as f:
        f.write(uploaded_file.read())

    df, skipped_df = extract_value_line_table_fixed_with_skipped("temp.pdf")

    st.success(f"✅ Successfully parsed {len(df)} rows")
    st.dataframe(df.head())

    output_filename = uploaded_file.name.replace(".pdf", "_output.xlsx")
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Parsed_Data", index=False)
        skipped_df.to_excel(writer, sheet_name="Skipped_Rows", index=False)

    with open(output_filename, "rb") as f:
        st.download_button(
            label="📥 Download Excel File",
            data=f,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    os.remove("temp.pdf")
    os.remove(output_filename)

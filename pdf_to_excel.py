import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# Set page config
st.set_page_config(page_title="Secure PDF Table Extractor", layout="wide")

# Custom CSS for better design
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
        font-family: 'Segoe UI', sans-serif;
    }
    h1 {
        color: #2c3e50;
    }
    .stButton button {
        background-color: #2980b9;
        color: white;
        border-radius: 6px;
        padding: 10px 20px;
    }
    .info-box {
        background-color: #ecf0f1;
        padding: 15px;
        border-left: 5px solid #2980b9;
        margin-bottom: 20px;
        font-size: 16px;
    }
    .footer {
        font-size: 14px;
        color: gray;
        text-align: center;
        margin-top: 50px;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.title("🔐 Secure PDF Table Extractor")
st.markdown("<p style='font-size:18px;'>Convert PDF tables into Excel — locally, securely, instantly.</p>", unsafe_allow_html=True)

# Privacy Banner
st.markdown('<div class="info-box">⚠️ This app runs entirely locally. No files are uploaded, stored, or shared.</div>', unsafe_allow_html=True)

# Sidebar for settings
with st.sidebar:
    st.header("🛠️ Settings")
    combine_tables = st.checkbox("✅ Combine all selected pages into one sheet", value=True)
    st.info("Select specific pages below to extract tables from.")

def deduplicate_columns(columns):
    """Handles duplicate or empty column names."""
    seen = {}
    for idx, col in enumerate(columns):
        if not col or col.strip() == '':
            col = 'Unnamed'
        if col in seen:
            seen[col] += 1
            columns[idx] = f"{col}_{seen[col]}"
        else:
            seen[col] = 1
    return columns

def is_valid_table(table_data, min_rows=2, min_cols=2):
    """
    Heuristic to determine if the extracted content is a valid table.
    """
    if not table_data or len(table_data) < min_rows:
        return False

    col_lengths = [len(row) for row in table_data]
    majority_length = max(set(col_lengths), key=col_lengths.count)
    majority_count = col_lengths.count(majority_length)

    if majority_count / len(col_lengths) < 0.7:
        return False

    if majority_length < min_cols:
        return False

    return True


def extract_tables_from_pdf(file, selected_pages):
    tables = []
    with pdfplumber.open(file) as pdf:
        total_pages = len(pdf.pages)
        progress_bar = st.progress(0)
        status_text = st.empty()

        for page_num in range(total_pages):
            if (page_num + 1) not in selected_pages:
                continue

            status_text.text(f"🔍 Processing page {page_num + 1}...")
            try:
                page = pdf.pages[page_num]

                # Try basic table extraction
                table = page.extract_table()
                if not table:
                    # Fallback strategy: try more aggressive detection
                    tables_on_page = page.extract_tables({
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "text"
                    })
                    if tables_on_page:
                        table = tables_on_page[0]

                if table and is_valid_table(table):
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df.columns = deduplicate_columns(df.columns.tolist())
                    tables.append((page_num + 1, df))
                else:
                    st.info(f"❌ Skipped non-table content on page {page_num + 1}")
            except Exception as e:
                st.warning(f"⚠️ Could not extract table from page {page_num + 1}: {str(e)}")
            progress_bar.progress((page_num + 1) / total_pages)
        status_text.text("✅ Processing complete.")
    return tables

def convert_to_excel(tables, combine_sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if combine_sheets:
            combined_df = pd.concat([df for _, df in tables], ignore_index=True)
            combined_df.to_excel(writer, index=False, sheet_name="Combined")
        else:
            for page_num, df in tables:
                sheet_name = f"Page_{page_num}"[:31]
                df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# Main logic
if 'tables' not in st.session_state:
    st.session_state.tables = []

uploaded_file = st.file_uploader("📂 Upload your PDF file", type="pdf", key="uploader")

if uploaded_file is not None:
    with pdfplumber.open(uploaded_file) as pdf:
        total_pages = len(pdf.pages)

    st.success(f"📘 Your PDF has **{total_pages} pages**. Select which ones to extract from:")
    page_options = list(range(1, total_pages + 1))
    selected_pages = st.multiselect("📌 Choose Pages", options=page_options, default=page_options)

    if st.button("🚀 Start Extraction"):
        tables = extract_tables_from_pdf(uploaded_file, selected_pages)
        st.session_state.tables = tables

        if tables:
            st.success(f"✅ Successfully extracted tables from **{len(tables)} pages**.")

            # Summary stats
            total_rows = sum(len(df) for _, df in tables)
            total_cols = max(len(df.columns) for _, df in tables) if tables else 0
            st.markdown("📊 **Summary:**")
            st.markdown(f"- Total Tables Extracted: `{len(tables)}`")
            st.markdown(f"- Total Rows: `{total_rows}`")
            st.markdown(f"- Max Columns in Any Table: `{total_cols}`")

            # Preview tables
            for page_num, df in tables:
                with st.expander(f"📄 Table from Page {page_num}"):
                    st.dataframe(df, use_container_width=True)

            excel_data = convert_to_excel(tables, combine_tables)

            st.download_button(
                label="📥 Download Excel File",
                data=excel_data,
                file_name="extracted_tables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ No valid tables found on the selected pages.")

# Footer
st.markdown('<div class="footer">Built by [Your Name] - MIS Data Analyst | Internal Use Only</div>', unsafe_allow_html=True)

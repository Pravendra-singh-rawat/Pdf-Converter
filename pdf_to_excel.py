import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# Set Streamlit app title and layout
st.title("PDF to Excel Converter")
st.write("Upload a PDF file with tables to convert it into an Excel spreadsheet.")

# File uploader for PDF
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

def deduplicate_columns(columns):
    # Function to handle duplicate column names by appending suffixes
    seen = {}
    for idx, col in enumerate(columns):
        if col in seen:
            seen[col] += 1
            columns[idx] = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
    return columns

def extract_table_from_pdf(file):
    tables = []
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            try:
                # Extract tables from each page
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df.columns = deduplicate_columns(df.columns.tolist())  # Apply deduplication
                    tables.append((page_num + 1, df))
            except Exception as e:
                st.warning(f"Could not extract table from page {page_num + 1}: {e}")
    return tables

def convert_to_excel(tables):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for page_num, table in tables:
            sheet_name = f"Page_{page_num}"
            table.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# If file is uploaded
if uploaded_file is not None:
    tables = extract_table_from_pdf(uploaded_file)
    
    if tables:
        st.write(f"Extracted {len(tables)} tables.")
        
        # Display tables and convert to Excel
        for page_num, table in tables:
            st.write(f"Page {page_num}")
            st.write(table)

        # Convert to Excel
        excel_data = convert_to_excel(tables)

        # Download link
        st.download_button(
            label="Download Excel file",
            data=excel_data,
            file_name="converted_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No tables found in PDF.")










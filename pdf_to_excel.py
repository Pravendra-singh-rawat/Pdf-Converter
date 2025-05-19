# import streamlit as st
# import pdfplumber
# import pandas as pd
# from io import BytesIO

# # Set Streamlit app title and layout
# st.title("PDF to Excel Converter")
# st.write("Upload a PDF file with tables to convert it into an Excel spreadsheet.")

# # File uploader for PDF
# uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

# def deduplicate_columns(columns):
#     # Function to handle duplicate column names by appending suffixes
#     seen = {}
#     for idx, col in enumerate(columns):
#         if col in seen:
#             seen[col] += 1
#             columns[idx] = f"{col}_{seen[col]}"
#         else:
#             seen[col] = 0
#     return columns

# def extract_table_from_pdf(file):
#     tables = []
#     with pdfplumber.open(file) as pdf:
#         for page_num, page in enumerate(pdf.pages):
#             try:
#                 # Extract tables from each page
#                 table = page.extract_table()
#                 if table:
#                     df = pd.DataFrame(table[1:], columns=table[0])
#                     df.columns = deduplicate_columns(df.columns.tolist())  # Apply deduplication
#                     tables.append((page_num + 1, df))
#             except Exception as e:
#                 st.warning(f"Could not extract table from page {page_num + 1}: {e}")
#     return tables

# def convert_to_excel(tables):
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="openpyxl") as writer:
#         for page_num, table in tables:
#             sheet_name = f"Page_{page_num}"
#             table.to_excel(writer, index=False, sheet_name=sheet_name)
#     output.seek(0)
#     return output

# # If file is uploaded
# if uploaded_file is not None:
#     tables = extract_table_from_pdf(uploaded_file)
    
#     if tables:
#         st.write(f"Extracted {len(tables)} tables.")
        
#         # Display tables and convert to Excel
#         for page_num, table in tables:
#             st.write(f"Page {page_num}")
#             st.write(table)

#         # Convert to Excel
#         excel_data = convert_to_excel(tables)

#         # Download link
#         st.download_button(
#             label="Download Excel file",
#             data=excel_data,
#             file_name="converted_tables.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
#     else:
#         st.error("No tables found in PDF.")
















import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# Set page config
st.set_page_config(page_title="Advanced PDF Table Converter", layout="wide")

# Title and description
st.title("📄 Advanced PDF Table Extractor")
st.markdown("Upload a PDF file and convert its tables into Excel. You can now select specific pages and combine tables into a single sheet.")

# Sidebar for settings
with st.sidebar:
    st.header("Settings")
    combine_tables = st.checkbox("Combine all selected tables into one sheet", value=True)

# Upload PDF
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

def deduplicate_columns(columns):
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

def extract_tables_from_pdf(file, selected_pages):
    tables = []
    with pdfplumber.open(file) as pdf:
        total_pages = len(pdf.pages)
        progress_bar = st.progress(0)
        status_text = st.empty()

        for page_num in range(total_pages):
            if (page_num + 1) not in selected_pages:
                continue

            status_text.text(f"Processing page {page_num + 1}...")
            try:
                page = pdf.pages[page_num]
                table = page.extract_table()
                if not table:
                    tables_on_page = page.extract_tables({
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "text"
                    })
                    if tables_on_page:
                        table = tables_on_page[0]

                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df.columns = deduplicate_columns(df.columns.tolist())
                    tables.append((page_num + 1, df))
            except Exception as e:
                st.warning(f"Could not extract table from page {page_num + 1}: {str(e)}")
            progress_bar.progress((page_num + 1) / total_pages)
        status_text.text("Processing complete.")
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
if uploaded_file is not None:
    with pdfplumber.open(uploaded_file) as pdf:
        total_pages = len(pdf.pages)
    
    st.info(f"📘 This PDF contains **{total_pages} pages**. Select pages below to extract tables from:")

    # Page selection
    page_options = list(range(1, total_pages + 1))
    selected_pages = st.multiselect("Select Pages", options=page_options, default=page_options)

    if st.button("Extract Tables"):
        tables = extract_tables_from_pdf(uploaded_file, selected_pages)

        if tables:
            st.success(f"✅ Successfully extracted tables from **{len(tables)} pages**.")

            # Summary stats
            total_rows = sum(len(df) for _, df in tables)
            total_cols = max(len(df.columns) for _, df in tables) if tables else 0
            st.markdown(f"**Summary:**")
            st.markdown(f"- Total Tables Extracted: `{len(tables)}`")
            st.markdown(f"- Total Rows: `{total_rows}`")
            st.markdown(f"- Max Columns in Any Table: `{total_cols}`")

            # Preview tables
            for page_num, df in tables:
                with st.expander(f"📊 Table from Page {page_num}"):
                    st.dataframe(df, use_container_width=True)

            # Convert to Excel
            excel_data = convert_to_excel(tables, combine_tables)

            # Download button
            st.download_button(
                label="📥 Download Excel File",
                data=excel_data,
                file_name="extracted_tables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ No tables found on the selected pages.")

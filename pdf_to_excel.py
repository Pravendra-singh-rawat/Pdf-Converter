import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF Fixer", layout="wide")

st.title("🔧 PDF to Excel (Debug Version)")

uploaded_file = st.file_uploader("Upload PDF", type="pdf")

if uploaded_file:
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            all_tables = []
            
            # Loop through pages
            for i, page in enumerate(pdf.pages):
                # Attempt to extract table
                table = page.extract_table()
                
                if table:
                    df = pd.DataFrame(table)
                    # Use first row as header
                    df.columns = df.iloc[0]
                    df = df[1:]
                    all_tables.append(df)
                else:
                    st.warning(f"Page {i+1}: No table structure detected.")

            if all_tables:
                final_df = pd.concat(all_tables, ignore_index=True)
                st.success("Data extracted!")
                st.dataframe(final_df)

                # Export
                output = io.BytesIO()
                final_df.to_excel(output, index=False)
                st.download_button("Download Excel", data=output.getvalue(), file_name="output.xlsx")
            else:
                st.error("Could not find any tables. Is this a scanned image?")
                
    except Exception as e:
        st.error(f"Logic Error: {e}")

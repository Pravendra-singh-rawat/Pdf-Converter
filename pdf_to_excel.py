import streamlit as st
import pandas as pd
import tabula
import io

# Page Configuration
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📊", layout="centered")

# Custom CSS for a polished look
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #007bff;
        color: white;
    }
    .stDownloadButton>button {
        width: 100%;
        border-radius: 5px;
        background-color: #28a745;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

## Header Section
st.title("📊 PDF to Excel Converter")
st.subheader("Extract tables from your PDF files with ease")
st.write("Upload a PDF file below, and we'll attempt to find and convert the tables into a downloadable Excel sheet.")

---

## File Upload Section
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    with st.spinner('Analyzing PDF...'):
        try:
            # Read PDF tables
            # multiple_tables=True extracts all tables; pages='all' scans the whole doc
            tables = tabula.read_pdf(uploaded_file, pages='all', multiple_tables=True)
            
            if len(tables) > 0:
                st.success(f"Found {len(tables)} table(s)!")
                
                # Combine tables or let user pick? Here we combine them for simplicity
                all_tables = pd.concat(tables)
                
                # Data Preview
                st.write("### Preview of Extracted Data")
                st.dataframe(all_tables.head(10), use_container_width=True)

                # Conversion to Excel Buffer
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    all_tables.to_excel(writer, index=False, sheet_name='Sheet1')
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=buffer.getvalue(),
                    file_name="converted_data.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.warning("No clear tables were found in this PDF. Try a document with a more defined grid structure.")
                
        except Exception as e:
            st.error(f"An error occurred: {e}")

else:
    st.info("Please upload a PDF file to begin.")

---
st.caption("Built with ❤️ using Streamlit and Tabula")

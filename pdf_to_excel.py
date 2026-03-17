import streamlit as st
import pdfplumber
import pandas as pd
import io

# -----------------------------------------------------------------------------
# Page Configuration & Custom CSS
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Injecting custom CSS for a modern look
st.markdown("""
<style>
    /* Main Background */
    .stApp {
        background-color: #f8f9fa;
    }
    
    /* Header Styling */
    h1 {
        color: #2c3e50;
        font-family: 'Helvetica Neue', sans-serif;
        text-align: center;
        padding-bottom: 20px;
    }
    
    /* Card Container */
    .upload-container {
        background-color: white;
        padding: 30px;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        margin-top: 20px;
    }
    
    /* Button Styling */
    .stButton > button {
        width: 100%;
        background-color: #3498db;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        height: 50px;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #2980b9;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    /* Success Message */
    .success-box {
        padding: 15px;
        background-color: #d4edda;
        color: #155724;
        border-radius: 8px;
        border: 1px solid #c3e6cb;
        margin-top: 20px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------

def extract_tables_from_pdf(pdf_file):
    """
    Extracts tables from a PDF file using pdfplumber.
    Returns a list of DataFrames (one per page/table found).
    """
    all_dataframes = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages):
            # Extract tables from the current page
            tables = page.extract_tables()
            
            if tables:
                for table in tables:
                    # Convert table list to DataFrame
                    # Usually the first row is headers, but we let pandas infer or handle it
                    df = pd.DataFrame(table[1:], columns=table[0])
                    
                    # Clean up column names (remove None/NaN)
                    df.columns = [str(col).strip() if col else f"Col_{j}" for j, col in enumerate(df.columns)]
                    
                    # Add a marker for which page it came from (optional but helpful)
                    df['Source_Page'] = i + 1
                    
                    all_dataframes.append(df)
    
    return all_dataframes

def convert_df_to_excel(dfs):
    """
    Combines multiple DataFrames into one Excel file with multiple sheets 
    or one single sheet depending on preference. 
    Here we combine them into one master sheet for simplicity, 
    or you can iterate to create multiple sheets.
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Option A: Save each table as a separate sheet named 'Table 1', 'Table 2'...
        for i, df in enumerate(dfs):
            sheet_name = f"Table_{i+1}"
            # Excel sheet names max 31 chars
            if len(sheet_name) > 31: sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    output.seek(0)
    return output

# -----------------------------------------------------------------------------
# Main App Layout
# -----------------------------------------------------------------------------

def main():
    # Header Section
    st.title("📄 PDF to Excel Converter")
    st.markdown("<p style='text-align: center; color: #7f8c8d;'>Convert your tabular PDF data into editable Excel spreadsheets instantly.</p>", unsafe_allow_html=True)

    # File Uploader inside a styled container simulation
    uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

    if uploaded_file is not None:
        # Display file info
        st.success(f"✅ File uploaded: **{uploaded_file.name}** ({round(uploaded_file.size / 1024, 2)} KB)")
        
        # Processing Button
        if st.button("🚀 Convert to Excel"):
            with st.spinner("⏳ Extracting tables... This may take a moment."):
                try:
                    # 1. Extract Data
                    dfs = extract_tables_from_pdf(uploaded_file)
                    
                    if not dfs:
                        st.error("❌ No tables detected in this PDF. The file might be image-based or unstructured.")
                        st.info("💡 Tip: This tool works best with PDFs containing selectable text tables.")
                    else:
                        # 2. Show Preview
                        st.subheader("👀 Data Preview")
                        st.caption("Showing the first table found. Download the full file for all tables.")
                        
                        # Show the first dataframe in an expandable section if there are many
                        with st.expander("View First Table Dataframe"):
                            st.dataframe(dfs[0], use_container_width=True)
                        
                        st.markdown(f"**Total Tables Found:** {len(dfs)}")

                        # 3. Generate Excel
                        excel_buffer = convert_df_to_excel(dfs)
                        
                        # 4. Download Button
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_buffer,
                            file_name=f"{uploaded_file.name.split('.')[0]}_converted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        st.balloons()

                except Exception as e:
                    st.error(f"An error occurred during conversion: {e}")
    
    else:
        # Empty state illustration
        st.markdown("""
        <div class="upload-container">
            <h3>How it works:</h3>
            <ol style="text-align: left; display: inline-block;">
                <li>Upload a PDF containing tables.</li>
                <li>Click <b>Convert to Excel</b>.</li>
                <li>Preview the data and download the .xlsx file.</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

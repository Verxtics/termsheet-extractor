import streamlit as st
import tempfile
import os
import pandas as pd

st.set_page_config(
    page_title="Termsheet Data Extractor",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Structured Product Termsheet Extractor")
st.write("Upload PDF termsheets to automatically extract key trading data")

# File uploader
uploaded_file = st.file_uploader(
    "Drop your PDF termsheet here",
    type=['pdf'],
    help="Upload a PDF termsheet to extract structured product data"
)

if uploaded_file is not None:
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_path = tmp_file.name
    
    if st.button("Extract Data", type="primary"):
        with st.spinner("Processing termsheet..."):
            try:
                # Extract data using your fixed script
                data = extract_termsheet_data(tmp_path)
                
                if 'error' not in data:
                    st.success(f"Successfully extracted {data.get('Fields_Extracted', 0)} fields!")
                    
                    # Display results in columns
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.subheader("Core Information")
                        st.write(f"**Issuer:** {data.get('Issuer', 'N/A')}")
                        st.write(f"**ISIN:** {data.get('ISIN', 'N/A')}")
                        st.write(f"**Currency:** {data.get('Currency', 'N/A')}")
                        st.write(f"**Notional:** {data.get('Notional_Amount', 'N/A')}")
                    
                    with col2:
                        st.subheader("Key Dates")
                        st.write(f"**Issue Date:** {data.get('Issue_Date', 'N/A')}")
                        st.write(f"**Strike Date:** {data.get('Strike_Date', 'N/A')}")
                        st.write(f"**Maturity:** {data.get('Maturity_Date', 'N/A')}")
                    
                    with col3:
                        st.subheader("Risk Parameters")
                        st.write(f"**Knock-In Barrier:** {data.get('Knock_In_Barrier', 'N/A')}")
                        st.write(f"**Coupon Rate:** {data.get('Coupon_Rate', 'N/A')}")
                        st.write(f"**Final Coupon:** {data.get('Final_Coupon', 'N/A')}")
                    
                    # Create and offer Excel download
                    excel_path = "extracted_data.xlsx"
                    create_professional_excel_output(data, excel_path)
                    
                    with open(excel_path, "rb") as file:
                        st.download_button(
                            label="📥 Download Excel Report",
                            data=file.read(),
                            file_name=f"{uploaded_file.name.replace('.pdf', '')}_extracted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    # Show detailed tables
                    if data.get('observation_schedule'):
                        st.subheader("Observation Schedule")
                        df_schedule = pd.DataFrame(data['observation_schedule'])
                        st.dataframe(df_schedule, use_container_width=True)
                    
                    if data.get('underlying_assets'):
                        st.subheader("Underlying Assets")
                        df_underlying = pd.DataFrame(data['underlying_assets'])
                        st.dataframe(df_underlying, use_container_width=True)
                
                else:
                    st.error(f"Extraction failed: {data['error']}")
                    
            except Exception as e:
                st.error(f"Processing error: {str(e)}")
        
        # Clean up temp file
        os.unlink(tmp_path)

st.markdown("---")
st.caption("Professional termsheet extraction for structured products")


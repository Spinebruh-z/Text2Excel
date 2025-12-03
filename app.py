"""
Text to Excel Converter - Streamlit Application
Converts unstructured text data into structured Excel spreadsheets
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Import custom modules
from utils.text_extractor import extract_text_from_file
from utils.parser import parse_key_value_pairs
from utils.excel_generator import generate_excel

# Page configuration
st.set_page_config(
    page_title="Text to Excel Converter",
    page_icon="📊",
    layout="wide"
)

def main():
    st.title("📊 Text to Excel Converter")
    st.markdown("---")
    st.markdown("Convert unstructured text data into structured Excel spreadsheets")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        st.subheader("🔑 Google Gemini API Key")
        api_key = st.text_input(
            "Enter your Gemini API key",
            type="password",
            help="Get your API key from https://aistudio.google.com/app/apikey"
        )
        
        st.subheader("Key Patterns (Optional)")
        custom_keys = st.text_area(
            "Enter custom keys to extract (one per line)",
            placeholder="Name\nEmail\nPhone\nAddress",
            height=150
        )
        
        st.subheader("Excel Options")
        include_header = st.checkbox("Include Headers", value=True)
        auto_width = st.checkbox("Auto-adjust Column Width", value=True)
        
    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📁 Upload Files")
        uploaded_files = st.file_uploader(
            "Choose text files",
            type=['txt', 'docx', 'pdf'],
            accept_multiple_files=True,
            help="Upload .txt, .docx, or .pdf files"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully")
            
            # Display uploaded files
            with st.expander("View uploaded files"):
                for file in uploaded_files:
                    st.write(f"📄 {file.name} ({file.size} bytes)")
    
    with col2:
        st.subheader("⚡ Process Files")
        
        if uploaded_files:
            if not api_key:
                st.warning("⚠️ Please enter your Gemini API key in the sidebar")
            elif st.button("🚀 Extract Data", type="primary", use_container_width=True):
                with st.spinner("Processing files with Gemini LLM..."):
                    try:
                        # Process files
                        all_data = []
                        
                        for file_idx, file in enumerate(uploaded_files, start=1):
                            # Generate doc_id (D1, D2, D3, etc.)
                            doc_id = f"D{file_idx}"
                            
                            # Extract text
                            text = extract_text_from_file(file)
                            
                            # Parse key-value pairs using Gemini LLM with doc_id
                            keys_list = [k.strip() for k in custom_keys.split('\n') if k.strip()] if custom_keys else None
                            rows = parse_key_value_pairs(text, api_key, keys_list, "Gemini LLM", doc_id)
                            
                            # Add source file to each row
                            for row in rows:
                                row['source_file'] = file.name
                                all_data.append(row)
                        
                        # Create DataFrame with required columns: key, value, comments
                        df = pd.DataFrame(all_data)
                        
                        # Reorder columns to ensure key, value, comments come first
                        cols = ['key', 'value', 'comments', 'source_file']
                        df = df[[c for c in cols if c in df.columns]]
                        
                        # Store in session state
                        st.session_state['extracted_data'] = df
                        st.session_state['processed'] = True
                        
                        st.success(f"✅ Data extracted successfully! Found {len(df)} records.")
                        
                    except Exception as e:
                        st.error(f"❌ Error processing files: {str(e)}")
                        st.exception(e)
    
    # Display results
    if 'processed' in st.session_state and st.session_state['processed']:
        st.markdown("---")
        st.subheader("📊 Extracted Data Preview")
        
        df = st.session_state['extracted_data']
        
        # Data editor for review and editing
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            height=400
        )
        
        # Update session state with edited data
        st.session_state['extracted_data'] = edited_df
        
        # Download section
        st.markdown("---")
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col1:
            st.subheader("💾 Export Data")
        
        with col2:
            # Export to Excel
            excel_buffer = generate_excel(edited_df, include_header, auto_width)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            st.download_button(
                label="📥 Download Excel",
                data=excel_buffer,
                file_name=f"Output_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        
        with col3:
            # Export to CSV
            csv = edited_df.to_csv(index=False)
            st.download_button(
                label="📥 Download CSV",
                data=csv,
                file_name=f"Output_{timestamp}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        # Statistics
        with st.expander("📈 Data Statistics"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", len(edited_df))
            with col2:
                st.metric("Total Columns", len(edited_df.columns))
            with col3:
                st.metric("Files Processed", edited_df['source_file'].nunique() if 'source_file' in edited_df.columns else 0)

if __name__ == "__main__":
    main()

"""
Excel File Merger - Streamlit Web Application
A simple web interface for merging multiple supplier Excel files into one master file.
"""

import streamlit as st
import pandas as pd
from excel_merger import ExcelMerger
import datetime

# Page configuration
st.set_page_config(
    page_title="Excel File Merger",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
        border-bottom: 2px solid #1f77b4;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        color: #2e8b57;
        margin: 1rem 0;
        padding: 0.5rem 0;
        border-left: 4px solid #2e8b57;
        padding-left: 1rem;
    }
    .info-box {
        background-color: #f0f8ff;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #f0fff0;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #32cd32;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #fff0f0;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ff6b6b;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main application function."""
    
    # Header
    st.markdown('<h1 class="main-header">📊 Excel File Merger</h1>', unsafe_allow_html=True)
    
    # Initialize session state
    if 'merger' not in st.session_state:
        st.session_state.merger = ExcelMerger()
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    
    # Sidebar with instructions
    with st.sidebar:
        st.markdown('<h2 class="section-header">📋 Instructions</h2>', unsafe_allow_html=True)
        st.markdown("""
        1. **Upload Files**: Select multiple Excel files (.xlsx or .xls) to merge
        2. **Review**: Check the preview and file information
        3. **Download**: Get your merged Excel file
        
        **Features:**
        - Automatically handles different column structures
        - Removes duplicate rows
        - Adds source file tracking
        - Provides data summary
        """)
        
        st.markdown('<h2 class="section-header">ℹ️ File Requirements</h2>', unsafe_allow_html=True)
        st.markdown("""
        - **Format**: Excel files (.xlsx, .xls)
        - **Size**: No specific limit
        - **Structure**: Any column structure (will be standardized)
        - **Content**: Tabular data preferred
        """)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown('<h2 class="section-header">📁 Upload Excel Files</h2>', unsafe_allow_html=True)
        
        # File uploader
        uploaded_files = st.file_uploader(
            "Choose Excel files to merge",
            accept_multiple_files=True,
            type=['xlsx', 'xls'],
            help="Select multiple Excel files that you want to merge into one master file."
        )
        
        # Process files when uploaded
        if uploaded_files and not st.session_state.processed:
            with st.spinner("Processing uploaded files..."):
                success, message = st.session_state.merger.process_uploaded_files(uploaded_files)
                
                if success:
                    st.session_state.processed = True
                    st.markdown(f'<div class="success-box">✅ {message}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="error-box">❌ {message}</div>', unsafe_allow_html=True)
        
        # Reset when files change
        if not uploaded_files and st.session_state.processed:
            st.session_state.processed = False
            st.session_state.merger = ExcelMerger()
    
    with col2:
        if uploaded_files:
            st.markdown('<h3 class="section-header">📊 Upload Status</h3>', unsafe_allow_html=True)
            st.write(f"**Files uploaded:** {len(uploaded_files)}")
            
            for file in uploaded_files:
                st.write(f"• {file.name} ({file.size:,} bytes)")
    
    # Show results if files are processed
    if st.session_state.processed:
        
        # Display summary information
        summary = st.session_state.merger.get_merged_data_summary()
        
        if summary:
            st.markdown('<h2 class="section-header">📈 Merge Summary</h2>', unsafe_allow_html=True)
            
            # Summary metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Rows", summary['total_rows'])
            with col2:
                st.metric("Total Columns", summary['total_columns'])
            with col3:
                st.metric("Source Files", len(summary['file_info']))
            with col4:
                st.metric("Unique Columns", len(summary['columns']))
            
            # File information table
            st.markdown('<h3 class="section-header">📋 File Details</h3>', unsafe_allow_html=True)
            file_df = pd.DataFrame(summary['file_info'])
            st.dataframe(file_df, use_container_width=True)
            
            # Column information
            st.markdown('<h3 class="section-header">📝 Column Information</h3>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Columns in merged data:**")
                for i, col in enumerate(summary['columns'], 1):
                    if col != 'Source_File':  # Hide the added source column from user view
                        st.write(f"{i}. {col}")
            
            with col2:
                st.write("**Data types:**")
                for col, dtype in summary['data_types'].items():
                    if col != 'Source_File':
                        st.write(f"• {col}: {dtype}")
            
            # Data preview
            st.markdown('<h3 class="section-header">👀 Data Preview</h3>', unsafe_allow_html=True)
            preview_data = st.session_state.merger.get_preview_data()
            if preview_data is not None:
                st.dataframe(preview_data, use_container_width=True)
            
            # Download section
            st.markdown('<h2 class="section-header">💾 Download Merged File</h2>', unsafe_allow_html=True)
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                if st.button("🔄 Generate Excel File", type="primary"):
                    with st.spinner("Generating Excel file..."):
                        excel_data = st.session_state.merger.export_to_excel()
                        if excel_data:
                            st.session_state.excel_data = excel_data
                            st.success("Excel file generated successfully!")
                        else:
                            st.error("Failed to generate Excel file.")
            
            with col2:
                if hasattr(st.session_state, 'excel_data'):
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"merged_excel_file_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="📥 Download Merged Excel File",
                        data=st.session_state.excel_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Click to download the merged Excel file to your computer."
                    )
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; font-size: 0.9rem;">
        Excel File Merger | Built with Streamlit 📊
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
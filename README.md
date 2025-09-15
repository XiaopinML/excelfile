# Excel File Merger 📊

A Python web application built with Streamlit that merges multiple supplier Excel files into one clean master Excel file.

## Features

- **Multi-file Upload**: Support for multiple Excel files (.xlsx, .xls)
- **Intelligent Merging**: Automatically handles different column structures
- **Data Cleaning**: Removes duplicate rows and empty data
- **Source Tracking**: Adds source file information to merged data
- **Data Summary**: Provides detailed analysis of merged data
- **Professional Export**: Generates clean Excel files with formatting
- **User-friendly Interface**: Simple drag-and-drop file upload

## Quick Start

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
streamlit run app.py
```

3. Open your browser to `http://localhost:8501`

4. Upload multiple Excel files and download the merged result!

## How It Works

1. **Upload**: Select multiple Excel files from different suppliers
2. **Process**: The app automatically standardizes column structures and merges data
3. **Review**: Check the merge summary, file details, and data preview
4. **Download**: Generate and download the clean master Excel file

## Technical Details

- **Framework**: Streamlit (lightweight Python web framework)
- **Excel Processing**: pandas + openpyxl for robust Excel file handling
- **File Support**: .xlsx and .xls formats
- **Data Handling**: Intelligent column alignment and duplicate removal
- **Export**: Professional Excel formatting with multiple sheets

## File Structure

```
excelfile/
├── app.py              # Main Streamlit application
├── excel_merger.py     # Core Excel merging functionality
├── requirements.txt    # Python dependencies
├── .gitignore         # Git ignore file
└── README.md          # This file
```

## Example Use Cases

- Merging product catalogs from multiple suppliers
- Consolidating inventory data from different sources
- Combining price lists with different column structures
- Creating master databases from supplier submissions

## Requirements

- Python 3.7+
- Streamlit 1.28.0+
- pandas 2.0.0+
- openpyxl 3.1.0+
- xlsxwriter 3.1.0+ 

"""
Excel File Merger Utility
Provides functionality to merge multiple Excel files into a single master file.
"""

import pandas as pd
import io
from typing import List, Dict, Any, Optional, Tuple
import tempfile
import os


class ExcelMerger:
    """Handles merging of multiple Excel files into a single master file."""
    
    def __init__(self):
        self.merged_data = None
        self.file_info = []
    
    def process_uploaded_files(self, uploaded_files: List[Any]) -> Tuple[bool, str]:
        """
        Process uploaded Excel files and merge them.
        
        Args:
            uploaded_files: List of uploaded file objects
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            if not uploaded_files:
                return False, "No files uploaded"
            
            all_dataframes = []
            self.file_info = []
            
            for uploaded_file in uploaded_files:
                # Read Excel file
                try:
                    # Try to read the file - support both .xlsx and .xls
                    if uploaded_file.name.endswith(('.xlsx', '.xls')):
                        df = pd.read_excel(uploaded_file)
                        
                        # Add source file column
                        df['Source_File'] = uploaded_file.name
                        
                        all_dataframes.append(df)
                        self.file_info.append({
                            'name': uploaded_file.name,
                            'rows': len(df),
                            'columns': len(df.columns)
                        })
                    else:
                        return False, f"Unsupported file type: {uploaded_file.name}. Please upload .xlsx or .xls files."
                        
                except Exception as e:
                    return False, f"Error reading {uploaded_file.name}: {str(e)}"
            
            if not all_dataframes:
                return False, "No valid Excel files found"
            
            # Merge all dataframes
            self.merged_data = self._merge_dataframes(all_dataframes)
            
            return True, f"Successfully merged {len(uploaded_files)} files with {len(self.merged_data)} total rows"
            
        except Exception as e:
            return False, f"Error processing files: {str(e)}"
    
    def _merge_dataframes(self, dataframes: List[pd.DataFrame]) -> pd.DataFrame:
        """
        Merge multiple dataframes intelligently.
        
        Args:
            dataframes: List of pandas DataFrames to merge
            
        Returns:
            Merged DataFrame
        """
        if not dataframes:
            return pd.DataFrame()
        
        if len(dataframes) == 1:
            return dataframes[0]
        
        # Get all unique columns across all dataframes
        all_columns = set()
        for df in dataframes:
            all_columns.update(df.columns)
        
        # Ensure all dataframes have the same columns (fill missing with NaN)
        standardized_dfs = []
        for df in dataframes:
            # Create a copy to avoid modifying original
            df_copy = df.copy()
            
            # Add missing columns with NaN values
            for col in all_columns:
                if col not in df_copy.columns:
                    df_copy[col] = None
            
            # Reorder columns to match
            df_copy = df_copy[sorted(all_columns)]
            standardized_dfs.append(df_copy)
        
        # Concatenate all dataframes
        merged_df = pd.concat(standardized_dfs, ignore_index=True)
        
        # Clean the data
        merged_df = self._clean_data(merged_df)
        
        return merged_df
    
    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean the merged data.
        
        Args:
            df: DataFrame to clean
            
        Returns:
            Cleaned DataFrame
        """
        # Remove completely empty rows
        df = df.dropna(how='all')
        
        # Remove duplicate rows (keep first occurrence)
        df = df.drop_duplicates()
        
        # Reset index
        df = df.reset_index(drop=True)
        
        return df
    
    def get_merged_data_summary(self) -> Dict[str, Any]:
        """
        Get summary information about the merged data.
        
        Returns:
            Dictionary with summary information
        """
        if self.merged_data is None:
            return {}
        
        return {
            'total_rows': len(self.merged_data),
            'total_columns': len(self.merged_data.columns),
            'columns': list(self.merged_data.columns),
            'file_info': self.file_info,
            'data_types': self.merged_data.dtypes.to_dict()
        }
    
    def export_to_excel(self) -> Optional[bytes]:
        """
        Export merged data to Excel format.
        
        Returns:
            Excel file content as bytes, or None if no data
        """
        if self.merged_data is None:
            return None
        
        # Create Excel file in memory
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write merged data to main sheet
            self.merged_data.to_excel(writer, sheet_name='Merged_Data', index=False)
            
            # Write summary information to a separate sheet
            summary_data = []
            for info in self.file_info:
                summary_data.append([info['name'], info['rows'], info['columns']])
            
            summary_df = pd.DataFrame(summary_data, columns=['File Name', 'Rows', 'Columns'])
            summary_df.to_excel(writer, sheet_name='File_Summary', index=False)
            
            # Get workbook and worksheets for formatting
            workbook = writer.book
            merged_sheet = writer.sheets['Merged_Data']
            summary_sheet = writer.sheets['File_Summary']
            
            # Add some basic formatting
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Format headers
            for col_num, value in enumerate(self.merged_data.columns.values):
                merged_sheet.write(0, col_num, value, header_format)
            
            for col_num, value in enumerate(summary_df.columns.values):
                summary_sheet.write(0, col_num, value, header_format)
        
        output.seek(0)
        return output.getvalue()
    
    def get_preview_data(self, num_rows: int = 5) -> Optional[pd.DataFrame]:
        """
        Get preview of merged data.
        
        Args:
            num_rows: Number of rows to preview
            
        Returns:
            DataFrame with preview data, or None if no data
        """
        if self.merged_data is None:
            return None
        
        return self.merged_data.head(num_rows)
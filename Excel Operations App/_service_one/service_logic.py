# ==========================================
# 1. IMPORTS
# ==========================================

# Third-Party Libraries
import os
import logging
import pandas as pd
import numpy as np

from _config.settings import load_config
from _utils.functions import auto_header_finder, clean_sheet_name, open_file

# Initialize Logger
logger = logging.getLogger('AppLogger')

# ==========================================
# 2. ANALYSIS LOGIC
# ==========================================

def analyse() -> None:
    """
    Performs the Forecast Comparison logic:
    1. Loads two Excel files.
    2. cleans and merges them based on Unique ID.
    3. Pivots data to compare Forecast columns.
    4. Exports results to Excel.
    """
    
    print("\n\nLoading data...")
    logger.info("Starting Forecast Analysis")

    config = load_config()
    file1_data = config["service_one"]["file1"]
    file2_data = config["service_one"]["file2"]

    try:
        # ==========================================
        # LOAD DATA
        # ==========================================
        
        #### Load 1st Excel File ####
        print("Reading File 1...")
        first_header_row = auto_header_finder(file1_data)
        excel1 = pd.read_excel(
            file1_data["file_path"], 
            sheet_name=file1_data["sheet_name"], 
            header=first_header_row
        )

        #### Load 2nd Excel File ####
        print("Reading File 2...")
        second_header_row = auto_header_finder(file2_data)
        excel2 = pd.read_excel(
            file2_data["file_path"], 
            sheet_name=file2_data["sheet_name"], 
            header=second_header_row
        )

        # ==========================================
        # DATA PREPARATION
        # ==========================================
        
        # Check if 'Subregion' exists in both files (Dynamic Handling)
        has_subregion = "Subregion" in excel1.columns and "Subregion" in excel2.columns
        if not has_subregion:
            logger.info("'Subregion' column missing in one or both files. Proceeding without subregion grouping.")

        # Define columns to pull
        id_col1 = file1_data["unique_id_column"]
        fc_col1 = file1_data["forecast_column"]
        extra_cols = file1_data["extra_columns"]
        
        id_col2 = file2_data["unique_id_column"]
        fc_col2 = file2_data["forecast_column"]

        # Build column list for File 1
        cols_pull1 = [id_col1, fc_col1] + extra_cols
        if has_subregion:
            cols_pull1.append("Subregion")
            
        # Build column list for File 2
        cols_pull2 = [id_col2, fc_col2]
        if has_subregion:
            cols_pull2.append("Subregion")

        # Trim DataFrames
        excel1_trim = excel1[cols_pull1].copy()
        excel2_trim = excel2[cols_pull2].copy()

        print("Processing and cleaning data...")

        # Normalize ID Columns (Strip whitespace & Lowercase) to ensure accurate merge
        excel1_trim[id_col1] = excel1_trim[id_col1].astype(str).str.strip().str.lower()
        excel2_trim[id_col2] = excel2_trim[id_col2].astype(str).str.strip().str.lower()

        # Rename File 2 columns to match File 1 for merging
        rename_map = {
            fc_col2: fc_col1,
            id_col2: id_col1
        }
        excel2_trim = excel2_trim.rename(columns=rename_map)

        # Merge Extra Columns from File 1 into File 2 (so File 2 has context)
        merge_cols = [id_col1] + extra_cols
        if has_subregion:
            merge_cols.append("Subregion")
            
        excel2_trim = excel2_trim.merge(
            excel1_trim[merge_cols], 
            on=id_col1, 
            how='left',
            suffixes=('', '_duplicate') # Handle potential collisions
        )
        
        # Ensure Excel 2 has the same column structure as Excel 1
        # (Intersect columns to avoid key errors if merge missed something)
        common_cols = [c for c in excel1_trim.columns if c in excel2_trim.columns]
        excel2_trim = excel2_trim[common_cols]

        # Convert Forecasts to Numeric (Coerce errors to NaN)
        excel1_trim[fc_col1] = pd.to_numeric(excel1_trim[fc_col1], errors='coerce').fillna(0)
        excel2_trim[fc_col1] = pd.to_numeric(excel2_trim[fc_col1], errors='coerce').fillna(0)

        # Add Source Identifiers
        excel1_trim['Source'] = 'source1'
        excel2_trim['Source'] = 'source2'

        # ==========================================
        # MERGE & PIVOT
        # ==========================================
        
        combined_df = pd.concat([excel1_trim, excel2_trim], ignore_index=True)

        # Define Index for Pivot
        pivot_index = [id_col1] + extra_cols
        if has_subregion:
            pivot_index.append("Subregion")

        pivot_table = combined_df.pivot_table(
            index=pivot_index,
            columns='Source',
            values=fc_col1,
            aggfunc='first' # Use first if duplicates exist
        ).fillna(0) # Fill NaNs with 0 for calculation

        # ==========================================
        # CALCULATIONS
        # ==========================================
        
        pivot_table['Difference'] = pivot_table['source1'] - pivot_table['source2']

        # Calculate Similarity %
        # Logic: (1 - |diff| / source1) * 100
        pivot_table['Similarity %'] = pivot_table.apply(
            lambda row: (1 - abs(row['source1'] - row['source2']) / row['source1']) * 100
            if row['source1'] != 0 else 0,
            axis=1
        )

        # Sort: Highest difference first
        pivot_table = pivot_table.sort_values(
            by=['Difference', 'Similarity %'],
            ascending=[False, True]
        )

        # ==========================================
        # EXPORT
        # ==========================================
        
        # Prepare output path
        file_name1 = os.path.splitext(os.path.basename(file1_data["file_path"]))[0]
        file_name2 = os.path.splitext(os.path.basename(file2_data["file_path"]))[0]
        
        output_dir = "outputs/forecast_comparisons"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        output_file_name = f"{output_dir}/comparison_{file_name1}_vs_{file_name2}.xlsx"

        print(f"Exporting to {output_file_name}...")

        with pd.ExcelWriter(output_file_name, engine="openpyxl") as writer:
            # 1. Info Sheet
            info_data = {
                "Metric": ["Source 1 File", "Source 2 File", "ID Column", "Forecast Column"],
                "Value": [file_name1, file_name2, id_col1, fc_col1]
            }
            pd.DataFrame(info_data).to_excel(writer, sheet_name="Info", index=False)
            
            # 2. All Data Sheet
            pivot_table.to_excel(writer, sheet_name="All Data", index=True)
            
            # 3. Subregion Sheets (Only if Subregion exists)
            if has_subregion:
                # Check if Subregion is in the index levels
                if "Subregion" in pivot_table.index.names:
                    unique_regions = pivot_table.index.get_level_values("Subregion").unique()
                    
                    for subregion in unique_regions:
                        clean_sub = clean_sheet_name(str(subregion))
                        try:
                            sub_data = pivot_table.xs(subregion, level="Subregion")
                            sub_data.to_excel(writer, sheet_name=clean_sub, index=True)
                        except Exception as e:
                            logger.warning(f"Could not write sheet for subregion {subregion}: {e}")

        print("Analysis Complete.")
        open_file(output_file_name)

    except Exception as e:
        logger.error(f"Analysis Failed: {e}")
        print(f"An error occurred during analysis: {e}")
# ==========================================
# 1. IMPORTS
# ==========================================

# Third-Party Libraries
import os
import pandas as pd
import logging
from typing import Optional

from _config.settings import load_config
from _utils.functions import open_file

# Initialize Logger
logger = logging.getLogger('AppLogger')

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================

def get_ordinal(n: int) -> str:
    """Returns the ordinal suffix for a number (e.g., 1st, 2nd, 3rd)."""
    if 10 <= n <= 20:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")


def get_ordinal_word(n: int) -> str:
    """Returns the ordinal word (e.g., first, second, third)."""
    # Simple mapping for common cases, falling back to numeric ordinal
    ordinals = {1: "first", 2: "second", 3: "third", 4: "fourth", 5: "fifth"}
    return ordinals.get(n, f"{n}{get_ordinal(n)}")

# ==========================================
# 3. CORE LOGIC
# ==========================================

def service_merge() -> Optional[pd.DataFrame]:
    """
    Merges multiple Excel files based on selected columns.
    Returns the merged DataFrame or None if failed.
    """
    config = load_config()

    file_paths = config["service_eight"]["file_paths"]
    sheet_names = config["service_eight"]["sheet_names"]
    merging_columns = config["service_eight"]["merging_columns"]

    merged_df = None
    
    print("Starting merge process...")

    for i, file_path in enumerate(file_paths):
        sheet_name = sheet_names[i]
        merging_column = merging_columns[i]
        
        print(f"Processing File {i+1}: {os.path.basename(file_path)}...")

        try:
            # Load Data
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

            # Validate Column Exists
            if merging_column not in df.columns:
                msg = f"Warning: Column '{merging_column}' not found in file {i+1}. Skipping."
                print(msg)
                logger.warning(msg)
                continue

            # Normalize Merge Key to String to avoid type mismatches
            df[merging_column] = df[merging_column].astype(str).fillna('')

            # Perform Merge
            if merged_df is None:
                # First file becomes the base
                merged_df = df.rename(columns={merging_column: 'merge_key'})
            else:
                # Subsequent files are merged onto the base
                current_df = df.rename(columns={merging_column: 'merge_key'})
                
                # Merge outer to keep all records
                suffix = f"_{get_ordinal_word(i + 1)}"
                merged_df = pd.merge(
                    merged_df, 
                    current_df, 
                    on='merge_key', 
                    how='outer', 
                    suffixes=('', suffix)
                )

        except Exception as e:
            msg = f"Error processing file {i+1}: {e}"
            print(msg)
            logger.error(msg)
            return None

    if merged_df is not None:
        print("\nMerging completed successfully.")
        return merged_df
    else:
        print("Merge failed: No data processed.")
        return None


def save_to_excel(merged_df: pd.DataFrame) -> None:
    """
    Trims the DataFrame to selected columns and saves it.
    """
    if merged_df is None or merged_df.empty:
        print("Nothing to save.")
        return

    config = load_config()
    selected_columns = config["service_eight"]["selected_columns"]
    file_paths = config["service_eight"]["file_paths"]
    sheet_names = config["service_eight"]["sheet_names"]

    # 1. Generate Safe Filename
    # Join first 3 filenames max to avoid hitting OS path limits
    base_names = [os.path.basename(p).rsplit('.', 1)[0] for p in file_paths[:3]]
    file_name_str = "_".join(base_names)
    
    if len(file_paths) > 3:
        file_name_str += "_and_others"
        
    # Ensure directory exists
    output_dir = "outputs/merged_files"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    filepath = f"{output_dir}/merged_{file_name_str}.xlsx"

    try:
        # 2. Filter Columns
        # Ensure selected columns actually exist in the dataframe
        valid_columns = [col for col in selected_columns if col in merged_df.columns]
        
        if not valid_columns:
            print("Error: None of the selected columns exist in the merged data.")
            return
            
        trimmed_df = merged_df[valid_columns]

        # 3. Write to Excel
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            trimmed_df.to_excel(writer, sheet_name='Data', index=False)

            # Create Information Sheet for traceability
            info_data = {
                "Source File": [os.path.basename(f) for f in file_paths],
                "Sheet Name": sheet_names,
                "Order": [i + 1 for i in range(len(file_paths))]
            }
            pd.DataFrame(info_data).to_excel(writer, sheet_name='Information', index=False)

        print(f"File saved successfully to:\n{filepath}")
        
        # 4. Open File
        open_file(filepath)

    except Exception as e:
        logger.error(f"Error saving file: {e}")
        print(f"Error saving file: {e}")
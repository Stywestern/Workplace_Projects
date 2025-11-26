# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import os
import subprocess
import platform
import logging
from datetime import datetime
from typing import Optional, List, Union, Dict, Any

import pandas as pd
import customtkinter as ctk

# Initialize Logger
logger = logging.getLogger('AppLogger')

# ==========================================
# 2. GUI UTILITIES
# ==========================================

def clear_frame(frame: Union[ctk.CTkFrame, ctk.CTkScrollableFrame]) -> None:
    """Destroys all child widgets within a given frame."""
    for widget in frame.winfo_children():
        widget.destroy()


# ==========================================
# 3. STRING & DATE PARSING
# ==========================================

def extract_date(filename: str) -> Optional[datetime]:
    """
    Extracts a date from a filename.
    Expected format: prefix%YYYY-MM-DD-HHMMSS.xlsx or similar.
    """
    # Remove extension
    date_str = os.path.splitext(filename)[0]
    
    # Try parsing standard formats
    # You may need to adjust these formats based on your actual file naming convention
    formats = [
        '%Y%m%d_%H%M',      # e.g., 20241218_0903
        '%Y-%m-%d-%H%M%S',  # e.g., 2024-12-18-090340
        '%Y-%m-%d'          # e.g., 2024-12-18
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
            
    return None


def clean_sheet_name(string: str) -> str:
    """Sanitizes a string to make it a valid Excel sheet name."""
    invalid_chars = ['/', '\\', '?', '*', '[', ']', ':']
    for char in invalid_chars:
        string = string.replace(char, '_') 
    
    # Excel sheet names are limited to 31 chars
    return string[:31]


# ==========================================
# 4. DATAFRAME UTILITIES
# ==========================================

def auto_header_finder(file_data: Dict[str, Any], column_name: str = "ID") -> int:
    """
    Scans the first few rows of an Excel sheet to find the header row 
    by looking for a specific column name (e.g., 'ID', 'Unique Key').
    """
    try:
        # Read headerless to scan rows
        excel1 = pd.read_excel(
            file_data["file_path"], 
            sheet_name=file_data["sheet_name"], 
            header=None,
            nrows=20 # Limit scan to first 20 rows for efficiency
        )

        # Loop through the rows to find the anchor column
        for i, row in excel1.iterrows():
            # Convert row to string to ensure safe search
            row_values = [str(val).strip() for val in row.values]
            
            if column_name in row_values:
                logger.info(f"Header found at row {i}")
                return i 

        # Default to 0 if not found, rather than crashing
        logger.warning(f"Column '{column_name}' not found. Defaulting to row 0.")
        return 0

    except Exception as e:
        logger.error(f"Error finding header: {e}")
        return 0


def dolarize(series: pd.Series) -> List[str]:
    """
    Formats a numeric series into US Dollar strings (e.g., $1,234).
    Avoids using 'locale' library to prevent global environment side effects.
    """
    formatted_list = []
    
    for amount in series:
        try:
            val = float(amount)
            # Format with commas and no decimals
            formatted = "${:,.0f}".format(abs(val))
            
            if val < 0:
                formatted = f"-{formatted}"
                
            formatted_list.append(formatted)
        except (ValueError, TypeError):
            # If not a number, return as is
            formatted_list.append(str(amount))
            
    return formatted_list


# ==========================================
# 5. OS SYSTEM UTILITIES
# ==========================================

def open_file(output_path: str) -> None:
    """Opens a file using the operating system's default application."""
    try:
        abs_path = os.path.abspath(output_path)
        
        if not os.path.exists(abs_path):
            logger.error(f"File not found: {abs_path}")
            return

        current_os = platform.system()

        if current_os == 'Windows':
            os.startfile(abs_path)
        elif current_os == 'Darwin':  # macOS
            subprocess.run(['open', abs_path], check=True)
        elif current_os == 'Linux':
            subprocess.run(['xdg-open', abs_path], check=True)
        else:
            logger.warning("Cannot open file: OS not supported.")
            
    except Exception as e:
        logger.error(f"Error opening file: {e}")
        print(f"Error opening the output file: {e}")
        
        
        
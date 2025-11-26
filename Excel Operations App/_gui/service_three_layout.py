# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import os
import logging
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, ttk

# Local Imports
from _config.settings import load_config, save_config
from _utils.functions import clear_frame
from _gui.monitor import setup_monitor

# Service Specific Imports
from _service_three.service_logic import analyse
from _service_three.service_constants import entry_message

# Initialize Logger
logger = logging.getLogger('AppLogger')

# ==========================================
# 2. SELECTION LOGIC
# ==========================================

def select_file(button: ctk.CTkButton, dropdown: ctk.CTkComboBox) -> None:
    """
    Opens file dialog, loads Excel sheet names, and populates the dropdown.
    """
    config = load_config()
    file_path = filedialog.askopenfilename(title="Select Usage Report File")
    
    if file_path:
        try:
            # Load Excel metadata
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            sheet_names_list = excel_file.sheet_names
            
            # Update Dropdown
            dropdown.configure(values=sheet_names_list)
            dropdown.set("Select sheet")

            # Update Button Text
            file_name = os.path.basename(file_path)
            button.configure(text=f"Selected file: {file_name}")
            print(f"Usage report file set to: {file_name}")

            # Save to Config
            config["service_three"]["file_path"] = file_path
            save_config(config)

        except Exception as e:
            logger.error(f"Failed to load file {file_path}: {e}")
            print(f"Error loading file: {e}")


def handle_sheet_selection(selected_sheet: str) -> None:
    """Handles the selection of the sheet and saves it to config."""
    print(f"Selected sheet for usage analysis: {selected_sheet}\n")
    
    config = load_config()
    config["service_three"]["sheet_name"] = selected_sheet
    save_config(config)


# ==========================================
# 3. LAYOUT INITIALIZATION
# ==========================================

def setup_service_three(left_frame: ctk.CTkFrame, right_frame: ctk.CTkFrame) -> None:
    """Sets up the layout for the Usage Change Analysis service."""
    
    logger.info('Entered Service Three')
    
    # Clean slate
    clear_frame(left_frame)
    clear_frame(right_frame)
    
    # Delayed import
    from _gui.layout import setup_default_layout

    # ---------------------------
    # Navigation
    # ---------------------------
    switch_button = ctk.CTkButton(
        left_frame, 
        text="< Main Menu", 
        command=lambda: setup_default_layout(left_frame, right_frame)
    )
    switch_button.grid(row=0, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # ---------------------------
    # Monitor Setup
    # ---------------------------
    setup_monitor(right_frame)
    print(entry_message)
    
    # ---------------------------
    # UI Components
    # ---------------------------

    # File Selection
    file_button = ctk.CTkButton(
        left_frame, 
        text="1. Select Report File",
        command=lambda: select_file(file_button, sheet_dropdown)
    )
    file_button.grid(row=1, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # Sheet Selection
    sheet_dropdown = ctk.CTkComboBox(
        left_frame, 
        values=[], 
        state="readonly",
        command=handle_sheet_selection
    )
    sheet_dropdown.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    sheet_dropdown.set("Select Sheet")
    
    # Analysis Action
    analysis_button = ctk.CTkButton(
        left_frame, 
        text="2. Analyse Report", 
        fg_color="green",
        hover_color="darkgreen",
        command=analyse
    )
    analysis_button.grid(row=3, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # ---------------------------
    # Grid Config
    # ---------------------------
    left_frame.grid_columnconfigure(0, weight=1)
    
    # Reset row weights to ensure compact layout
    for i in range(5):
        left_frame.grid_rowconfigure(i, weight=0)
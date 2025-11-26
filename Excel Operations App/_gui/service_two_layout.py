# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import os
import logging
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog
from typing import Dict

# Local Imports
from _config.settings import load_config, save_config
from _utils.functions import clear_frame
from _gui.monitor import setup_monitor

# Service Specific Imports
from _service_two.service_logic import compare
from _service_two.service_constants import entry_message

# Initialize Logger
logger = logging.getLogger('AppLogger')

# Global storage for checkbox variables to ensure persistence between clicks
checkbox_vars: Dict[str, ctk.BooleanVar] = {}

# ==========================================
# 2. SELECTION LOGIC
# ==========================================

def select_file(button: ctk.CTkButton, dropdown: ctk.CTkComboBox, file_index: int) -> None:
    """Opens file dialog and populates the sheet dropdown."""
    config = load_config()
    file_path = filedialog.askopenfilename(title=f"Select Excel file {file_index}")
    
    if file_path:
        try:
            # Read the excel file and get sheet names
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            sheet_names_list = excel_file.sheet_names
            
            # Update the dropdown menu with sheet names
            dropdown.configure(values=sheet_names_list)
            dropdown.set("Select sheet")

            # Update Button Text safely
            file_name = os.path.basename(file_path)
            button.configure(text=f"Selected file: {file_name}")
            print(f"File {file_index} is set to: {file_name}")

            # Save to config
            config["service_two"][f"file{file_index}"]["file_path"] = file_path
            save_config(config)
            
        except Exception as e:
            logger.error(f"Failed to load file {file_path}: {e}")
            print(f"Error loading file: {e}")


def handle_dropdown_selection(selected_sheet: str, 
                              file_index: int, 
                              unique_key_dropdown: ctk.CTkComboBox, 
                              checkbox_frame: ctk.CTkScrollableFrame) -> None:
    """
    Loads headers from the selected sheet.
    If File 1 is selected, populates the unique key dropdown and column checkboxes.
    """
    print(f"File {file_index} selected sheet is set to: {selected_sheet}")

    config = load_config()
    config["service_two"][f"file{file_index}"]["sheet_name"] = selected_sheet
    save_config(config)

    file_path = config["service_two"][f"file{file_index}"]["file_path"]
    
    try:
        # Load headers
        df = pd.read_excel(file_path, sheet_name=selected_sheet, engine='openpyxl', nrows=0)
        column_list = df.columns.tolist()

        # Logic specific to File 1 (The Reference File)
        if file_index == 1:
            # Fill unique key dropdown
            unique_key_dropdown.configure(values=column_list)
            unique_key_dropdown.set("Select Unique Key")

            # Clear existing checkboxes
            for widget in checkbox_frame.winfo_children():
                widget.destroy()
            checkbox_vars.clear()
            
            # --- Helper: Select All ---
            def select_all() -> None:
                for var in checkbox_vars.values():
                    var.set(True)

            select_all_button = ctk.CTkButton(
                checkbox_frame, 
                text="Select All Columns", 
                command=select_all, 
                height=24,
                width=100
            )
            select_all_button.pack(anchor='w', pady=(0, 5))
            # --------------------------

            # Generate Checkboxes
            for col in column_list:
                var = ctk.BooleanVar()
                cb = ctk.CTkCheckBox(checkbox_frame, text=col, variable=var)
                cb.pack(anchor='w', pady=2)
                checkbox_vars[col] = var

    except Exception as e:
        logger.error(f"Error reading sheet {selected_sheet}: {e}")
        print(f"Error reading columns: {e}")


def on_unique_key_selected(selected: str) -> None:
    """Callback when a unique key is chosen."""
    config = load_config()
    config["service_two"]["unique_id_column"] = selected
    save_config(config)
    print(f"Unique key set to: {selected}")


def save_checked_columns() -> None:
    """Callback to save the columns selected via checkboxes."""
    selected_columns = [col for col, var in checkbox_vars.items() if var.get()]
    
    config = load_config()
    config["service_two"]["compare_columns"] = selected_columns
    save_config(config)
    
    print(f"Columns selected for comparison: {len(selected_columns)}")

# ==========================================
# 3. LAYOUT INITIALIZATION
# ==========================================

def setup_service_two(left_frame: ctk.CTkFrame, right_frame: ctk.CTkFrame) -> None:
    """Sets up the layout for the File Comparison service."""
    
    logger.info('Entered Service Two')

    # Clear frames & Reset state
    clear_frame(left_frame)
    clear_frame(right_frame)
    checkbox_vars.clear()

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
    # File Selection Area
    # ---------------------------
    
    # File 1 (Reference)
    file_button1 = ctk.CTkButton(
        left_frame, 
        text="1. Select Reference File", 
        command=lambda: select_file(file_button1, dropdown1, 1)
    )
    file_button1.grid(row=1, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    dropdown1 = ctk.CTkComboBox(
        left_frame, 
        values=[], 
        state="readonly", 
        command=lambda selected: handle_dropdown_selection(selected, 1, unique_key_dropdown, checkbox_frame)
    )
    dropdown1.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    dropdown1.set("Select Sheet")

    # File 2 (Target)
    file_button2 = ctk.CTkButton(
        left_frame, 
        text="2. Select Comparison File", 
        command=lambda: select_file(file_button2, dropdown2, 2)
    )
    file_button2.grid(row=3, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    dropdown2 = ctk.CTkComboBox(
        left_frame, 
        values=[], 
        state="readonly", 
        command=lambda selected: handle_dropdown_selection(selected, 2, unique_key_dropdown, checkbox_frame)
    )
    dropdown2.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    dropdown2.set("Select Sheet")

    # ---------------------------
    # Column Configuration
    # ---------------------------

    # Unique Key
    ctk.CTkLabel(left_frame, text="Unique Key:").grid(row=5, column=0, padx=10, sticky="w")
    unique_key_dropdown = ctk.CTkComboBox(
        left_frame, 
        state="readonly",
        command=on_unique_key_selected
    )
    unique_key_dropdown.grid(row=5, column=1, padx=10, sticky="ew")

    # Checkboxes Frame (Upgraded to ScrollableFrame)
    checkbox_frame = ctk.CTkScrollableFrame(left_frame, height=200)
    checkbox_frame.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
    
    # Save Columns Button
    save_button = ctk.CTkButton(
        left_frame, 
        text="3. Save Column Selection", 
        command=save_checked_columns
    )
    save_button.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

    # ---------------------------
    # Action
    # ---------------------------
    analyse_button = ctk.CTkButton(
        left_frame, 
        text="4. Run Comparison", 
        fg_color="green",
        hover_color="darkgreen",
        command=compare
    )
    analyse_button.grid(row=8, column=0, columnspan=2, pady=20, padx=10, sticky="ew")
    
    # ---------------------------
    # Grid Config
    # ---------------------------
    left_frame.grid_columnconfigure(0, weight=1)
    left_frame.grid_columnconfigure(1, weight=2)
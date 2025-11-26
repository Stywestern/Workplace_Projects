# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import logging
from typing import List, Optional
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog

# Local Imports
from _config.settings import load_config, save_config, pretty_print_config
from _utils.functions import clear_frame, auto_header_finder
from _gui.monitor import setup_monitor

# Service Specific Imports
from _service_one.service_logic import analyse
from _service_one.service_constants import entry_message

# Initialize Logger
logger = logging.getLogger('AppLogger')

# ==========================================
# 2. FILE & SHEET SELECTION LOGIC
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

            # Update Button Text
            file_name = file_path.split("/")[-1]
            button.configure(text=f"Selected file: {file_name}")
            print(f"File {file_index} is set to: {file_name}")

            # Save to config
            config["service_one"][f"file{file_index}"]["file_path"] = file_path
            save_config(config)
            
        except Exception as e:
            logger.error(f"Failed to load file {file_path}: {e}")
            print(f"Error loading file: {e}")


def handle_dropdown_selection(selected_sheet: str, 
                              file_index: int, 
                              unique_id_dropdown: ctk.CTkComboBox, 
                              forecast_dropdown: ctk.CTkComboBox, 
                              extra_columns_dropdown: Optional[ctk.CTkComboBox] = None) -> None:
    """
    Loads headers from the selected sheet and populates column dropdowns.
    """
    config = load_config()
    
    # Store the selected sheet name
    config["service_one"][f"file{file_index}"]["sheet_name"] = selected_sheet
    save_config(config)
    print(f"File {file_index} selected sheet is set to: {selected_sheet}")

    # Load the selected file's sheet
    file_data = config["service_one"][f"file{file_index}"]

    if file_data and file_data.get("file_path"):
        try:
            # Initial read to find header
            header_row = auto_header_finder(file_data)
            df = pd.read_excel(file_data["file_path"], sheet_name=selected_sheet, header=header_row)

            # Populate column dropdowns
            filtered_columns = populate_column_dropdowns(
                df, unique_id_dropdown, forecast_dropdown, extra_columns_dropdown, file_index
            )
            
        except Exception as e:
            logger.error(f"Error reading sheet {selected_sheet}: {e}")
            print(f"Error reading sheet: {e}")


# ==========================================
# 3. COLUMN HANDLING LOGIC
# ==========================================

def populate_column_dropdowns(df: pd.DataFrame, 
                              unique_id_dropdown: ctk.CTkComboBox, 
                              forecast_dropdown: ctk.CTkComboBox, 
                              extra_columns_dropdown: Optional[ctk.CTkComboBox], 
                              file_index: int) -> List[str]:
    """Filters columns and updates dropdown widgets."""
    config = load_config()
    
    # Get the column names (headers)
    col_names = sorted(list(df.columns))
    excluded_columns = ["Subregion"] # Logic to exclude specific internal columns
    
    # Get the already selected extra columns from the config
    selected_extra_columns = config["service_one"][f"file{file_index}"].get("extra_columns", [])

    # Filter columns
    filtered_columns = [col for col in col_names if col not in excluded_columns and col not in selected_extra_columns]

    # Update standard dropdowns
    unique_id_dropdown.configure(values=filtered_columns)
    forecast_dropdown.configure(values=filtered_columns)
    
    # Update extra columns dropdown only if it exists (Fixes bug where File 2 updated File 1)
    if extra_columns_dropdown is not None:
        extra_columns_dropdown.configure(values=filtered_columns)
    
    return filtered_columns


def select_extra_columns(extra_column: str, extra_columns_dropdown: ctk.CTkComboBox, file_index: int = 1) -> None:
    """Adds a column to the 'Extra Columns' list in config and removes it from the dropdown."""
    config = load_config()

    if extra_column:
        selected_extra_columns = config["service_one"][f"file{file_index}"].get("extra_columns", [])

        if extra_column not in selected_extra_columns:
            selected_extra_columns.append(extra_column)
            
            config["service_one"][f"file{file_index}"]["extra_columns"] = selected_extra_columns
            save_config(config)
            
            print(f"Selected Extra Columns: {', '.join(selected_extra_columns)}")

    # Update dropdown to remove selected item
    current_values = list(extra_columns_dropdown.cget('values'))
    new_values = [col for col in current_values if col != extra_column]
    extra_columns_dropdown.configure(values=new_values)
    extra_columns_dropdown.set("Select more columns")


def clear_extra_columns(file_index: int = 1) -> None:
    """Clears the extra column selection in config."""
    config = load_config()
    config["service_one"][f"file{file_index}"]["extra_columns"] = []
    save_config(config)
    print("Cleared all selected extra columns")


def confirm_inputs(unique_id_dropdown1: ctk.CTkComboBox, forecast_dropdown1: ctk.CTkComboBox, 
                   unique_id_dropdown2: ctk.CTkComboBox, forecast_dropdown2: ctk.CTkComboBox) -> None:
    """Validates and saves the final column mappings."""
    config = load_config()

    config["service_one"]["file1"]["unique_id_column"] = unique_id_dropdown1.get()
    config["service_one"]["file1"]["forecast_column"] = forecast_dropdown1.get()

    config["service_one"]["file2"]["unique_id_column"] = unique_id_dropdown2.get()
    config["service_one"]["file2"]["forecast_column"] = forecast_dropdown2.get()

    save_config(config)
    
    logger.info("User inputs confirmed and stored")
    print("User inputs confirmed and stored:")
    pretty_print_config("service_one")


# ==========================================
# 4. LAYOUT INITIALIZATION
# ==========================================

def setup_service_one(left_frame: ctk.CTkFrame, right_frame: ctk.CTkFrame) -> None:
    """Initializes the UI layout for Service One."""
    
    logger.info('Entered Service One')
    
    # Clear frames
    clear_frame(left_frame)
    clear_frame(right_frame)

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
    # File 1 Section
    # ---------------------------
    
    # File Selection
    file_button1 = ctk.CTkButton(
        left_frame, 
        text="1. Select Reference File", 
        command=lambda: select_file(file_button1, dropdown1, 1)
    )
    file_button1.grid(row=1, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # Sheet Selection
    dropdown1 = ctk.CTkComboBox(
        left_frame, 
        values=[], 
        state="readonly", 
        command=lambda selected: handle_dropdown_selection(
            selected, 1, unique_id_dropdown1, forecast_dropdown1, extra_columns_dropdown1
        )
    )
    dropdown1.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    dropdown1.set("Select Sheet")

    # Unique ID
    ctk.CTkLabel(left_frame, text="ID column:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
    unique_id_dropdown1 = ctk.CTkComboBox(left_frame, state="readonly")
    unique_id_dropdown1.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
    
    # Forecast Column
    ctk.CTkLabel(left_frame, text="Forecast column:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
    forecast_dropdown1 = ctk.CTkComboBox(left_frame, state="readonly")
    forecast_dropdown1.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

    # Extra Columns
    ctk.CTkLabel(left_frame, text="Extra columns:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
    extra_columns_dropdown1 = ctk.CTkComboBox(
        left_frame, 
        values=[], 
        state="readonly", 
        command=lambda selected: select_extra_columns(selected, extra_columns_dropdown1)
    )
    extra_columns_dropdown1.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

    # Clear Extra Columns
    clear_button1 = ctk.CTkButton(left_frame, text="Clear Extra Columns", command=clear_extra_columns)
    clear_button1.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

    # ---------------------------
    # File 2 Section
    # ---------------------------

    # File Selection
    file_button2 = ctk.CTkButton(
        left_frame, 
        text="2. Select Comparison File", 
        command=lambda: select_file(file_button2, dropdown2, 2)
    )
    file_button2.grid(row=7, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # Sheet Selection
    dropdown2 = ctk.CTkComboBox(
        left_frame, 
        values=[], 
        state="readonly", 
        # Note: extra_columns_dropdown is None here to prevent File 2 from updating File 1's UI
        command=lambda selected: handle_dropdown_selection(
            selected, 2, unique_id_dropdown2, forecast_dropdown2, extra_columns_dropdown=None
        )
    )
    dropdown2.grid(row=8, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    dropdown2.set("Select Sheet")

    # Unique ID
    ctk.CTkLabel(left_frame, text="ID column:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
    unique_id_dropdown2 = ctk.CTkComboBox(left_frame, state="readonly")
    unique_id_dropdown2.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
    
    # Forecast Column
    ctk.CTkLabel(left_frame, text="Forecast column:").grid(row=10, column=0, padx=5, pady=5, sticky="w")
    forecast_dropdown2 = ctk.CTkComboBox(left_frame, state="readonly")
    forecast_dropdown2.grid(row=10, column=1, padx=5, pady=5, sticky="ew")

    # ---------------------------
    # Actions
    # ---------------------------

    # Confirm Inputs
    confirm_button = ctk.CTkButton(
        left_frame, 
        text="3. Confirm Inputs", 
        command=lambda: confirm_inputs(unique_id_dropdown1, forecast_dropdown1, unique_id_dropdown2, forecast_dropdown2)
    )
    confirm_button.grid(row=11, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # Analyse
    analyse_button = ctk.CTkButton(
        left_frame, 
        text="4. Run Analysis", 
        fg_color="green", 
        hover_color="darkgreen",
        command=analyse
    )
    analyse_button.grid(row=13, column=0, columnspan=2, pady=20, padx=10, sticky="ew")
    
    # ---------------------------
    # Grid Config
    # ---------------------------
    left_frame.grid_columnconfigure(0, weight=1) 
    left_frame.grid_columnconfigure(1, weight=2)

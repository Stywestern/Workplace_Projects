# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import logging
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog

# Local Imports
from _config.settings import load_config, save_config
from _utils.functions import clear_frame
from _gui.monitor import setup_monitor

# Service Specific Imports
from _service_eight.service_logic import service_merge, save_to_excel
from _service_eight.service_constants import entry_message

# Initialize Logger
logger = logging.getLogger('AppLogger')

# ==========================================
# 2. SELECTION LOGIC
# ==========================================

def select_files(button: ctk.CTkButton, frame: ctk.CTkFrame) -> None:
    """
    Opens a file dialog for the user to select Excel files.
    Populates the frame with dropdowns to select specific sheets.
    """
    config = load_config()

    # Open file dialog
    selected_files = filedialog.askopenfilenames(
        title="Select Excel files", 
        filetypes=[("Excel files", "*.xlsx")]
    )

    if selected_files:
        # Update Config
        config["service_eight"]["file_paths"] = selected_files
        # Initialize sheet_names with empty strings
        config["service_eight"]["sheet_names"] = [""] * len(selected_files)
        save_config(config)

        # Clear previous widgets in the frame if any
        clear_frame(frame)

        print("Selected files:")
        for index, file_path in enumerate(selected_files):
            print(f" - {file_path}")
            
            try:
                # Read Excel file to get sheet names
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                sheet_names = excel_file.sheet_names

                # Create Label
                label = ctk.CTkLabel(frame, text=f"File {index + 1} Sheet:")
                label.grid(row=index, column=0, padx=5, pady=5, sticky='w')

                # Create Dropdown
                var = ctk.StringVar()
                dropdown = ctk.CTkComboBox(frame, values=sheet_names, variable=var)
                dropdown.grid(row=index, column=1, padx=5, pady=5, sticky='ew')

                # Define callback for this specific dropdown
                def on_sheet_select(choice: str, idx: int = index) -> None:
                    print(f"Selected file {idx + 1} sheet: {choice}")
                    # Reload config to ensure we don't overwrite other updates
                    current_config = load_config() 
                    current_config["service_eight"]["sheet_names"][idx] = choice
                    save_config(current_config)

                dropdown.configure(command=on_sheet_select)
            
            except Exception as e:
                logger.error(f"Error reading file {file_path}: {e}")
                print(f"Error reading file {file_path}")

        # Update button text to reflect status
        button.configure(text=f"Selected files: {len(selected_files)} files")
        print()


def create_column_dropdowns(frame: ctk.CTkFrame) -> None:
    """
    Reads the selected sheets and creates dropdowns to select 
    the 'merge key' column for each file.
    """
    config = load_config()
    
    row_offset = len(config["service_eight"]["file_paths"])

    # Initialize merging_columns config
    config["service_eight"]["merging_columns"] = [""] * row_offset
    save_config(config)

    for index, file_path in enumerate(config["service_eight"]["file_paths"]):
        sheet_name = config["service_eight"]["sheet_names"][index]

        if not sheet_name:
            print(f"Skipping file {index + 1}: No sheet selected.")
            continue

        try:
            # Read header only to get columns (efficient)
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', nrows=0)
            column_names = df.columns.tolist()

            # Create Label
            column_label = ctk.CTkLabel(frame, text=f"File {index + 1} Key Column:")
            column_label.grid(row=index + row_offset, column=0, padx=5, pady=5, sticky='w')

            # Create Dropdown
            column_var = ctk.StringVar()
            column_dropdown = ctk.CTkComboBox(frame, values=column_names, variable=column_var)
            column_dropdown.grid(row=index + row_offset, column=1, padx=5, pady=5, sticky='ew')

            # Define callback
            def on_column_select(choice: str, idx: int = index) -> None:
                print(f"Selected file {idx + 1} merge column: {choice}")
                current_config = load_config()
                current_config["service_eight"]["merging_columns"][idx] = choice
                save_config(current_config)

            column_dropdown.configure(command=on_column_select)

        except Exception as e:
            logger.error(f"Error processing columns for file {index + 1}: {e}")
            print(f"Error reading columns: {e}")


# ==========================================
# 3. MERGE & OUTPUT LOGIC
# ==========================================

def display_column_checkboxes(frame: ctk.CTkFrame) -> None:
    """
    Triggers the merge logic to preview columns, then displays 
    checkboxes for the user to select which columns to keep.
    """
    clear_frame(frame)
    
    try:
        print("Analyzing merged data structure...")
        # Assume service_merge() reads from config and returns a DataFrame
        merged_df = service_merge()
        
        if merged_df.empty:
            print("Warning: Merged dataset is empty.")
            return

        merged_columns = merged_df.columns.tolist()
        
        for col in merged_columns:
            checkbox = ctk.CTkCheckBox(frame, text=col)
            checkbox.pack(padx=10, pady=5, anchor='w')
            
    except Exception as e:
        logger.error(f"Merge failed: {e}")
        print(f"Merge process failed: {e}")


def get_selected_checkboxes(checkbox_frame: ctk.CTkFrame) -> None:
    """
    Retrieves selected columns from checkboxes, updates config, 
    and triggers the final save to Excel.
    """
    config = load_config()
    selected_columns = []
    
    # Iterate through children widgets to find checked boxes
    for widget in checkbox_frame.winfo_children():
        if isinstance(widget, ctk.CTkCheckBox):
            if widget.get() == 1:  # CheckBox value is 1 if checked
                selected_columns.append(widget.cget("text")) 
                
    config["service_eight"]["selected_columns"] = selected_columns
    save_config(config)
    
    print("\nSelected columns for export:")
    for column in selected_columns:
        print(f"- {column}")
        
    # Perform final save
    save_to_excel(service_merge())


# ==========================================
# 4. LAYOUT INITIALIZATION
# ==========================================

def setup_service_eight(left_frame: ctk.CTkFrame, right_frame: ctk.CTkFrame) -> None:
    """Initializes the UI layout for Service Eight."""
    
    logger.info('Entered Service Eight')
    
    # Clean slate
    clear_frame(left_frame)
    clear_frame(right_frame)

    # Delayed import to avoid circular dependency
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
    
    # File Selection Button
    file_button = ctk.CTkButton(
        left_frame, 
        text="1. Select Files", 
        command=lambda: select_files(file_button, dropdown_frame)
    )
    file_button.grid(row=1, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # Column Selection Button
    column_button = ctk.CTkButton(
        left_frame, 
        text="2. Select Merge Keys", 
        command=lambda: create_column_dropdowns(dropdown_frame)
    )
    column_button.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    
    # Scrollable Area for Dropdowns (Sheets & Columns)
    dropdown_frame = ctk.CTkScrollableFrame(left_frame, height=200)
    dropdown_frame.grid(row=3, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    
    # Merge Preview Button
    merge_button = ctk.CTkButton(
        left_frame, 
        text="3. Analyze & Filter", 
        command=lambda: display_column_checkboxes(checkbox_frame)
    )
    merge_button.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # Scrollable Area for Column Checkboxes
    checkbox_frame = ctk.CTkScrollableFrame(left_frame, height=200)
    checkbox_frame.grid(row=5, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

    # Final Save Button
    confirm_button = ctk.CTkButton(
        left_frame, 
        text="4. Confirm & Save", 
        fg_color="green", # Visual distinction for final action
        hover_color="darkgreen",
        command=lambda: get_selected_checkboxes(checkbox_frame)
    )
    confirm_button.grid(row=6, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    
    # Grid Configuration
    left_frame.grid_columnconfigure(0, weight=1)
    left_frame.grid_columnconfigure(1, weight=1)

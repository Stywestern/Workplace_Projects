# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import os
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk

# Local Utilities
from _utils.functions import clear_frame
from _utils.constants_n_variables import welcome_message
from _gui.monitor import setup_monitor

# Constants
ICON_PATH = "inputs/Microsoft.ico" 

# ==========================================
# 2. LOADING SCREENS
# ==========================================

def show_loading_screen(app: ctk.CTk) -> ctk.CTkToplevel:
    """Displays a simple loading modal (CustomTkinter version)."""
    loading_screen = ctk.CTkToplevel(app)
    loading_screen.title("Session started")
    loading_screen.geometry("300x100")
    
    # Safe icon loading
    if os.path.exists(ICON_PATH):
        loading_screen.iconbitmap(ICON_PATH)
        
    loading_label = ctk.CTkLabel(
        loading_screen, 
        text="Session started, please wait...", 
        font=("Arial", 14)
    )
    loading_label.pack(pady=20)

    return loading_screen


def show_loading_screen_in() -> tk.Toplevel:
    """Displays a complex loading modal with progress bar (Standard Tkinter version)."""
    
    def cancel_action() -> None:
        print("Process cancelled")
        loading_screen.destroy()

    loading_screen = tk.Toplevel()
    loading_screen.title("Loading")
    loading_screen.geometry("680x200")
    loading_screen.configure(bg='#333333')

    # Safe icon loading
    if os.path.exists(ICON_PATH):
        loading_screen.iconbitmap(ICON_PATH)

    label = tk.Label(
        loading_screen, 
        text="Processing, please wait...", 
        font=("Helvetica", 20), 
        fg="white", 
        bg="#333333"
    )
    label.pack(pady=20)

    # Configure Progress Bar Style
    style = ttk.Style()
    style.theme_use('clam') # Note: This sets the global theme for ttk widgets
    style.configure(
        "TProgressbar", 
        thickness=30, 
        troughcolor='#333333', 
        background='gray'
    )

    progress_bar = ttk.Progressbar(
        loading_screen, 
        mode='indeterminate', 
        style="TProgressbar"
    )
    progress_bar.pack(expand=True, pady=20, padx=20, fill='x')
    progress_bar.start()

    cancel_button = tk.Button(
        loading_screen, 
        text="Stop Operation",
        command=cancel_action, 
        font=("Helvetica", 18), 
        bg="gray", 
        fg="white"
    )
    cancel_button.pack(pady=10)
    
    return loading_screen


# ==========================================
# 3. MAIN LAYOUT LOGIC
# ==========================================

def setup_default_layout(left_frame: ctk.CTkFrame, right_frame: ctk.CTkFrame) -> None:
    """Initializes the main navigation menu and default view."""
    
    # Delayed imports to avoid circular dependencies
    from _gui.service_one_layout import setup_service_one
    from _gui.service_two_layout import setup_service_two
    from _gui.service_eight_layout import setup_service_eight
    from _gui.service_three_layout import setup_service_three
    
    # Clear the frames to ensure a clean slate
    clear_frame(left_frame)
    clear_frame(right_frame)
    
    # Setup Monitor View (Default Right Frame Content)
    setup_monitor(right_frame)
    
    # Optional: Log the welcome message if needed
    # print(welcome_message) 

    # ------------------------------------------
    # Navigation Buttons (Left Frame)
    # ------------------------------------------

    # Service 8: Merge Excels
    service8_button = ctk.CTkButton(
        master=left_frame, 
        text="Merge Excels", 
        command=lambda: setup_service_eight(left_frame, right_frame)
    )
    service8_button.pack(padx=10, pady=10)
    
    # Service 1: Forecast Comparison
    service1_button = ctk.CTkButton(
        master=left_frame, 
        text="Forecast Comparison", 
        command=lambda: setup_service_one(left_frame, right_frame)
    )
    service1_button.pack(padx=10, pady=10)
    
    # Service 2: Detect Changes
    service2_button = ctk.CTkButton(
        master=left_frame, 
        text="Detect Changes", 
        command=lambda: setup_service_two(left_frame, right_frame)
    )
    service2_button.pack(padx=10, pady=10)
    
    # Service 3: Service Usage Report
    service3_button = ctk.CTkButton(
        master=left_frame, 
        text="Service Usage Report", 
        command=lambda: setup_service_three(left_frame, right_frame)
    )
    service3_button.pack(padx=10, pady=10)

# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import sys
import os
import logging
import warnings
import customtkinter as ctk

# Local Modules
from _config.settings import load_config, terminate_program
from _gui.layout import setup_default_layout

# Suppress specific warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl", message="Workbook contains no default style")
warnings.simplefilter(action='ignore', category=FutureWarning)

# ==========================================
# 2. CONFIGURATION & LOGGING
# ==========================================

# Load configuration
config = load_config()

original_stdout = sys.stdout
original_stderr = sys.stderr

# Configure logging
if not os.path.exists("outputs"):
    os.makedirs("outputs")

with open("outputs/app.log", "w"):
    pass

logging.basicConfig(
    filename="outputs/app.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(funcName)s() - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

logger = logging.getLogger('AppLogger')
logger.info("Logging system initialized")


# ==========================================
# 3. APP INITIALIZATION
# ==========================================

app = ctk.CTk()
app.title("Pipeline Velocity Tracker")
app.geometry("920x520")

# Icon setup
app.iconbitmap("inputs/Microsoft.ico")
    
ctk.set_appearance_mode("dark")

# ==========================================
# 4. LAYOUT & GRID CONFIGURATION
# ==========================================

# Configure main app grid weights
app.grid_rowconfigure(0, weight=1)
app.grid_columnconfigure(0, weight=1) # Left frame expandable
app.grid_columnconfigure(1, weight=1) # Right frame expandable


# ==========================================
# 5. UI COMPONENTS
# ==========================================

# Left Frame
left_frame = ctk.CTkFrame(master=app, width=400, height=800)
left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=2)
left_frame.grid_rowconfigure(0, weight=1)
left_frame.grid_columnconfigure(0, weight=1)

# Scrollable Frame inside Left Frame
scroll_frame = ctk.CTkScrollableFrame(master=left_frame)
scroll_frame.grid(row=0, column=0, sticky="nsew")

# Right Frame
right_frame = ctk.CTkFrame(master=app, width=650, height=800)
right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=7)
right_frame.grid_rowconfigure(0, weight=1)
right_frame.grid_columnconfigure(0, weight=1)

# Footer Label
footer_label = ctk.CTkLabel(
        master=app,
        text="Data Analysis Services - Local Version",
        font=("Arial", 10)
    )
footer_label.grid(row=1, column=0, columnspan=2, pady=10)


# ==========================================
# 6. LOGIC & PROTOCOLS
# ==========================================

# Initialize the default layout logic
setup_default_layout(scroll_frame, right_frame)

def on_closing() -> None:
    """Handle the graceful shutdown of the application."""
    try:
        terminate_program()
    finally:
        sys.stdout = original_stdout
        sys.stderr = original_stderr
        app.quit()

# Bind the custom closing event
app.protocol("WM_DELETE_WINDOW", on_closing)


# ==========================================
# 7. EXECUTION
# ==========================================

if __name__ == "__main__":
    app.mainloop()




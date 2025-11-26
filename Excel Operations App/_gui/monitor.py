# ==========================================
# 1. IMPORTS
# ==========================================

# Third-Party Libraries
import sys
import customtkinter as ctk

# ==========================================
# 2. CLASSES
# ==========================================

class MonitorRedirector:
    """
    A class to redirect sys.stdout and sys.stderr to a generic
    CustomTkinter text widget, creating an internal console.
    """
    def __init__(self, output_widget: ctk.CTkTextbox):
        self.output_widget = output_widget
    
    def write(self, text: str) -> None:
        """Writes text to the widget and scrolls to the bottom."""
        self.output_widget.insert('end', text)
        self.output_widget.see('end')

    def flush(self) -> None:
        """
        Required method for file-like objects (sys.stdout).
        Left empty as GUI updates happen immediately in 'write'.
        """
        pass


# ==========================================
# 3. SETUP FUNCTIONS
# ==========================================

def setup_monitor(right_frame: ctk.CTkFrame) -> None:
    """Initializes the text box and redirects system output to it."""
    
    # Create the display widget
    monitor = ctk.CTkTextbox(
        master=right_frame, 
        height=400, 
        width=400, 
        wrap="word"
    )
    monitor.pack(padx=10, pady=10, fill="both", expand=True)
    
    # Initialize the redirector
    stdout_redirector = MonitorRedirector(monitor)
    
    # Override system standard output and error
    sys.stdout = stdout_redirector 
    sys.stderr = stdout_redirector
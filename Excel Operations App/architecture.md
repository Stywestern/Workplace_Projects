# System Architecture

## 1. High-Level Data Flow

1.  **Initialization:**
    * The user runs `main.py`.
    * The system initializes the root `CustomTkinter` window, sets up the `AppLogger`, and creates the `outputs/` directory structure.
    * `_config/settings.py` loads `config.json` to restore the last known state (e.g., last used file paths).

2.  **Navigation & Selection (GUI Layer):**
    * The `_gui/layout.py` orchestrates the main navigation.
    * When a user selects a tool (e.g., "Anomaly Detection"), the main frame is cleared, and the specific service layout (e.g., `_gui/service_three_layout.py`) is injected.
    * User inputs (file selections, dropdowns) are immediately persisted to `config.json` via the `save_config()` utility.

3.  **Execution Trigger:**
    * The user clicks an Action Button (e.g., "Analyze", "Merge").
    * The GUI component calls the corresponding function in the **Logic Layer** (e.g., `_service_three/service_logic.py -> analyse()`).

4.  **Data Processing (Logic Layer):**
    * The logic module reads parameters from the `config.json` (decoupling it from the GUI widgets).
    * **Pandas** reads the Excel files.
    * Specific algorithms (comparison, anomaly detection, merging) process the dataframes.

5.  **Output Generation:**
    * Results are generated using **OpenPyXL** (for Excel reports) or **Matplotlib + Jinja2** (for HTML/Image reports).
    * Files are saved to the `outputs/` directory.

6.  **Presentation:**
    * The system uses `os.startfile` (or platform equivalent) to automatically open the generated report for the user.
    * Status updates and errors are redirected to the GUI "Monitor" (text box) via `sys.stdout` interception.

## 2. Key Design Decisions

### Modular Service Architecture (Micro-Service Style)
Instead of a monolithic script, the app is split into isolated services directories (`_service_one`, `_service_two`, etc.).

* **Benefit:** A bug in the "Excel Merger" will not break or stop the "Anomaly Detection" from working.
* **Scalability:** Adding "Service Five" only requires creating a new folder and adding one button to the main layout.

### Configuration-Driven State
We avoid passing dozens of variables between functions. Instead, the application uses a "Single Source of Truth": `config/config.json`.

* **Pattern:** `GUI updates JSON` -> `Logic reads JSON`.
* **Benefit:** This provides persistent memory. If the user closes the app and reopens it, their selected file paths are remembered, improving the UX.

### Separation of Concerns (GUI vs. Logic)
Strict separation is enforced between how the app *looks* and what the app *does*.

* **`_gui/`**: Contains only CustomTkinter widgets and layout grids. It knows *nothing* about math or data analysis.
* **`_service_*/service_logic.py`**: Contains only Pandas and Math logic. It knows *nothing* about buttons or windows.
* **Bridge:** The logic functions are triggered by the GUI command callbacks.

### HTML Reporting with Jinja2
For complex analysis (Service 3), simple Excel dumps are insufficient.

* **Decision:** We use **Jinja2** templates combined with **Base64 encoded Matplotlib plots**.
* **Benefit:** This creates portable, self-contained HTML reports that can be emailed or shared without needing to send a folder of separate image files.

## 3. Process Chart

The following diagram illustrates the lifecycle of a typical user interaction (e.g., running the Anomaly Detection Service).

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant Main as Main Window (main.py)
    participant Layout as Service Layout (_gui/*_layout.py)
    participant Config as Config Store (config.json)
    participant Logic as Business Logic (_service_*/logic.py)
    participant FS as File System

    Note over User, Main: Phase 1: Setup
    User->>Main: Launch Application
    Main->>Config: Load Settings
    Main->>Layout: Render Main Menu

    Note over User, Config: Phase 2: Input & State
    User->>Layout: Click "Service 3 (Anomaly)"
    Layout->>User: Display Service 3 UI
    User->>Layout: Select Excel File
    Layout->>Config: Save "file_path" to JSON
    User->>Layout: Select Sheet Name
    Layout->>Config: Save "sheet_name" to JSON

    Note over User, Logic: Phase 3: Execution
    User->>Layout: Click "Run Analysis"
    Layout->>Logic: Call analyse()
    
    Logic->>Config: Read File Paths & Parameters
    Logic->>FS: Read Input Excel File (Pandas)
    
    loop Data Processing
        Logic->>Logic: Clean Data
        Logic->>Logic: Detect Anomalies
        Logic->>Logic: Generate Plots (Matplotlib)
    end

    Note over Logic, FS: Phase 4: Output
    Logic->>FS: Save HTML Report (Jinja2 Render)
    Logic->>FS: Save Excel Data Dump
    Logic->>FS: Open Result File (OS Default App)
    
    Logic-->>Layout: Print "Done" to Monitor
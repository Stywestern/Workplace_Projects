# Data Analysis & Reporting Suite

A modular Python desktop application designed to streamline data operations on Excel files. This tool provides a modern graphical interface for comparing datasets, detecting anomalies, and merging multiple data sources.

## Key Features

The application is divided into specialized services:

* **Forecast Comparison (Service 1)**
    * Compares two Excel datasets based on a unique ID.
    * Aligns data and calculates difference and similarity percentages between forecast columns.
    * Groups data by "Subregion" (if available) for detailed granular reporting.

* **Data Change Detection (Service 2)**
    * Detects cell-level changes between two versions of a dataset.
    * Generates a highlighted Excel report showing exactly which values changed (e.g., `"Old Value → New Value"`).
    * Optimized to filter and display only the rows where changes occurred.

* **Anomaly Detection & Usage Reporting (Service 3)**
    * Analyzes time-series data to flag statistical anomalies in service usage.
    * Generates interactive HTML reports containing trend lines, bar charts, and detailed metrics.
    * Uses statistical thresholds to identify sustained shifts in data.

* **Excel Merger (Service 8)**
    * Merges multiple Excel files (or specific sheets) into a single master file.
    * Aligns data based on a user-selected key column.

## Tech Stack

* **Language:** Python 3.10+
* **GUI Framework:** CustomTkinter
* **Data Processing:** Pandas, NumPy, OpenPyXL
* **Visualization:** Matplotlib
* **Reporting:** Jinja2 (HTML Generation)

## Installation & Setup

1.  **Clone the repository**
    ```bash
    git clone <your-repository-url>
    cd excel_operations_app
    ```

2.  **Install Dependencies**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Ensure your virtual environment is activated.
2.  Run the main application entry point:
    ```bash
    python main.py
    ```
3.  The GUI will launch. Use the left-hand navigation menu to select the desired service.
4.  Follow the on-screen instructions to select your Input files.
5.  **Outputs:** All generated reports, merged files, and plots are automatically saved to the `outputs/` directory.

## Project Structure

```text
├── main.py                     # Application Entry Point
├── requirements.txt            # Python Dependencies
├── .gitignore                  # Git Ignore rules
├── html_files/                 # Jinja2 templates for HTML reports
│   └── s4_anomaly_report.html
├── _config/                    # Configuration loading and saving logic
│   └── settings.py
├── _gui/                       # GUI Layouts and Monitor components
│   ├── layout.py
│   ├── monitor.py
│   └── ... (Service layouts)
├── _utils/                     # Shared helper functions (Date parsing, UI clearing)
│   └── functions.py
├── _service_one/               # Logic: Forecast Comparison
├── _service_two/               # Logic: Change Detection
├── _service_three/             # Logic: Anomaly Detection
└── _service_eight/             # Logic: Excel Merging

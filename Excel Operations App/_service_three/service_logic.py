# ==========================================
# 1. IMPORTS
# ==========================================

# Third-Party Libraries
import os
import json
import base64
import logging
import traceback
import webbrowser
import warnings
from io import BytesIO
from typing import Optional, Any

import pandas as pd
import matplotlib
matplotlib.use('Agg') # Non-interactive backend for generating images
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from jinja2 import Environment, FileSystemLoader
from markupsafe import Markup

# Local Imports
from _config.settings import load_config

# Initialize Logger
logger = logging.getLogger('AppLogger')

# Suppress specific warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# ==========================================
# 2. CONSTANTS (Column Mappings)
# ==========================================
# Modify these if your Excel headers change
COL_COMPANY = 'TopParent'
COL_ID = 'TPID'
COL_DATE = 'FiscalDate'
COL_SERVICE = 'ServiceLevel2'
COL_USAGE = 'Acr'

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================

def tojson_filter(value: Any) -> Markup:
    """Custom Jinja2 filter to serialize data to JSON."""
    return Markup(json.dumps(value))

def calculate_total_usage(group: pd.DataFrame) -> pd.DataFrame:
    """Calculates total usage for a group and appends it as a summary row."""
    total_usage = group[COL_USAGE].sum()
    
    # Create summary row matching the dataframe structure
    total_usage_row = pd.DataFrame({
        COL_ID: [group[COL_ID].iloc[0]],
        COL_COMPANY: [group[COL_COMPANY].iloc[0]],
        COL_DATE: [group[COL_DATE].iloc[0]],
        COL_SERVICE: ['Total Usage'],
        COL_USAGE: [total_usage]
    })
    return pd.concat([group, total_usage_row], ignore_index=True)


# ==========================================
# 4. ANOMALY DETECTION LOGIC
# ==========================================

def flag_anomalies(df: pd.DataFrame, 
                  anomaly_threshold_percent: int = 20, 
                  anomaly_threshold_abs: int = 60, 
                  recent_window_size: int = 7, 
                  shift_window: int = 3) -> pd.DataFrame:
    """
    Flags anomalies based on a sustained shift in usage.
    """
    df_flags = df.copy()
    df_flags['anomaly'] = False
    df_flags['anomaly_direction'] = None

    # Group by Company and Service
    for _, group in df_flags.groupby([COL_COMPANY, COL_SERVICE]):
        group = group.sort_values(by=COL_DATE).reset_index()
        consecutive_anomalies = 0

        for i in range(recent_window_size, len(group)):
            # Define window
            historical_window = group.loc[i - recent_window_size:i - 1, COL_USAGE]
            current_usage = group.loc[i, COL_USAGE]

            if historical_window.empty:
                continue

            baseline_usage = historical_window.median()

            if baseline_usage != 0:
                percent_change = abs((current_usage - baseline_usage) / baseline_usage) * 100
                absolute_change = abs(current_usage - baseline_usage)

                is_single_anomaly = (percent_change >= anomaly_threshold_percent) and (absolute_change >= anomaly_threshold_abs)

                if is_single_anomaly:
                    consecutive_anomalies += 1
                    if consecutive_anomalies >= shift_window:
                        # Flag the shift window
                        for j in range(i - shift_window + 1, i + 1):
                            original_index = group.loc[j, 'index']
                            # Only flag if not already flagged to preserve direction
                            if not df_flags.loc[original_index, 'anomaly']:
                                df_flags.loc[original_index, 'anomaly'] = True
                                direction = 'increase' if group.loc[j, COL_USAGE] > baseline_usage else 'decrease'
                                df_flags.loc[original_index, 'anomaly_direction'] = direction
                else:
                    consecutive_anomalies = 0

    return df_flags


def get_anomalous_instances(df: pd.DataFrame) -> pd.DataFrame:
    """Returns rows for (company, service) pairs that have at least one anomaly."""
    # Identify pairs that have an anomaly
    anomalous_pairs = set(
        df[df['anomaly'] == True][[COL_COMPANY, COL_SERVICE]]
        .itertuples(index=False, name=None)
    )

    # Filter main dataframe to only include those pairs
    def is_anomalous_pair(row):
        return (row[COL_COMPANY], row[COL_SERVICE]) in anomalous_pairs

    return df[df.apply(is_anomalous_pair, axis=1)].copy()


# ==========================================
# 5. PLOTTING FUNCTIONS
# ==========================================

def plot_acr_with_outliers(dates, values, anomalies, title, save_path=None) -> None:
    """Plots Usage over time, highlighting outlier points."""
    fig, ax = plt.subplots(figsize=(12, 6))
    
    ax.scatter(dates, values, label='Usage', s=30)
    ax.scatter(dates[anomalies], values[anomalies], color='red', label='Anomaly', s=50)
    
    ax.set_xlabel('Date')
    ax.set_ylabel('Usage')
    ax.set_title(title)
    ax.legend()
    ax.grid(True)
    
    # Date formatting
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    fig.autofmt_xdate()

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path)
        plt.close(fig)
    else:
        plt.show()


def generate_flags_per_service_chart(df_anomalies_final: pd.DataFrame) -> Optional[str]:
    """Generates a base64 string of a horizontal bar chart."""
    service_counts = {}
    
    for service in df_anomalies_final[COL_SERVICE].unique():
        subset = df_anomalies_final[
            (df_anomalies_final[COL_SERVICE] == service) & 
            (df_anomalies_final['anomaly'] == True)
        ]
        count = subset[COL_COMPANY].nunique()
        if count > 0:
            service_counts[service] = count

    if not service_counts:
        return None

    top_10 = pd.Series(service_counts).sort_values(ascending=False).head(10)

    try:
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.barh(top_10.index, top_10.values, color='#8fbc8f')
        ax.set_xlabel('Unique Companies Flagging Service')
        ax.set_title('Top 10 Anomalous Services')
        plt.tight_layout()

        buffer = BytesIO()
        plt.savefig(buffer, format='png')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        return image_base64
    except Exception as e:
        logger.error(f"Error plotting flags chart: {e}")
        return None


def generate_anomaly_trend_chart(df_anomalies_final: pd.DataFrame) -> Optional[str]:
    """Generates a base64 string of the anomaly trend line chart."""
    anomalies_by_date = df_anomalies_final[df_anomalies_final['anomaly'] == True].groupby(COL_DATE)[COL_COMPANY].nunique()

    if anomalies_by_date.empty:
        return None

    try:
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.plot(anomalies_by_date.index, anomalies_by_date.values, marker='o', linestyle='-', color='#4c72b0')
        ax.set_title('Trend of Anomalies Over Time')
        ax.grid(True)
        
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
        fig.autofmt_xdate()
        plt.tight_layout()

        buffer = BytesIO()
        plt.savefig(buffer, format='png')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        return image_base64
    except Exception as e:
        logger.error(f"Error plotting trend chart: {e}")
        return None


def generate_usage_pie_chart(total_usage_per_service: pd.Series) -> Optional[str]:
    """Generates a base64 string of the usage pie chart."""
    if total_usage_per_service.empty:
        return None

    sorted_usage = total_usage_per_service.sort_values(ascending=False)
    top_9 = sorted_usage.head(9)
    other = sorted_usage[9:].sum()

    labels = top_9.index.tolist()
    sizes = top_9.values.tolist()

    if other > 0:
        labels.append('Other')
        sizes.append(other)

    try:
        fig, ax = plt.subplots(figsize=(8, 8))
        ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
        ax.set_title('Usage Share by Top 10 Services')
        plt.tight_layout()

        buffer = BytesIO()
        plt.savefig(buffer, format='png')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        return image_base64
    except Exception as e:
        logger.error(f"Error plotting pie chart: {e}")
        return None


# ==========================================
# 6. MAIN ANALYSIS ORCHESTRATOR
# ==========================================

def analyse() -> None:
    print("\nStarting Usage Report Analysis...")
    logger.info("Started Service Three Analysis")

    # 1. Load Config & Data
    try:
        config = load_config()
        file_path = config["service_three"]["file_path"]
        sheet = config["service_three"]["sheet_name"]
        
        print("Reading Excel file (this may take a moment)...")
        df = pd.read_excel(file_path, sheet_name=sheet)
        
        # Verify columns exist
        required_cols = [COL_ID, COL_COMPANY, COL_DATE, COL_SERVICE, COL_USAGE]
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            raise ValueError(f"Missing required columns in Excel: {missing}")

    except Exception as e:
        logger.error(f"Initialization Error: {e}")
        print(f"Error loading data: {e}")
        return

    # 2. Preprocess Data
    try:
        print("Preprocessing data...")
        df_trim = df[required_cols].copy()
        df_trim[COL_DATE] = pd.to_datetime(df_trim[COL_DATE])
        
        # Filter last 30 days
        latest_date = df_trim[COL_DATE].max()
        x_days_ago = latest_date - pd.Timedelta(days=30)
        df_trim = df_trim[df_trim[COL_DATE] >= x_days_ago].copy()

        # Aggregation
        df_agg = df_trim.groupby([COL_ID, COL_COMPANY, COL_DATE, COL_SERVICE])[COL_USAGE].sum().reset_index()
        df_agg = df_agg.groupby([COL_ID, COL_DATE], group_keys=False).apply(calculate_total_usage).reset_index(drop=True)

    except Exception as e:
        logger.error(f"Processing Error: {e}")
        print(f"Error processing data: {e}")
        return

    # 3. Detect Anomalies
    try:
        print("Running anomaly detection algorithms...")
        df_anomalies = flag_anomalies(df_agg)
        df_final = get_anomalous_instances(df_anomalies)
        
        # Calculate Stats
        total_companies = df_agg[COL_COMPANY].nunique()
        total_anomalous_services = df_final.groupby([COL_COMPANY, COL_SERVICE]).ngroups
        
        # Direction counts
        direction_counts = df_final[df_final['anomaly']].groupby(COL_SERVICE)['anomaly_direction'].apply(lambda x: x.value_counts().idxmax())
        increase_count = len(direction_counts[direction_counts == 'increase'])
        decrease_count = len(direction_counts[direction_counts == 'decrease'])

    except Exception as e:
        logger.error(f"Algorithm Error: {e}")
        print(f"Error in detection algorithm: {e}")
        return

    # 4. Generate Visualizations
    try:
        print("Generating charts...")
        total_usage_per_service = df_agg[df_agg[COL_SERVICE] != 'Total Usage'].groupby(COL_SERVICE)[COL_USAGE].sum()
        
        img_flags = generate_flags_per_service_chart(df_final)
        img_trend = generate_anomaly_trend_chart(df_final)
        img_pie = generate_usage_pie_chart(total_usage_per_service)
        
        # Generate Individual Plots
        plot_data = []
        file_basename = os.path.basename(file_path).rsplit('.', 1)[0]
        output_dir = f'outputs/anomaly_detection/plots_{file_basename}'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        grouped = df_final.groupby([COL_COMPANY, COL_SERVICE])
        
        print(f"Generating {len(grouped)} individual time-series plots...")
        for (company, service), group in grouped:
            safe_comp = "".join(x for x in str(company) if x.isalnum() or x in " -_").strip()
            safe_serv = "".join(x for x in str(service) if x.isalnum() or x in " -_").strip()
            
            plot_title = f"{company} - {service}"
            filename = f"{safe_comp}_{safe_serv}.png"
            save_path = os.path.join(output_dir, filename)
            
            plot_acr_with_outliers(group[COL_DATE], group[COL_USAGE], group['anomaly'], plot_title, save_path)
            
            plot_data.append({
                'path': os.path.abspath(save_path),
                'company_service': plot_title,
                'median_acr': group[COL_USAGE].median()
            })

    except Exception as e:
        logger.error(f"Visualization Error: {e}")
        print(f"Error generating charts: {e}")
        return

    # 5. Generate HTML Report
    try:
        print("Compiling HTML report...")
        template_path = "html_files/s4_anomaly_report.html"
        
        if not os.path.exists(template_path):
            logger.warning(f"Template not found at {template_path}. Skipping HTML generation.")
            print(f"Warning: HTML template missing at {template_path}. Analysis finished but report not generated.")
            return

        env = Environment(loader=FileSystemLoader(os.path.dirname(template_path)))
        env.filters['tojson'] = tojson_filter 
        template = env.get_template(os.path.basename(template_path))

        # Sort data for display
        plot_data_company = sorted(plot_data, key=lambda x: x['company_service'])
        plot_data_acr = sorted(plot_data, key=lambda x: x['median_acr'], reverse=True)
        
        # Mocking top affected (simplified)
        top_affected = df_final.groupby(COL_COMPANY)[COL_SERVICE].nunique().sort_values(ascending=False).head(3)

        html_content = template.render(
            report_title="Anomalous Service Report",
            report_heading="Anomalous Service Analysis",
            summary_start=df_agg[COL_DATE].min().strftime('%B %d, %Y'),
            summary_end=df_agg[COL_DATE].max().strftime('%B %d, %Y'),
            total_companies=total_companies,
            total_anomalous_services=total_anomalous_services,
            increase_count=increase_count,
            decrease_count=decrease_count,
            top_affected_companies=top_affected,
            flags_per_service_chart=img_flags,
            anomaly_trend_chart=img_trend,
            usage_pie_chart=img_pie,
            plot_data_company=plot_data_company,
            plot_data_acr=plot_data_acr,
            direction_service_company_counts={} # Simplified for now
        )

        report_path = f'outputs/anomaly_detection/Report_{file_basename}.html'
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        print("Analysis Complete.")
        webbrowser.open('file://' + os.path.abspath(report_path))

    except Exception as e:
        logger.error(f"Reporting Error: {e}")
        print(f"Error generating HTML report: {e}")
        traceback.print_exc()
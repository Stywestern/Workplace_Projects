# ==========================================
# 1. IMPORTS & SETUP
# ==========================================

# Third-Party Libraries
import os
import json
import logging
from typing import Dict, Any

# Define the config path relative to the current working directory or absolute
CONFIG_FILE = "config/config.json"

# Initialize Logger
logger = logging.getLogger('AppLogger')

# ==========================================
# 2. CORE IO FUNCTIONS
# ==========================================

def load_config() -> Dict[str, Any]:
    """Loads the configuration from the JSON file."""
    dir_path = os.path.dirname(CONFIG_FILE)
    
    # Create the directory if it doesn't exist
    if dir_path and not os.path.exists(dir_path):
        os.makedirs(dir_path)   

    # Check if the config file exists and load it
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                logger.info("Config file loaded successfully.")
                return json.load(f)
        except json.JSONDecodeError:
            logger.error("Config file is corrupted. Returning empty config.")
            return {}
            
    # Return an empty dictionary if the file doesn't exist
    return {}


def save_config(config: Dict[str, Any]) -> None:
    """Saves the dictionary configuration to the JSON file."""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4) # Added indent for readability
            logger.info("Config file saved successfully.")
    except IOError as e:
        logger.error(f"Failed to save config: {e}")


# ==========================================
# 3. SERVICE CONFIGURATION MANAGEMENT
# ==========================================

def clear_service_one_config() -> None:
    """Resets configuration for Service One."""
    config = load_config()

    # Ensure keys exist before clearing to avoid KeyError if config is empty
    if "service_one" not in config:
        config["service_one"] = {}

    config["service_one"]["file1"] = {
        "file_path": "",
        "sheet_name": "",
        "unique_id_column": "",
        "forecast_column": "",
        "extra_columns": []
    }
    config["service_one"]["file2"] = {
        "file_path": "",
        "sheet_name": "",
        "unique_id_column": "",
        "forecast_column": ""
    }

    logger.info("Config for service_one cleared")
    save_config(config)


def clear_service_two_config() -> None:
    """Resets configuration for Service Two."""
    config = load_config()

    if "service_two" not in config:
        config["service_two"] = {}

    # Fixed: Now correctly targets service_two instead of service_one
    config["service_two"]["file1"] = {
        "file_path": "",
        "sheet_name": "",
    }
    config["service_two"]["file2"] = {
        "file_path": "",
        "sheet_name": "",
    }
    config["service_two"]["unique_id_column"] = ""
    config["service_two"]["compare_columns"] = []

    logger.info("Config for service_two cleared")
    save_config(config)


def clear_service_eight_config() -> None:
    """Resets configuration for Service Eight."""
    config = load_config()

    if "service_eight" not in config:
        config["service_eight"] = {}

    config["service_eight"]["file_paths"] = []
    config["service_eight"]["sheet_names"] = []
    config["service_eight"]["merging_columns"] = []
    config["service_eight"]["selected_columns"] = []
    
    logger.info("Config for service_eight cleared")
    save_config(config)


# ==========================================
# 4. UTILITIES & PROTOCOLS
# ==========================================

def terminate_program() -> None:
    """Clears all temporary configurations before exit."""
    clear_service_one_config()
    clear_service_two_config()
    clear_service_eight_config()


def pretty_print_config(service: str) -> None:
    """Debug utility to print the current config of a specific service."""
    config = load_config()
    
    if service == "service_one" and "service_one" in config:
        for file_key, file_data in config["service_one"].items():
            print(f"\n##### {file_key.capitalize()} #####")
            print(f"File Path: {file_data.get('file_path', 'N/A')}")
            print(f"Sheet Name: {file_data.get('sheet_name', 'N/A')}")
            print(f"Unique ID Column: {file_data.get('unique_id_column', 'N/A')}")
            print(f"Forecast Column: {file_data.get('forecast_column', 'N/A')}")

            if file_data.get('extra_columns'):
                print(f"Extra Columns: {', '.join(file_data['extra_columns'])}")
            else:
                print("Extra Columns: None")
import os
from load_config import load_config
from pathlib import Path
from combine_raw_data import CombineRawData
from highlight_switch_points import HighlightSwitchPoints
from extract_switch_events import ExtractSwitchEvents
from highlight_registry import HighlightRegistry
from extract_resgistry import export_registry_in_excel

def ask_model_number() -> int:
    while True:
        try:
            user_input = int(input("Enter TMA model number (3–8): "))
            if 3 <= user_input <= 8:
                return user_input
            print("Model number must be between 3 and 8.")
        except ValueError:
            print("Please enter a valid number.")

model_num = ask_model_number()
MODEL = f"TMA{model_num}"

# Path to this script
script_dir = Path(__file__).resolve().parent

config = load_config(model_num)
# Relative path to data
data_folder = os.path.normpath(os.path.join(
    script_dir, "..", "..", "Temperature_Performance", "TMA DAQ", MODEL
))

# List DAQ folders
folders = [f for f in os.listdir(data_folder) if f.startswith("DAQ_")]

print(f"Found folders: {folders}")

for folder in folders:
    folder_path = os.path.join(data_folder, folder)

    # Combine CSVs
    combiner = CombineRawData(folder_path, config)
    combined_file = combiner.combine_csvs()

    # --- NEW: Initialize the Registry for THIS folder ---
    registry = HighlightRegistry()

    # Highlight switch points (passing the registry)
    highlightSwitchPoints = HighlightSwitchPoints(combined_file, config, registry)
    highlightSwitchPoints.highlight_switch_points()

    # Extract switch events (passing the registry)
    extractor = ExtractSwitchEvents(combined_file, config, registry)
    # Note: No need to pass green_rows/yellow_rows anymore! The registry has them.
    extractor.create_switch_events_sheet()

    # Extract switch events
    extractor = ExtractSwitchEvents(combined_file, config, registry)
    extractor.create_switch_events_sheet()

    # Export registry for inspection in the same Excel file
    export_registry_in_excel(combined_file, registry)

print("Pipeline complete!")

"""
import os
from load_config import load_config
from pathlib import Path
from combine_raw_data import CombineRawData
from highlight_switch_points import HighlightSwitchPoints
from extract_switch_events import ExtractSwitchEvents

def ask_model_number() -> int:
    while True:
        try:
            user_input = int(input("Enter TMA model number (3–8): "))
            if 3 <= user_input <= 8:
                return user_input
            print("Model number must be between 3 and 8.")
        except ValueError:
            print("Please enter a valid number.")

model_num = ask_model_number()

MODEL = f"TMA{model_num}"

# Path to this script
script_dir = Path(__file__).resolve().parent

config = load_config(model_num)
# Relative path to data
data_folder = os.path.normpath(os.path.join(
    script_dir, "..", "..", "Temperature_Performance", "TMA DAQ", MODEL
))

# List DAQ folders
folders = [f for f in os.listdir(data_folder) if f.startswith("DAQ_")]

print(f"Found folders: {folders}")

for folder in folders:
    folder_path = os.path.join(data_folder, folder)

    # Combine CSVs
    combiner = CombineRawData(folder_path, config)
    combined_file = combiner.combine_csvs()

    # Highlight switch points
    highlightSwitchPoints = HighlightSwitchPoints(combined_file, config)
    green_rows, yellow_rows = highlightSwitchPoints.highlight_switch_points()

    # Extract switch events
    extractor = ExtractSwitchEvents(combined_file, config)
    extractor.create_switch_events_sheet(green_rows, yellow_rows)

print("Pipeline complete!")
"""
"""
# Path to main config file
main_config_path = Path(script_dir) / "main_config.json"

# Path to model config file
json_path = Path(script_dir) / ".." / ".." / "Temperature_Performance" / "TMA DAQ" / MODEL / f"tma{model_num}_config.json"
json_path = json_path.resolve()  # get absolute path

config = json.loads(main_config_path.read_text()) | json.loads(json_path.read_text())
print(config)
"""
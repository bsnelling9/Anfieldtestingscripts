import os
from load_config import load_config
from pathlib import Path
from combine_raw_data import CombineRawData
from highlight_switch_points import HighlightSwitchPoints
from extract_switch_events import ExtractSwitchEvents
from highlight_registry import HighlightRegistry
import time
#from extract_resgistry import export_registry_in_excel

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
    start_processing = time.time()

   
    highlightSwitchPoints = HighlightSwitchPoints(combined_file, config, registry)
    highlightSwitchPoints.highlight_switch_points()

    end_processing = time.time()
    print(f"Processing rows took {end_processing - start_processing:.2f} seconds")
    # Extract switch events (passing the registry)
    
    start_processing = time.time()
    extractor = ExtractSwitchEvents(combined_file, config, registry)
    extractor.create_switch_events_sheet()
    
    end_processing = time.time()
    print(f"ExtractSwitchEvents took {end_processing - start_processing:.2f} seconds")

    # Export registry for inspection in the same Excel file
    #export_registry_in_excel(combined_file, registry)

print("Pipeline complete!")

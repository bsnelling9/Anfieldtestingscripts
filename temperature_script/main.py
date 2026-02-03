import time
from load_config import load_config
from pathlib import Path
from combine_raw_data import CombineRawData
from highlight_switch_points import HighlightSwitchPoints
from extract_switch_events import ExtractSwitchEvents
from highlight_registry import HighlightRegistry
#extract registry is on only used for testing
#from extract_registry import export_registry_in_excel

def ask_model_number() -> int:
    
    while True:
        
        try:
            user_input = int(input("Enter TMA model number (3–8): "))
            if 3 <= user_input <= 8:
                return user_input
            print("Model number must be between 3 and 8.")
        
        except ValueError:
            print("Please enter a valid number.")

def main():
    
    model_num = ask_model_number()
    MODEL = f"TMA{model_num}"

    # Path to this script
    script_dir = Path(__file__).resolve().parent

    data_folder = (script_dir / ".." / ".." / "Temperature_Performance" / "TMA DAQ" / MODEL).resolve()

    config = load_config(model_num)
    
    # List DAQ folders
    folders = [folder for folder in data_folder.iterdir() if folder.name.startswith("DAQ_")]

    print("Found folders:")
    for i, folder in enumerate(folders, start=1):
        print(f"{i}. {folder.name}")

    for folder in folders:

        combiner = CombineRawData(folder, config)
        combined_file = combiner.combine_csvs()

        registry = HighlightRegistry()

        start_processing = time.time()
    
        highlightSwitchPoints = HighlightSwitchPoints(combined_file, config, registry)
        highlightSwitchPoints.highlight_switch_points()
        
        end_processing = time.time()
        print(f"Processing rows for {folder.name} took: {end_processing - start_processing:.2f} seconds")
        
        start_processing = time.time()
        extractor = ExtractSwitchEvents(combined_file, config, registry)
        extractor.create_switch_events_sheet()
        
        end_processing = time.time()
        print(f"Calculating differential for {folder.name} took: {end_processing - start_processing:.2f} seconds")

        # Export registry for inspection in the same Excel file
        #export_registry_in_excel(combined_file, registry)

    print("Data processing complete!")

if __name__ == "__main__":
    main()
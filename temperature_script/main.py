import time
from load_config import load_config, get_model_folder, list_models, PRODUCT_FOLDERS
from combine_raw_data import CombineRawData
from highlight_switch_points import HighlightSwitchPoints
from extract_switch_events import ExtractSwitchEvents
from highlight_registry import HighlightRegistry


def ask_product() -> str:

    options = list(PRODUCT_FOLDERS)

    while True:
        user_input = input(f"Enter product type ({' / '.join(options)}): ").strip().upper()
        if user_input in options:
            return user_input
        print(f"Please enter one of: {', '.join(options)}")


def ask_model_number(product: str) -> int:

    available = list_models(product)
    if not available:
        raise FileNotFoundError(
            f"No {product} model folders found in {get_model_folder(product, 0).parent}"
        )

    while True:

        try:
            user_input = int(input(f"Enter {product} model number {available}: "))
            if user_input in available:
                return user_input
            print(f"No {product}{user_input} folder. Available: {available}")

        except ValueError:
            print("Please enter a valid number.")


def main():

    product = ask_product()
    model_num = ask_model_number(product)

    data_folder = get_model_folder(product, model_num).resolve()
    print(f"Looking in: {data_folder}")

    config = load_config(product, model_num)

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

    print("Data processing complete!")

if __name__ == "__main__":
    main()
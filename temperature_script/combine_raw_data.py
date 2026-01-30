import os
import pandas as pd
from typing import Optional

class CombineRawData:
    """
    Combines Analog and Digital CSVs into a single Excel sheet,
    adds Pressure column, and maintains spacing.
    """
    def __init__(self, folder_path: str, config: dict):
        self.folder_path = folder_path
        self.config = config
        self.resistor: float = config.resistor
        self.pressure: float = config.pressure
        self.output_min: float = config.outputMin
        self.output_max: float = config.outputMax
        self.digital_start_col: int = config.digitalStartCol
        self.daq_meta_data: int = config.daqMetaData

    def compute_pressure(self, voltage: float) -> float:
        """
        Converts T200 output to pressure in psi
        Excel Formula: max((((C/270)-0.004)/(0.016/3000)), 0)
        Need to do one for voltage, have an error that only checks for current
        """
        outputRange = self.output_max - self.output_min
        pressure = ((voltage / self.resistor) - self.output_min) / (outputRange / self.pressure)
        return abs(max(pressure, 0))

    def combine_csvs(self, output_file: Optional[str] = None) -> str:
        
        analog_file = None
        digital_file = None

        # Look in the folder to find the Analog and Digital csv files
        for folder in os.listdir(self.folder_path):
            if folder.startswith("Analog") and folder.endswith(".csv"):
                analog_file = os.path.join(self.folder_path, folder)
            elif folder.startswith("Digital") and folder.endswith(".csv"):
                digital_file = os.path.join(self.folder_path, folder)

        if not analog_file or not digital_file:
            raise FileNotFoundError("Missing Analog or Digital CSV file.")

        skip_rows = self.daq_meta_data
        
        # Read CSVs and skip DAQ metadata
        df_analog = pd.read_csv(analog_file, skiprows=skip_rows)
        df_digital = pd.read_csv(digital_file, skiprows=skip_rows)
        
        # Remove row 2 from digital to match analog
        if len(df_digital) >= 1:
            df_digital = df_digital.drop(index=0).reset_index(drop=True)
        
        # removes repeated columns in the digital csv
        if df_digital.shape[1] > 2:
            df_digital = df_digital.iloc[:, 2:]

        #Calculate pressure using iloc[:, 2] which selects the 3rd column
        df_pressure = df_analog.iloc[:, 2].apply(self.compute_pressure).to_frame("Pressure (psi)")

        df_combined = pd.concat([df_analog, df_pressure, df_digital], axis=1)

        # Determine output file
        if output_file is None:
            
            folder_name = os.path.basename(self.folder_path)
            # Remove "DAQ_" prefix for naming
            folder_name = folder_name.replace("DAQ_", "")
            output_file = os.path.join(os.path.dirname(self.folder_path), f"{folder_name}_Processed.xlsx")

        df_combined.to_excel(output_file, index=False)
        
        return output_file

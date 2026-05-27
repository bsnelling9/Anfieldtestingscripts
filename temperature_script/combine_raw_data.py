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
        self.daq_meta_data: int = config.daqMetaData

    def calculate_pressure(self, voltage: float) -> float:
        """
        Converts T200 output to pressure in psi
        Excel Formula: max((((C/270)-0.004)/(0.016/3000)), 0)
        Need to do one for voltage, have an error that only checks for current
        """
        outputRange = self.output_max - self.output_min
        pressure = ((voltage / self.resistor) - self.output_min) / (outputRange / self.pressure)
        
        return max(pressure, 0)

    def combine_csvs(self, output_file: Optional[str] = None) -> str:
        
        analog_csv = None
        digital_csv = None

        for file in self.folder_path.iterdir():
            if file.is_file() and file.name.startswith("Analog") and file.suffix == ".csv":
                analog_csv = file
            elif file.is_file() and file.name.startswith("Digital") and file.suffix == ".csv":
                digital_csv = file

        if not analog_csv or not digital_csv:
            raise FileNotFoundError("Missing Analog or Digital CSV file.")

        skip_rows = self.daq_meta_data
        
        df_analog = pd.read_csv(analog_csv, skiprows=skip_rows)
        df_digital = pd.read_csv(digital_csv, skiprows=skip_rows)
        
        if len(df_digital) >= 1:
            df_digital = df_digital.drop(index=0).reset_index(drop=True)
        
        if df_digital.shape[1] > 2:
            df_digital = df_digital.iloc[:, 2:]

        df_pressure = df_analog.iloc[:, 2].apply(self.calculate_pressure).to_frame("Pressure (psi)")

        df_combined = pd.concat([df_analog, df_pressure, df_digital], axis=1)

        if output_file is None:
            
            file_name = self.folder_path.name.replace("DAQ_", "")
            output_file = (
                self.folder_path.parent
                / f"{file_name}_Processed.xlsx"
            )

        df_combined.to_excel(output_file, index=False)
        
        return output_file

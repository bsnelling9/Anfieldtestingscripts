import pandas as pd
from pathlib import Path
from typing import Optional


class CombineRawData:

    SIGNAL_KEYS = ("outputType", "resistor", "outputMin", "outputMax", "pressure")

    def __init__(self, folder_path, config):
        self.folder_path = Path(folder_path)
        self.config = config
        self.daq_meta_data: int = config.daqMetaData

    def resolve_transducer(self, analog_header: str) -> dict:

        settings = {key: getattr(self.config, key) for key in self.SIGNAL_KEYS}

        profile = self.config.transducers.get(analog_header)
        if profile:
            settings.update({k: v for k, v in profile.items() if k in self.SIGNAL_KEYS})

        self.validate_settings(analog_header, settings)
        
        return settings

    @staticmethod
    def validate_settings(analog_header: str, s: dict):
        output_type = str(s["outputType"]).lower()
        where = f"transducer '{analog_header}'"

        if output_type not in ("current", "voltage"):
            raise ValueError(f"{where}: outputType must be 'Current' or 'Voltage', got '{s['outputType']}'")

        for key in ("outputMin", "outputMax", "pressure"):
            if s[key] is None:
                raise ValueError(f"{where}: '{key}' is not set. Fill it in from the transducer datasheet in the config.")

        if s["outputMax"] <= s["outputMin"]:
            raise ValueError(f"{where}: outputMax ({s['outputMax']}) must be greater than outputMin ({s['outputMin']})")

        if output_type == "current" and not s["resistor"]:
            raise ValueError(f"{where}: a current-output transducer needs 'resistor' (shunt resistance in ohms)")


        if output_type == "voltage" and s["outputMax"] < 0.1:
            raise ValueError(
                f"{where}: outputType is Voltage but outputMax is {s['outputMax']}. "
                "That looks like a current range in amps; set outputMin/outputMax in volts."
            )

    @staticmethod
    def calculate_pressure(readings: pd.Series, s: dict) -> pd.Series:

        if str(s["outputType"]).lower() == "current":
            signal = readings / s["resistor"]
        else:
            signal = readings

        output_range = s["outputMax"] - s["outputMin"]
        pressure = (signal - s["outputMin"]) / output_range * s["pressure"]

        return pressure.clip(lower=0)

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

        analog_header = df_analog.columns[2]
        settings = self.resolve_transducer(analog_header)
        print(
            f"  {self.folder_path.name}: '{analog_header}' -> {settings['outputType']} output, "
            f"{settings['outputMin']}-{settings['outputMax']}, {settings['pressure']} psi full scale"
        )

        df_pressure = self.calculate_pressure(df_analog.iloc[:, 2], settings).to_frame("Pressure (psi)")

        df_combined = pd.concat([df_analog, df_pressure, df_digital], axis=1)

        if output_file is None:

            file_name = self.folder_path.name.replace("DAQ_", "")
            output_file = (
                self.folder_path.parent
                / f"{file_name}_Processed.xlsx"
            )

        df_combined.to_excel(output_file, index=False)

        return output_file
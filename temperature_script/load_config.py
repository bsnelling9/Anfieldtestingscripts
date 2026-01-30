import json
from pathlib import Path
from config_types import TMAConfig, HighlightColors


def load_config(model_num: int) -> TMAConfig:
    model = f"TMA{model_num}"

    script_dir = Path(__file__).resolve().parent

    default_path = script_dir / "main_config.json"
    model_path = (
        script_dir
        / ".."
        / ".."
        / "Temperature_Performance"
        / "TMA DAQ"
        / model
        / f"tma{model_num}_config.json"
    ).resolve()

    if not default_path.exists():
        raise FileNotFoundError(f"Default config not found: {default_path}")

    if not model_path.exists():
        raise FileNotFoundError(f"Model config not found: {model_path}")

    defaults = json.loads(default_path.read_text())
    overrides = json.loads(model_path.read_text())

    # Shallow merge
    merged = {**defaults, **overrides}

    # Nested merge for highlightColors
    merged["highlightColors"] = {
        **defaults.get("highlightColors", {}),
        **overrides.get("highlightColors", {}),
    }

    return TMAConfig(
        model=merged["model"],
        transducer=merged["transducer"],
        pressure=merged["pressure"],

        outputType=merged["outputType"],
        resistor=merged["resistor"],
        outputMin=merged["outputMin"],
        outputMax=merged["outputMax"],

        digitalStartCol=merged["digitalStartCol"],
        daqMetaData=merged["daqMetaData"],
        pressureCol=merged["pressureCol"],

        highlightColors=HighlightColors(**merged["highlightColors"]),
        protectedHeaders=merged["protectedHeaders"],
    )

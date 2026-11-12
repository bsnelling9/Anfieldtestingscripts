import json
from pathlib import Path
import re
from config_types import TMAConfig, HighlightColors

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_ROOT = (SCRIPT_DIR / ".." / ".." / "Temperature_Performance").resolve()

PRODUCT_FOLDERS = {
    "TMA": "TMA DAQ",
    "TPA": "TPA DAQ",
}


def get_model_folder(product: str, model_num: int) -> Path:
    return DATA_ROOT / PRODUCT_FOLDERS[product] / f"{product}{model_num}"


def list_models(product: str) -> list[int]:

    product_dir = DATA_ROOT / PRODUCT_FOLDERS[product]
    if not product_dir.exists():
        return []
    pattern = re.compile(rf"^{product}(\d+)$", re.IGNORECASE)
    models = []
    for folder in product_dir.iterdir():
        match = pattern.match(folder.name)
        if folder.is_dir() and match:
            models.append(int(match.group(1)))
    
    return sorted(models)


def merge_dicts(defaults: dict, overrides: dict) -> dict:
    
    merged = {**defaults, **overrides}
    merged["highlightColors"] = {**defaults.get("highlightColors", {}), **overrides.get("highlightColors", {})}

    default_tx = defaults.get("transducers", {})
    override_tx = overrides.get("transducers", {})
    merged["transducers"] = {
        header: {**default_tx.get(header, {}), **override_tx.get(header, {})}
        for header in {**default_tx, **override_tx}
    }
    
    return merged

def load_config(product: str, model_num: int) -> TMAConfig:

    default_path = SCRIPT_DIR / "main_config.json"
    model_path = get_model_folder(product, model_num) / f"{product.lower()}{model_num}_config.json"

    if not default_path.exists():
        raise FileNotFoundError(f"Default config not found: {default_path}")

    if not model_path.exists():
        raise FileNotFoundError(f"Model config not found: {model_path}")

    defaults = json.loads(default_path.read_text())
    overrides = json.loads(model_path.read_text())
    merged = merge_dicts(defaults, overrides)

    return TMAConfig(
        model=merged["model"],
        transducer=merged["transducer"],
        pressure=merged["pressure"],

        outputType=merged["outputType"],
        resistor=merged.get("resistor"),
        outputMin=merged["outputMin"],
        outputMax=merged["outputMax"],

        digitalStartCol=merged["digitalStartCol"],
        daqMetaData=merged["daqMetaData"],
        pressureCol=merged["pressureCol"],

        highlightColors=HighlightColors(**merged["highlightColors"]),
        protectedHeaders=merged["protectedHeaders"],
        transducers=merged["transducers"],
    )
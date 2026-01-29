import json
from pathlib import Path
from config import config

def load_config(model_num: int) -> dict:
    """
    Load default config, then override with model-specific JSON.
    """

    model = f"TMA{model_num}"

    script_dir = Path(__file__).resolve().parent

    json_path = (
        script_dir
        / ".."
        / ".."
        / "Temperature_Performance"
        / "TMA DAQ"
        / model
        / f"tma{model_num}_config.json"
    ).resolve()

    config = DEFAULT_CONFIG.copy()

    if not json_path.exists():
        raise FileNotFoundError(f"Config file not found: {json_path}")

    with open(json_path, "r") as f:
        model_config = json.load(f)

    config.update(model_config)  # model overrides defaults
    return config
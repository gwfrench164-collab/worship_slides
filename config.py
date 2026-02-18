print("USING CONFIG FILE:", __file__)
from pathlib import Path
import json

CONFIG_FILE = Path.home() / ".worship_slides_config.json"

REQUIRED_FOLDERS = ["songs", "templates", "output", "SongPDFs", "notes_refs"]

def _load_config():
    if not CONFIG_FILE.exists():
        return {}
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def _save_config(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def load_data_root():
    return _load_config().get("data_root")

def save_data_root(path):
    print(">>> SAVE_DATA_ROOT EXECUTED")
    data = _load_config()
    data["data_root"] = str(path)
    _save_config(data)

def load_build_prefs():
    cfg = _load_config()
    return {
        "last_template": cfg.get("last_template"),
        "last_output": cfg.get("last_output"),
    }

def save_build_prefs(template_name, output_name):
    print(">>> SAVE_BUILD_PREFS EXECUTED")
    data = _load_config()
    data["last_template"] = template_name
    data["last_output"] = output_name
    _save_config(data)

def ensure_data_root_structure(data_root: str | None) -> None:
    """
    Ensures the standard folder structure exists inside data_root.
    Safe to call at startup every time.
    """
    if not data_root:
        return

    root = Path(data_root)
    root.mkdir(parents=True, exist_ok=True)

    for name in REQUIRED_FOLDERS:
        (root / name).mkdir(parents=True, exist_ok=True)

from pathlib import Path

def load_bible_json_path():
    data = _load_config()
    return data.get("bible_json_path", "")

def save_bible_json_path(path):
    data = _load_config()
    data["bible_json_path"] = str(path)
    _save_config(data)

def auto_find_kjv_json():
    here = Path(__file__).resolve().parent
    candidate = here / "kjv.json"
    if candidate.exists():
        return str(candidate)

    data_root = load_data_root()
    if data_root:
        candidate2 = Path(data_root) / "kjv.json"
        if candidate2.exists():
            return str(candidate2)

    return
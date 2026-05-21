"""Settings management using config.json."""
import json
from pathlib import Path
from .paths import PROJECT_ROOT

CONFIG_PATH = PROJECT_ROOT / "config.json"

_DEFAULTS = {
    "nb_kb_path": "",
    "dt_kb_path": "",
    "peripheral_path": "",
    "output_nb_path": "",
    "output_dt_path": "",
    "output_peripheral_path": "",
    "output_consolidate_all_path": "",
}


def load() -> dict:
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return {**_DEFAULTS, **data}
        except Exception:
            pass
    return dict(_DEFAULTS)


def save(cfg: dict) -> None:
    tmp = CONFIG_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(CONFIG_PATH)

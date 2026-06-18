import os
import json
import sys
from typing import Any

CONFIG_FILE = "config.json"


def _ruta_config() -> str:
    if getattr(sys, "frozen", False):
        base = os.environ.get("APPDATA", "")
        return os.path.join(base, "GeneradorDocumentos", CONFIG_FILE)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE)


def cargar() -> dict[str, Any]:
    ruta = _ruta_config()
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def guardar(config: dict[str, Any]) -> None:
    ruta = _ruta_config()
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

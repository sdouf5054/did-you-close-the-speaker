"""
Configuration loader.
Reads config.json from the project root directory.
"""

import json
import sys
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.json"

REQUIRED_KEYS = ["tapo_email", "tapo_password", "plug_ip"]

DEFAULTS = {
    "delay_after_power_off_sec": 2,
    "timeout_sec": 5,
}


def load_config() -> dict:
    """Load and validate config.json."""
    if not CONFIG_PATH.exists():
        print(f"❌ Config file not found: {CONFIG_PATH}")
        print("   Copy config.example.json to config.json and fill in your details.")
        sys.exit(1)

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    missing = [k for k in REQUIRED_KEYS if not cfg.get(k)]
    if missing:
        print(f"❌ Missing required config keys: {', '.join(missing)}")
        sys.exit(1)

    for key, default in DEFAULTS.items():
        cfg.setdefault(key, default)

    return cfg

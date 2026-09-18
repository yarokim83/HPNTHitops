"""Atomic preferences, separate from ERP credentials."""
import json
import os
from pathlib import Path
import tempfile

VERSION = '1.1.1'
DATA_DIR = Path(os.getenv('LOCALAPPDATA') or Path.home()) / 'PRMaker'
INSTALL_EXE = DATA_DIR / 'app' / 'PRMakerWidget.exe'


def load():
    try:
        data = json.loads((DATA_DIR / 'preferences.json').read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save(data):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    handle, path = tempfile.mkstemp(dir=DATA_DIR, suffix='.tmp')
    try:
        with os.fdopen(handle, 'w', encoding='utf-8') as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
        os.replace(path, DATA_DIR / 'preferences.json')
    finally:
        if os.path.exists(path):
            os.unlink(path)

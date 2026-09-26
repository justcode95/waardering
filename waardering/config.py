"""Instellingen (API-sleutel, model, template) in %APPDATA%\\Waardering\\config.json."""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from .uitlezen import MODEL


def programmamap() -> Path:
    """Map met het meegeleverde template (werkt ook in de .exe van PyInstaller)."""
    if getattr(sys, 'frozen', False):
        return Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent))
    return Path(__file__).resolve().parent.parent


def configpad() -> Path:
    basis = Path(os.environ.get('APPDATA') or Path.home() / '.config')
    pad = basis / 'Waardering'
    pad.mkdir(parents=True, exist_ok=True)
    return pad / 'config.json'


def laad() -> dict:
    standaard = {'api_key': '', 'model': MODEL,
                 'template': str(programmamap() / 'template' / 'Waardering_template.xlsx')}
    p = configpad()
    if p.exists():
        standaard.update(json.loads(p.read_text(encoding='utf-8')))
    if not Path(standaard['template']).exists():
        standaard['template'] = str(programmamap() / 'template' / 'Waardering_template.xlsx')
    return standaard


def bewaar(cfg: dict) -> None:
    configpad().write_text(json.dumps(cfg, indent=1), encoding='utf-8')

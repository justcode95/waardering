"""Invoer uit het formulier omzetten naar getallen."""
from __future__ import annotations

import re


def getal(s: str) -> float | None:
    """Leest Belgische en gewone notatie: 75.000 / 75000 / 4,5 / 1.253,20 / 0.12."""
    s = s.strip().replace('€', '').replace(' ', '').replace('%', '')
    if not s:
        return None
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    elif re.fullmatch(r'-?\d{1,3}(\.\d{3})+', s):
        s = s.replace('.', '')
    return float(s)

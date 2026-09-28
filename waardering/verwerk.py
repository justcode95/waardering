"""De twee stappen van het programma, los van het venster: dossier uitlezen en waardering maken."""
from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Callable

from . import herberekenen, rapport
from .invullen import Invuller
from .model import Beslissingen, Dossier, bewaar_json, dossier_naar_json
from .samenstellen import samenstellen, signalen, voorstel
from .uitlezen import Uitlezer

CACHE = '.waardering_cache'
KEUZES = 'waardering_keuzes.json'


def pdfs_in(map_: Path) -> list[Path]:
    return sorted(p for p in map_.rglob('*.pdf') if CACHE not in p.parts)


def lees_dossier(map_: Path, cfg: dict, log: Callable[[str], None]) -> tuple[Dossier, Beslissingen, float]:
    bestanden = pdfs_in(map_)
    if not bestanden:
        raise ValueError(f'Geen PDF-bestanden gevonden in {map_}')
    lezer = Uitlezer(cfg.get('api_key') or None, cfg.get('model') or 'claude-opus-5', log)
    resultaten = []
    for pdf in bestanden:
        r = lezer.lees(pdf, map_ / CACHE)
        log(f'   → {r["documenttype"]}, {len(r["periodes"])} periode(s), {len(r["activa"])} activa'
            + (f', {len(r["_controle"])} rubriek(en) sluiten niet' if r['_controle'] else ''))
        resultaten.append(r)
    d = samenstellen(resultaten)
    (map_ / CACHE).mkdir(exist_ok=True)
    bewaar_json(map_ / CACHE / 'dossier.json', dossier_naar_json(d))
    b = voorstel(d)
    oude = map_ / KEUZES
    if oude.exists():
        import json
        try:
            vorige = Beslissingen.van_json(json.loads(oude.read_text(encoding='utf-8')))
            if vorige.afsluitdatum in [p.einde for p in d.periodes]:
                for k, datum in vorige.kolommen.items():   # periode intussen anders gedateerd: voorstel nemen
                    if datum and d.periode_op(datum) is None:
                        vorige.kolommen[k] = b.kolommen.get(k)
                        if not vorige.kolommen[k]:
                            vorige.weging[k] = 0
                b = vorige
                log('Keuzes van een vorige run geladen.')
        except (KeyError, ValueError, TypeError):
            pass
    log(f'Kost uitlezen: ± ${lezer.kost:.2f}')
    return d, b, lezer.kost


def bestandsnaam(d: Dossier, b: Beslissingen) -> str:
    naam = re.sub(r'[\\/:*?"<>|]+', '', d.naam or 'Waardering').strip()
    return f'{naam} - Waardebepaling - {b.afsluitdatum.strftime("%d-%m-%Y")}.xlsx'


def maak_waardering(map_: Path, d: Dossier, b: Beslissingen, cfg: dict, log: Callable[[str], None],
                    kost: float = 0.0) -> tuple[Path, dict, list, list]:
    bewaar_json(map_ / KEUZES, b.naar_json())
    uit = map_ / bestandsnaam(d, b)
    if uit.exists():
        uit = uit.with_name(uit.stem + ' (nieuw).xlsx')
    werk = Path(tempfile.mkdtemp(prefix='waardering_'))
    try:
        log('Template invullen …')
        inv = Invuller(Path(cfg['template']), d, b, werk / 'x')
        inv.vul(uit)
        log('Herberekenen en controleren …')
        waarden = None
        if herberekenen.met_excel(uit):
            waarden = herberekenen.waarden_openpyxl(uit)
        else:
            waarden = herberekenen.met_libreoffice(uit)
            if waarden:                           # berekende waarden in het bestand bewaren
                inv2 = Invuller(Path(cfg['template']), d, b, werk / 'y')
                inv2.vul(uit, cached=waarden)
        ok, fout, res = [], [], {}
        if waarden:
            afsl = d.periode_op(b.afsluitdatum)
            ok, fout = herberekenen.controles(waarden, b.kolommen, d.afschrijvingstabel_totaal is not None,
                                              afsl.balanstotaal if afsl else None)
            res = herberekenen.resultaten(waarden)
        else:
            log('Geen Excel of LibreOffice gevonden: controles overgeslagen (Excel rekent bij het openen).')
        meldingen = d.meldingen + inv.meldingen
        if not b.multiple:
            meldingen.append('Multiple (Marktwaarde D19) niet ingevuld: de marktwaarde is onvolledig.')
        rapport.schrijf(uit.with_suffix('.md'), d, b, ok, fout, res, meldingen, signalen(d, b), kost)
        for x in ok:
            log(f'  OK  {x}')
        for x in fout:
            log(f'  FOUT {x}')
        log(f'Klaar: {uit.name}')
        return uit, res, ok, fout
    finally:
        shutil.rmtree(werk, ignore_errors=True)

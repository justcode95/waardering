"""PDF's uitlezen met Claude (ook gescande bundels en rapporten van elk boekhoudkantoor).

Per PDF één verzoek met gestructureerde JSON-uitvoer. Het resultaat wordt gecontroleerd (elke rubriek moet
optellen tot haar totaal); bij afwijkingen volgt één correctieronde in hetzelfde gesprek. Resultaten worden
per bestand (hash) bewaard in <dossiermap>/.waardering_cache, zodat een tweede run niets kost.
"""
from __future__ import annotations

import base64
import hashlib
import json
from pathlib import Path
from typing import Callable

import anthropic

MODEL = 'claude-opus-5'
PRIJS_PER_MTOK = {'claude-opus-5': (5.0, 25.0), 'claude-opus-5-5': (4.0, 20.0), 'claude-sonnet-5': (2.0, 10.0)}

_REK = {
    'type': 'object',
    'properties': {'nummer': {'type': 'string'}, 'omschrijving': {'type': 'string'}, 'saldo': {'type': 'number'}},
    'required': ['nummer', 'omschrijving', 'saldo'], 'additionalProperties': False}
_RUBRIEK = {
    'type': 'object',
    'properties': {'naam': {'type': 'string'}, 'totaal': {'type': 'number'},
                   'rekeningen': {'type': 'array', 'items': _REK}},
    'required': ['naam', 'totaal', 'rekeningen'], 'additionalProperties': False}
_PERIODE = {
    'type': 'object',
    'properties': {
        'periode_begin': {'type': 'string', 'description': 'JJJJ-MM-DD of leeg'},
        'periode_einde': {'type': 'string', 'description': 'JJJJ-MM-DD'},
        'tussentijds': {'type': 'boolean'},
        'rubrieken': {'type': 'array', 'items': _RUBRIEK},
        'bedrijfswinst': {'type': 'number'},
        'winst_voor_belasting': {'type': 'number'},
        'balanstotaal': {'type': 'number'}},
    'required': ['periode_begin', 'periode_einde', 'tussentijds', 'rubrieken', 'bedrijfswinst',
                 'winst_voor_belasting', 'balanstotaal'],
    'additionalProperties': False}
_ACTIEF = {
    'type': 'object',
    'properties': {'rekening': {'type': 'string'}, 'omschrijving': {'type': 'string'},
                   'datum_aanschaf': {'type': 'string', 'description': 'JJJJ-MM-DD'},
                   'aanschafwaarde': {'type': 'number'}, 'netto_boekwaarde': {'type': 'number'}},
    'required': ['rekening', 'omschrijving', 'datum_aanschaf', 'aanschafwaarde', 'netto_boekwaarde'],
    'additionalProperties': False}
_PERSOON = {
    'type': 'object',
    'properties': {'naam': {'type': 'string'}, 'rol': {'type': 'string'}, 'aandelen': {'type': 'number'},
                   'bezoldiging': {'type': 'number'}},
    'required': ['naam', 'rol', 'aandelen', 'bezoldiging'], 'additionalProperties': False}

SCHEMA = {
    'type': 'object',
    'properties': {
        'documenttype': {'type': 'string', 'enum': [
            'jaarrekening_of_fiscale_bundel', 'tussentijdse_cijfers', 'afschrijvingstabel',
            'openstaande_klanten', 'openstaande_leveranciers', 'andere']},
        'vennootschap': {
            'type': 'object',
            'properties': {'naam': {'type': 'string'}, 'straat': {'type': 'string'},
                           'postcode_gemeente': {'type': 'string'}, 'ondernemingsnummer': {'type': 'string'},
                           'activiteit': {'type': 'string'}},
            'required': ['naam', 'straat', 'postcode_gemeente', 'ondernemingsnummer', 'activiteit'],
            'additionalProperties': False},
        'periodes': {'type': 'array', 'items': _PERIODE},
        'activa': {'type': 'array', 'items': _ACTIEF},
        'afschrijvingstabel_totaal_aanschaf': {'type': 'number'},
        'personen': {'type': 'array', 'items': _PERSOON},
        'openstaand_totaal': {'type': 'number'},
        'opmerkingen': {'type': 'string'}},
    'required': ['documenttype', 'vennootschap', 'periodes', 'activa', 'afschrijvingstabel_totaal_aanschaf',
                 'personen', 'openstaand_totaal', 'opmerkingen'],
    'additionalProperties': False}

INSTRUCTIE = """Je leest een document uit het waarderingsdossier van een Belgische vennootschap en zet het om naar het
gevraagde JSON-formaat. Een accountant gebruikt de cijfers rechtstreeks in een waardering, dus volledigheid en
exactheid per rekening zijn belangrijker dan snelheid.

1. documenttype: jaarrekening_of_fiscale_bundel (afgesloten boekjaar, vaak een ingescande bundel met detailbalans,
   bijlagen, notulen en aangifte), tussentijdse_cijfers (voorlopige of tussentijdse balans/resultatenrekening),
   afschrijvingstabel, openstaande_klanten, openstaande_leveranciers of andere.
2. periodes: neem voor elke periode waarvoor het document een DETAIL PER REKENING bevat (rapporteringsbalans,
   detailbalans, proef- en saldibalans) alle rekeningen van balans én resultatenrekening over, gegroepeerd per
   rubriek zoals afgedrukt, met het rubriektotaal zoals afgedrukt. Gebruik het rekeningnummer exact zoals gedrukt
   (bv. 618000, 440, 5500001). Neem geen vergelijkende cijfers van een vorig jaar over als aparte periode tenzij
   ze per rekening gedetailleerd zijn. Neem de verkorte NBB-jaarrekening niet over als er een detail per rekening is.
   Tekens: activa positief (afschrijvingen en waardeverminderingen op activa negatief), eigen vermogen en
   schulden positief, opbrengsten positief, kosten negatief, ongeacht hoe het rapport ze toont. Het rubriektotaal
   volgt dezelfde conventie en moet gelijk zijn aan de som van de rekeningen van die rubriek; controleer dat.
   bedrijfswinst en winst_voor_belasting zoals afgedrukt (0 als niet aanwezig); balanstotaal = totaal activa.
   tussentijds = true als de periode niet een volledig afgesloten boekjaar is. Datums als JJJJ-MM-DD.
3. activa (alleen bij een afschrijvingstabel): één regel per actief met rekening (bv. 230000), omschrijving,
   datum van aanschaf/investering, aanschafwaarde en de laatst vermelde netto boekwaarde (restwaarde).
   afschrijvingstabel_totaal_aanschaf = het algemeen totaal van de aanschafwaarden.
4. vennootschap: naam, straat + nummer, postcode + gemeente, ondernemingsnummer (formaat 0123.456.789) en een
   korte omschrijving van de activiteit (bv. "uitbating van een sauna- en wellnesscentrum"), voor zover het
   document ze vermeldt; anders een lege string.
5. personen: bestuurders en aandeelhouders uit notulen, verslagen of lijsten (rol "bestuurder", "aandeelhouder" of
   beide), met aantal aandelen en totale bezoldiging voor zover vermeld (anders 0).
6. openstaand_totaal: bij een lijst openstaande klanten of leveranciers het totaal; anders 0.
7. opmerkingen: kort wat opvalt en relevant is voor een waardering (bv. geen afschrijvingen geboekt in tussentijdse
   cijfers, uitzonderlijke posten, pagina's die onleesbaar waren). Lege lijsten en 0 voor wat niet van toepassing is.
"""


class Uitlezer:
    def __init__(self, api_key: str | None = None, model: str = MODEL, log: Callable[[str], None] = print):
        self.client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
        self.model = model
        self.log = log
        self.kost = 0.0

    # ------------------------------------------------------------------
    def lees(self, pdf: Path, cachemap: Path) -> dict:
        data = pdf.read_bytes()
        sleutel = hashlib.sha256(data).hexdigest()[:24]
        cachemap.mkdir(exist_ok=True)
        cache = cachemap / f'{sleutel}.json'
        if cache.exists():
            self.log(f'{pdf.name}: uit cache')
            return json.loads(cache.read_text(encoding='utf-8'))
        self.log(f'{pdf.name}: uitlezen met {self.model} …')
        document = {'type': 'document',
                    'source': {'type': 'base64', 'media_type': 'application/pdf',
                               'data': base64.standard_b64encode(data).decode('ascii')},
                    'cache_control': {'type': 'ephemeral'}}
        berichten = [{'role': 'user', 'content': [document, {'type': 'text', 'text': INSTRUCTIE}]}]
        antwoord, resultaat = self._vraag(berichten)
        fouten = controleer(resultaat)
        if fouten:
            self.log(f'{pdf.name}: {len(fouten)} rubriek(en) sluiten niet, correctieronde …')
            berichten += [{'role': 'assistant', 'content': antwoord.content},
                          {'role': 'user', 'content': [{'type': 'text', 'text':
                              'Deze rubrieken tellen niet op tot hun totaal:\n' + '\n'.join(fouten) +
                              '\nLees die pagina\'s opnieuw nauwkeurig (let op mintekens, weggevallen regels en '
                              'verwisselde cijfers) en geef het volledige JSON-resultaat opnieuw, gecorrigeerd.'}]}]
            _, resultaat = self._vraag(berichten)
            fouten = controleer(resultaat)
        resultaat['_controle'] = fouten
        resultaat['_bestand'] = pdf.name
        cache.write_text(json.dumps(resultaat, ensure_ascii=False), encoding='utf-8')
        return resultaat

    def _vraag(self, berichten: list) -> tuple:
        try:
            with self.client.beta.messages.stream(
                    model=self.model,
                    max_tokens=64000,
                    betas=['server-side-fallback-2026-07-01'],
                    fallbacks='default',
                    thinking={'type': 'adaptive'},
                    output_config={'format': {'type': 'json_schema', 'schema': SCHEMA}},
                    messages=berichten) as stream:
                antwoord = stream.get_final_message()
        except anthropic.AuthenticationError:
            raise RuntimeError('De API-sleutel is ongeldig. Controleer ze in Instellingen.')
        except anthropic.RateLimitError:
            raise RuntimeError('Te veel verzoeken tegelijk bij Anthropic. Probeer het over een minuut opnieuw.')
        except anthropic.APIConnectionError:
            raise RuntimeError('Geen verbinding met de Anthropic API. Controleer de internetverbinding.')
        except anthropic.BadRequestError as e:
            raise RuntimeError(f'De API weigerde het verzoek: {e.message}')
        self._boek_kost(antwoord)
        if antwoord.stop_reason == 'refusal':
            raise RuntimeError('Het model weigerde dit document te verwerken.')
        if antwoord.stop_reason == 'max_tokens':
            raise RuntimeError('Het antwoord was te lang (max_tokens bereikt); splits het document op.')
        tekst = next(b.text for b in antwoord.content if b.type == 'text')
        return antwoord, json.loads(tekst)

    def _boek_kost(self, antwoord) -> None:
        u = antwoord.usage
        pin, pout = PRIJS_PER_MTOK.get(self.model, (5.0, 25.0))
        cache_read = getattr(u, 'cache_read_input_tokens', 0) or 0
        cache_write = getattr(u, 'cache_creation_input_tokens', 0) or 0
        kost = (u.input_tokens * pin + cache_write * pin * 1.25 + cache_read * pin * 0.1 + u.output_tokens * pout) / 1e6
        self.kost += kost
        self.log(f'   tokens in {u.input_tokens + cache_read + cache_write:,} / uit {u.output_tokens:,} '
                 f'(± ${kost:.2f})')


def controleer(resultaat: dict) -> list[str]:
    """Rubrieken waarvan de som van de rekeningen niet gelijk is aan het rubriektotaal."""
    fouten = []
    for p in resultaat.get('periodes', []):
        for r in p.get('rubrieken', []):
            som = round(sum(x['saldo'] for x in r['rekeningen']), 2)
            if r['rekeningen'] and abs(som - r['totaal']) > 0.011:
                fouten.append(f"{p['periode_einde']} – {r['naam']}: som {som:,.2f} ≠ totaal {r['totaal']:,.2f}")
    return fouten

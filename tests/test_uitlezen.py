"""Uitlezer tegen een lokale, nagebootste Messages API (streaming), zodat het volledige codepad getest wordt
zonder API-sleutel of kosten: verzoekopbouw, gestructureerde uitvoer, correctieronde en cache."""
import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

from waardering.uitlezen import Uitlezer


def antwoord(rubriek_totaal: float) -> dict:
    return {'documenttype': 'jaarrekening_of_fiscale_bundel',
            'vennootschap': {'naam': 'Testbedrijf BV', 'straat': '', 'postcode_gemeente': '',
                             'ondernemingsnummer': '', 'activiteit': ''},
            'periodes': [{'periode_begin': '2025-01-01', 'periode_einde': '2025-12-31', 'tussentijds': False,
                          'bedrijfswinst': 0, 'winst_voor_belasting': 0, 'balanstotaal': 0,
                          'rubrieken': [{'naam': 'Omzet', 'totaal': rubriek_totaal,
                                         'rekeningen': [{'nummer': '700000', 'omschrijving': 'Verkopen',
                                                         'saldo': 1000}]}]}],
            'activa': [], 'afschrijvingstabel_totaal_aanschaf': 0, 'personen': [], 'openstaand_totaal': 0,
            'opmerkingen': ''}


class NepAPI(BaseHTTPRequestHandler):
    verzoeken: list = []
    antwoorden: list = []

    def log_message(self, *a):
        pass

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers['content-length'])))
        NepAPI.verzoeken.append({'pad': self.path, 'headers': dict(self.headers), 'body': body})
        tekst = json.dumps(NepAPI.antwoorden.pop(0))
        events = [
            ('message_start', {'type': 'message_start', 'message': {
                'id': 'msg_1', 'type': 'message', 'role': 'assistant', 'model': body['model'], 'content': [],
                'stop_reason': None, 'stop_sequence': None,
                'usage': {'input_tokens': 1000, 'output_tokens': 1}}}),
            ('content_block_start', {'type': 'content_block_start', 'index': 0,
                                     'content_block': {'type': 'text', 'text': ''}}),
            ('content_block_delta', {'type': 'content_block_delta', 'index': 0,
                                     'delta': {'type': 'text_delta', 'text': tekst}}),
            ('content_block_stop', {'type': 'content_block_stop', 'index': 0}),
            ('message_delta', {'type': 'message_delta', 'delta': {'stop_reason': 'end_turn', 'stop_sequence': None},
                               'usage': {'output_tokens': 200}}),
            ('message_stop', {'type': 'message_stop'})]
        self.send_response(200)
        self.send_header('content-type', 'text/event-stream')
        self.end_headers()
        for naam, data in events:
            self.wfile.write(f'event: {naam}\ndata: {json.dumps(data)}\n\n'.encode())


@pytest.fixture
def api(monkeypatch):
    server = HTTPServer(('127.0.0.1', 0), NepAPI)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    monkeypatch.setenv('ANTHROPIC_BASE_URL', f'http://127.0.0.1:{server.server_port}')
    NepAPI.verzoeken, NepAPI.antwoorden = [], []
    yield NepAPI
    server.shutdown()


def test_uitlezen_met_correctieronde_en_cache(api, tmp_path):
    api.antwoorden = [antwoord(999), antwoord(1000)]       # eerst een rubriek die niet sluit, dan gecorrigeerd
    pdf = tmp_path / 'jaarrekening.pdf'
    pdf.write_bytes(b'%PDF-1.4 test')
    lezer = Uitlezer(api_key='test', log=lambda *_: None)
    r = lezer.lees(pdf, tmp_path / 'cache')
    assert r['_controle'] == [] and r['periodes'][0]['rubrieken'][0]['totaal'] == 1000
    assert len(api.verzoeken) == 2
    eerste, tweede = api.verzoeken
    assert eerste['pad'].startswith('/v1/messages')
    assert 'server-side-fallback-2026-07-01' in eerste['headers'].get('anthropic-beta', '')
    b = eerste['body']
    assert b['model'] == 'claude-opus-5' and b['fallbacks'] == 'default' and b['stream'] is True
    assert b['output_config']['format']['type'] == 'json_schema'
    assert b['messages'][0]['content'][0]['type'] == 'document'
    assert [m['role'] for m in tweede['body']['messages']] == ['user', 'assistant', 'user']
    assert 'niet op tot hun totaal' in tweede['body']['messages'][2]['content'][0]['text']
    assert lezer.kost > 0
    # tweede keer: uit de cache, geen nieuw verzoek
    lezer.lees(pdf, tmp_path / 'cache')
    assert len(api.verzoeken) == 2

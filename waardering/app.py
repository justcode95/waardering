"""Programmavenster (tkinter): dossiermap kiezen → uitlezen → keuzes nakijken → waardering maken."""
from __future__ import annotations

import datetime as dt
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import config
from .model import Beslissingen, Dossier
from .invoer import getal as _getal
from .samenstellen import bedrijfswinst, eur, signalen
from .verwerk import lees_dossier, maak_waardering


def openen(pad: Path) -> None:
    if sys.platform == 'win32':
        os.startfile(pad)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(['xdg-open', str(pad)])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Waardering')
        self.geometry('1100x760')
        self.cfg = config.laad()
        self.q: queue.Queue = queue.Queue()
        self.dossier: Dossier | None = None
        self.beslissing: Beslissingen | None = None
        self.kost = 0.0
        self.uitvoer: Path | None = None
        self._bouw()
        self.after(100, self._poll)

    # ------------------------------------------------------------------ opbouw
    def _bouw(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill='x')
        ttk.Label(top, text='Dossiermap').grid(row=0, column=0, sticky='w')
        self.map_var = tk.StringVar(value=self.cfg.get('laatste_map', ''))
        ttk.Entry(top, textvariable=self.map_var, width=90).grid(row=0, column=1, sticky='we', padx=4)
        ttk.Button(top, text='Bladeren…', command=self._kies_map).grid(row=0, column=2)
        ttk.Button(top, text='Instellingen…', command=self._instellingen).grid(row=0, column=3, padx=(8, 0))
        top.columnconfigure(1, weight=1)
        knoppen = ttk.Frame(self, padding=(8, 0))
        knoppen.pack(fill='x')
        self.b1 = ttk.Button(knoppen, text='1. Bestanden uitlezen', command=self._stap1)
        self.b1.pack(side='left')
        self.b2 = ttk.Button(knoppen, text='2. Waardering maken', command=self._stap2, state='disabled')
        self.b2.pack(side='left', padx=8)
        self.b3 = ttk.Button(knoppen, text='Excel openen', command=lambda: openen(self.uitvoer), state='disabled')
        self.b3.pack(side='left')
        self.b4 = ttk.Button(knoppen, text='Map openen', command=lambda: openen(Path(self.map_var.get())))
        self.b4.pack(side='left', padx=8)
        self.status = ttk.Label(knoppen, text='')
        self.status.pack(side='right')

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=8, pady=8)
        self.nb = nb
        self.overzicht = tk.Text(nb, wrap='word', font=('Consolas', 10))
        nb.add(self.overzicht, text='Overzicht')
        self.keuzes = ttk.Frame(nb, padding=10)
        nb.add(self.keuzes, text='Keuzes')
        self.log = tk.Text(nb, wrap='word', font=('Consolas', 9))
        nb.add(self.log, text='Logboek')
        self._keuzeformulier()

    def _keuzeformulier(self):
        f = self.keuzes
        self.v = {}
        rij = 0

        def veld(label, sleutel, uitleg='', breedte=18):
            nonlocal rij
            ttk.Label(f, text=label).grid(row=rij, column=0, sticky='w', pady=2)
            self.v[sleutel] = tk.StringVar()
            lang = breedte > 18                      # lange velden over beide kolommen, uitleg eronder
            ttk.Entry(f, textvariable=self.v[sleutel], width=breedte).grid(row=rij, column=1,
                                                                         columnspan=2 if lang else 1, sticky='w')
            if uitleg and not lang:
                ttk.Label(f, text=uitleg, foreground='#666').grid(row=rij, column=2, sticky='w', padx=8)
            rij += 1

        ttk.Label(f, text='Afsluitdatum (waarderingsdatum)').grid(row=rij, column=0, sticky='w')
        self.afsl_cb = ttk.Combobox(f, width=16, state='readonly')
        self.afsl_cb.grid(row=rij, column=1, sticky='w')
        rij += 1
        ttk.Label(f, text='Kolommen en weging (%)').grid(row=rij, column=0, sticky='nw', pady=(8, 2))
        kf = ttk.Frame(f)
        kf.grid(row=rij, column=1, columnspan=2, sticky='w', pady=(8, 6))
        self.kol_cb, self.weg = {}, {}
        for i, (k, naam) in enumerate((('H', 'boekjaar 1'), ('K', 'boekjaar 2'), ('N', 'boekjaar 3'),
                                       ('R', 'prognose / tussentijds'))):
            ttk.Label(kf, text=f'{k}  {naam}', width=24).grid(row=i, column=0, sticky='w')
            self.kol_cb[k] = ttk.Combobox(kf, width=12, state='readonly')
            self.kol_cb[k].grid(row=i, column=1, padx=(0, 8), pady=1)
            self.weg[k] = tk.StringVar()
            ttk.Entry(kf, textvariable=self.weg[k], width=7).grid(row=i, column=2)
            ttk.Label(kf, text='%').grid(row=i, column=3, sticky='w')
        rij += 1
        veld('Marktconforme bestuurdersvergoeding', 'bestuurder', 'EUR/jaar, totale kost; leeg = niet normaliseren')
        self.info618 = ttk.Label(f, text='', foreground='#666')
        self.info618.grid(row=rij, column=1, columnspan=2, sticky='w')
        rij += 1
        veld('Marktconforme huur', 'huur', 'EUR/jaar; leeg = werkelijke huur behouden')
        veld('EBITDA-multiple', 'multiple', 'bv. 4,5 (Vlerick M&A Monitor)')
        veld('Bron multiple (bv. VLERICK M&A Monitor 2026, horeca)', 'multiple_bron', '', 40)
        veld('Venale waarde onroerend goed', 'vastgoed', 'EUR volgens schattingsverslag; leeg = boekwaarde')
        veld('Datum schattingsverslag', 'vastgoed_datum', '')
        veld('Correctie financiering werkkapitaal', 'wk', 'EUR; 0 = geen')
        veld('Goodwill (aantal jaar)', 'gw_jaren', '')
        veld('Vereist rendement op EV', 'rendement', '%; leeg = volgens samenstelling actief (Goodwill E31)')
        veld('Activiteit', 'activiteit', '', 60)
        self.rc_var = tk.BooleanVar()
        ttk.Checkbutton(f, text='Intrest op de R/C van de bestuurder neutraliseren', variable=self.rc_var).grid(
            row=rij, column=0, columnspan=3, sticky='w', pady=(6, 0))

    # ------------------------------------------------------------------ acties
    def _kies_map(self):
        m = filedialog.askdirectory(initialdir=self.map_var.get() or None)
        if m:
            self.map_var.set(m)

    def _instellingen(self):
        w = tk.Toplevel(self)
        w.title('Instellingen')
        w.transient(self)
        sleutel = tk.StringVar(value=self.cfg.get('api_key', ''))
        model = tk.StringVar(value=self.cfg.get('model', 'claude-opus-5'))
        tpl = tk.StringVar(value=self.cfg.get('template', ''))
        ttk.Label(w, text='Anthropic API-sleutel').grid(row=0, column=0, sticky='w', padx=8, pady=4)
        ttk.Entry(w, textvariable=sleutel, show='•', width=60).grid(row=0, column=1, padx=8)
        ttk.Label(w, text='Model').grid(row=1, column=0, sticky='w', padx=8, pady=4)
        ttk.Combobox(w, textvariable=model, values=['claude-opus-5', 'claude-opus-5-5', 'claude-sonnet-5'],
                     width=20).grid(row=1, column=1, sticky='w', padx=8)
        ttk.Label(w, text='Template').grid(row=2, column=0, sticky='w', padx=8, pady=4)
        ttk.Entry(w, textvariable=tpl, width=60).grid(row=2, column=1, padx=8)
        ttk.Button(w, text='…', width=3, command=lambda: tpl.set(
            filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')]) or tpl.get())).grid(row=2, column=2)

        def ok():
            self.cfg.update(api_key=sleutel.get().strip(), model=model.get().strip(), template=tpl.get().strip())
            config.bewaar(self.cfg)
            w.destroy()
        ttk.Button(w, text='Bewaren', command=ok).grid(row=3, column=1, sticky='e', padx=8, pady=8)

    def _stap1(self):
        map_ = Path(self.map_var.get())
        if not map_.is_dir():
            messagebox.showerror('Waardering', 'Kies eerst een dossiermap.')
            return
        if not (self.cfg.get('api_key') or os.environ.get('ANTHROPIC_API_KEY')):
            messagebox.showerror('Waardering', 'Vul eerst de Anthropic API-sleutel in bij Instellingen.')
            return
        self.cfg['laatste_map'] = str(map_)
        config.bewaar(self.cfg)
        self._bezig(True, 'Bestanden uitlezen …')
        self.nb.select(self.log)
        self._draai(lambda: lees_dossier(map_, self.cfg, self._log), self._na_stap1)

    def _na_stap1(self, res):
        self.dossier, self.beslissing, self.kost = res
        self._toon_overzicht()
        self._vul_formulier()
        self._bezig(False, 'Kijk de keuzes na en klik op "2. Waardering maken".')
        self.b2.configure(state='normal')
        self.nb.select(self.keuzes)

    def _stap2(self):
        try:
            b = self._lees_formulier()
        except ValueError as e:
            messagebox.showerror('Waardering', str(e))
            return
        self.beslissing = b
        self._bezig(True, 'Waardering maken …')
        self.nb.select(self.log)
        map_ = Path(self.map_var.get())
        self._draai(lambda: maak_waardering(map_, self.dossier, b, self.cfg, self._log, self.kost), self._na_stap2)

    def _na_stap2(self, res):
        self.uitvoer, resultaat, ok, fout = res
        self._bezig(False, f'Klaar: {self.uitvoer.name}')
        self.b3.configure(state='normal')
        tekst = [f'WAARDERING {self.dossier.naam}', '']
        for k, v in resultaat.items():
            tekst.append(f'{k:32} {eur(v) if isinstance(v, (int, float)) else "–":>14}')
        tekst += ['', 'CONTROLES'] + [f'  OK    {x}' for x in ok] + [f'  FOUT  {x}' for x in fout]
        tekst += ['', f'Excel:   {self.uitvoer}', f'Verslag: {self.uitvoer.with_suffix(".md")}']
        self.overzicht.delete('1.0', 'end')
        self.overzicht.insert('1.0', '\n'.join(tekst))
        self.nb.select(self.overzicht)
        if fout:
            messagebox.showwarning('Waardering', f'{len(fout)} controle(s) niet in orde; zie Overzicht.')

    # ------------------------------------------------------------------ formulier
    def _vul_formulier(self):
        d, b = self.dossier, self.beslissing
        eindes = sorted({p.einde for p in d.periodes})
        self.afsl_cb['values'] = [x.isoformat() for x in eindes if not d.periode_op(x).tussentijds]
        self.afsl_cb.set(b.afsluitdatum.isoformat())
        for k in 'HKNR':
            self.kol_cb[k]['values'] = ['–'] + [x.isoformat() for x in eindes]
            self.kol_cb[k].set(b.kolommen[k].isoformat() if b.kolommen.get(k) else '–')
            self.weg[k].set(f'{b.weging.get(k, 0) * 100:.2f}'.rstrip('0').rstrip('.'))
        zet = lambda k, x: self.v[k].set('' if x is None else (str(int(x)) if float(x).is_integer()
                                                                else str(x).replace('.', ',')))
        zet('bestuurder', b.bestuurdersvergoeding)
        zet('huur', b.marktconforme_huur)
        zet('multiple', b.multiple)
        self.v['multiple_bron'].set(b.multiple_bron)
        zet('vastgoed', b.vastgoed_marktwaarde)
        self.v['vastgoed_datum'].set(b.vastgoed_schatting_datum)
        zet('wk', b.werkkapitaalcorrectie)
        self.v['gw_jaren'].set(str(b.goodwill_jaren))
        zet('rendement', None if b.vereist_rendement is None else round(b.vereist_rendement * 100, 2))
        self.v['activiteit'].set(b.activiteit)
        self.rc_var.set(b.rc_intrest_neutraliseren)
        huidig = [f'{p.einde.year}: {eur(-p.som("618"))}' for p in d.afgesloten()[-3:] if p.som('618')]
        self.info618.configure(text='Nu geboekt (618): ' + ', '.join(huidig) if huidig else '')

    def _lees_formulier(self) -> Beslissingen:
        kol = {k: (dt.date.fromisoformat(cb.get()) if cb.get() not in ('', '–') else None)
               for k, cb in self.kol_cb.items()}
        weging = {k: (_getal(self.weg[k].get()) or 0) / 100 for k in 'HKNR'}
        if abs(sum(weging.values()) - 1) > 0.001:
            raise ValueError('De wegingen moeten samen 100% zijn.')
        if any(weging[k] and not kol[k] for k in 'HKNR'):
            raise ValueError('Een kolom zonder periode kan geen weging krijgen.')
        rend = _getal(self.v['rendement'].get())
        return Beslissingen(
            afsluitdatum=dt.date.fromisoformat(self.afsl_cb.get()), kolommen=kol, weging=weging,
            bestuurdersvergoeding=_getal(self.v['bestuurder'].get()), marktconforme_huur=_getal(self.v['huur'].get()),
            multiple=_getal(self.v['multiple'].get()), multiple_bron=self.v['multiple_bron'].get().strip(),
            vastgoed_marktwaarde=_getal(self.v['vastgoed'].get()),
            vastgoed_schatting_datum=self.v['vastgoed_datum'].get().strip(),
            werkkapitaalcorrectie=_getal(self.v['wk'].get()) or 0.0,
            rc_intrest_neutraliseren=self.rc_var.get(), goodwill_jaren=int(_getal(self.v['gw_jaren'].get()) or 7),
            vereist_rendement=None if rend is None else rend / 100, activiteit=self.v['activiteit'].get().strip())

    def _toon_overzicht(self):
        d, b = self.dossier, self.beslissing
        t = [f'{d.naam}  –  {d.straat}, {d.postcode_gemeente}  –  ON {d.ondernemingsnummer}',
             f'Activiteit: {d.activiteit}', '',
             f'{"Periode":24}{"mnd":>4}{"Omzet":>14}{"Bedrijfswinst":>15}{"Afschr. (63)":>14}{"Bestuurder 618":>16}']
        for p in sorted(d.periodes, key=lambda p: p.einde):
            t.append(f'{p.einde.isoformat() + (" (tussentijds)" if p.tussentijds else ""):24}{p.maanden:>4}'
                     f'{eur(p.som("70")):>14}{eur(bedrijfswinst(p)):>15}{eur(-p.som("63")):>14}'
                     f'{eur(-p.som("618")):>16}')
        t += ['', f'Activa uit afschrijvingstabel: {len(d.activa)}',
              f'Bestuurders: {", ".join(x["naam"] for x in d.bestuurders) or "?"}',
              f'Aandeelhouders: {", ".join(x["naam"] + (" (" + str(int(x["aandelen"])) + ")" if x.get("aandelen") else "") for x in d.aandeelhouders) or "?"}',
              '', 'MELDINGEN'] + [f'  - {m}' for m in d.meldingen] + ['', 'SIGNALEN VOOR NORMALISATIE'] + \
             [f'  - {s}' for s in signalen(d, b)]
        self.overzicht.delete('1.0', 'end')
        self.overzicht.insert('1.0', '\n'.join(t))

    # ------------------------------------------------------------------ achtergrondtaken
    def _log(self, tekst: str):
        self.q.put(('log', tekst))

    def _draai(self, taak, klaar):
        def werk():
            try:
                self.q.put(('klaar', (klaar, taak())))
            except Exception as e:  # noqa: BLE001 - fout tonen aan de gebruiker
                self.q.put(('fout', e))
        threading.Thread(target=werk, daemon=True).start()

    def _poll(self):
        try:
            while True:
                soort, inhoud = self.q.get_nowait()
                if soort == 'log':
                    self.log.insert('end', inhoud + '\n')
                    self.log.see('end')
                elif soort == 'klaar':
                    inhoud[0](inhoud[1])
                elif soort == 'fout':
                    self._bezig(False, 'Fout')
                    self.log.insert('end', f'FOUT: {inhoud}\n')
                    messagebox.showerror('Waardering', str(inhoud))
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _bezig(self, bezig: bool, tekst: str):
        self.status.configure(text=tekst)
        for b in (self.b1, self.b2):
            b.configure(state='disabled' if bezig else ('normal' if b is self.b1 or self.dossier else 'disabled'))


def main():
    App().mainloop()


if __name__ == '__main__':
    main()

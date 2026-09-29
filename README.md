# Waardering

Windows-programma dat een KMO waardeert op basis van de bestanden in een dossiermap (jaarrekeningen of fiscale
bundels, tussentijdse cijfers, afschrijvingstabel, lijsten openstaande klanten/leveranciers) en het vaste
waarderingstemplate. Wat per bedrijf verschilt (bestuurdersvergoeding, weging, multiple, vastgoed …) vraagt het
programma aan jou, met een voorstel.

## Hoe het werkt

1. **Dossiermap kiezen** met alle PDF's van het bedrijf (ook submappen). Scans mogen; het boekhoudkantoor maakt
   niet uit.
2. **1. Bestanden uitlezen**: elke PDF wordt door Claude (Anthropic API) omgezet naar rekeningnummers en saldi per
   rubriek, de activa van de afschrijvingstabel, identificatie, bestuurders en aandeelhouders. Het programma
   controleert zelf of elke rubriek optelt tot haar totaal en of bedrijfswinst en winst vóór belasting kloppen;
   wat niet klopt wordt één keer opnieuw gelezen en anders gemeld. Resultaten worden per bestand bewaard in
   `.waardering_cache`, dus opnieuw draaien kost niets.
3. **Keuzes nakijken** (tabblad Keuzes): afsluitdatum, welke jaren in welke kolom en met welke weging,
   marktconforme bestuurdersvergoeding, huur, multiple, venale waarde vastgoed, werkkapitaalcorrectie, goodwill.
   Het tabblad Overzicht toont de cijfers per jaar, meldingen en signalen (bv. geboekte 618-vergoeding, eenmalige
   posten, cadeaubonnen in 493).
4. **2. Waardering maken**: het template wordt ingevuld (alles als traceerbare formules `=1234,56+789`),
   herberekend met Excel en gecontroleerd. In de dossiermap komen:
   - `<Naam> - Waardebepaling - <dd-mm-jjjj>.xlsx`
   - `<Naam> - Waardebepaling - <dd-mm-jjjj>.md`: verslag met resultaat, keuzes, controles en openstaande punten
   - `waardering_keuzes.json`: je keuzes, zodat een volgende run ze opnieuw voorstelt.

Controles na het invullen: EBITDA per jaar = bedrijfswinst + afschrijvingen (tabblad FCF, V58:X58), MVA
aanschafwaarde = afschrijvingstabel, MVA netto boekwaarde = balans, balanstotaal = jaarrekening, geen foutwaarden.

## Installatie (Windows)

**Optie A – kant-en-klare .exe** (na `build_exe.bat`, zie onder): dubbelklik `Waardering.exe`.

**Optie B – met Python:**

1. Installeer Python 3.11 of nieuwer van python.org (vink *Add python.exe to PATH* aan).
2. In deze map: `pip install -r requirements.txt`
3. Starten: dubbelklik `Waardering starten.bat` (of in PowerShell in deze map: `python -m waardering`)

Eenmalig: klik **Instellingen…** en vul je Anthropic API-sleutel in (aan te maken op console.anthropic.com; betalen
per gebruik, los van een Claude-abonnement). De sleutel wordt bewaard in `%APPDATA%\Waardering\config.json`.

Excel moet op de pc staan voor het herberekenen en de controles. Zonder Excel gebruikt het programma LibreOffice
als dat geïnstalleerd is; anders rekent Excel bij het openen en worden de controles overgeslagen.

## Kost

Het uitlezen gebruikt standaard Claude Opus 5 (instelbaar). Een tekst-PDF van enkele pagina's kost enkele
eurocenten; een ingescande fiscale bundel van ±90 pagina's ongeveer een halve tot één dollar. Het logboek toont
per bestand de gebruikte tokens en de geschatte kost; bestanden in de cache worden niet opnieuw aangerekend.

## Vertrouwelijkheid

De PDF's worden voor het uitlezen naar de Anthropic API gestuurd. Toets dit aan het beleid van je kantoor.
Zet geen klantdossiers in deze repository (`.gitignore` sluit ze uit).

## Het template

`template/Waardering_template.xlsx` volgt de indeling van de bestaande dossiers (Lopritec, Dembofisk, Duinenhoeve)
met de verbeteringen uit de recentere dossiers:

- weging over H/K/N én de prognosekolom R, EBITDA-formule van de prognose gecorrigeerd;
- vervangingsinvesteringen per sectie = af te schrijven waarde roerende goederen / gebruiksduur, onroerend goed
  niet inbegrepen; TOTAAL MVA telt alle secties;
- rendementswaarde 2: te ontlenen kapitaal met een exacte formule in plaats van handmatig itereren;
- vereist rendement op EV volgens samenstelling van het actief (Goodwill E31), overschrijfbaar in het programma;
- normalisaties: bestuurdersvergoeding en huur via een invoercel (leeg = geen normalisatie), financiële
  normalisatie gekoppeld aan de werkkapitaalcorrectie (4%);
- geen klantgegevens, geen oude hulpbladen of DataSnipper-koppelingen; het verborgen SWOT-blad blijft.

Je kan een eigen versie van het template kiezen bij Instellingen, zolang de indeling (rijen en kolommen) gelijk
blijft; de celadressen staan in `waardering/invullen.py`.

## .exe bouwen

Op een Windows-pc met Python: `build_exe.bat`. Het resultaat staat in `dist\Waardering\`.

## Ontwikkeling

```
pip install -r requirements.txt pytest
python -m pytest -q
```

De tests gebruiken een verzonnen testbedrijf en een nagebootste API; de volledige-waarderingstests draaien alleen
als LibreOffice (`soffice`) beschikbaar is.

Onderdelen (`waardering/`):

| Bestand | Rol |
|---|---|
| `app.py` | venster (tkinter) |
| `verwerk.py` | de twee stappen: dossier uitlezen, waardering maken |
| `uitlezen.py` | PDF → JSON via de Anthropic API, controle per rubriek, cache |
| `samenstellen.py` | documenten samenvoegen tot een dossier, voorstellen en signalen |
| `invullen.py` | template invullen (FCF, MVA, eigen vermogen, goodwill, teksten) |
| `herberekenen.py` | herberekenen met Excel/LibreOffice en controles |
| `rapport.py` | verslag in Markdown |
| `xlsxedit.py` | Excel-bestanden bewerken zonder opmaak, afbeeldingen of koppelingen te verliezen |

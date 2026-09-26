# Waardering

Windows-programma (Python/tkinter) dat een Belgische KMO waardeert: PDF's uit een dossiermap worden via de
Anthropic API uitgelezen, samengevoegd, en het vaste template `template/Waardering_template.xlsx` wordt ingevuld.

## Regels

- Nooit klantgegevens committen (PDF's, ingevulde Excel-bestanden, `.waardering_cache`). Tests gebruiken het
  verzonnen testbedrijf in `tests/`.
- Het template bewerken we enkel via `waardering/xlsxedit.py` (XML), nooit via openpyxl-save: dat verliest
  afbeeldingen, koppelingen en opmaak.
- Celadressen van het template staan in `waardering/invullen.py`; wijzigt het template, pas dan beide aan en draai
  de tests (`python -m pytest -q`, volledige waarderingstests vragen LibreOffice).
- Normalisaties zijn beslissingen van de gebruiker: het programma stelt voor en meldt, maar kiest niet zelf een
  bestuurdersvergoeding, huur, multiple of venale waarde.
- Vervangingsinvesteringen = af te schrijven waarde roerende goederen / gebruiksduur per sectie.

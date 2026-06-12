
# HS-code invoice tool

Deze tool vult automatisch HS-codes in Excel-invoices.

## Wat doet de tool?
- Zoekt de kolom met `Article No.` / `Artikelnummer`.
- Zoekt of maakt de kolom `HS code`.
- Koppelt artikelnummer aan HS-code via `hs_mapping.csv`.
- Gebruikt geen HSN-codes uit de invoice.
- Maakt een nieuwe Excel-output; het originele bestand wordt niet overschreven.

## Starten als upload-tool

1. Installeer Python 3.10 of nieuwer.
2. Open Terminal / Command Prompt in deze map.
3. Installeer benodigdheden:

```bash
pip install -r requirements.txt
```

4. Start de app:

```bash
streamlit run app.py
```

5. Upload je invoice en download de versie met HS-codes.

## Gebruiken via command line

```bash
python fill_hs_codes.py "invoice.xlsx" "invoice_met_HS_codes.xlsx"
```

## Artikeloverzicht bijwerken

Vervang `hs_mapping.csv` door een nieuwe mapping met minimaal deze kolommen:

```csv
artikelnummer,hs_code
7300 ZW 38,64039110
```

De meegeleverde mapping is gemaakt op basis van het bestand `Artikeloverzicht compleet HS codes 12-5-2026`.




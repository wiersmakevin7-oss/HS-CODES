# HS-code invoice tool

Streamlit-app om HS-codes op invoices aan te vullen op basis van artikelnummer.

## Wat doet de tool?

- Upload een Excel- of PDF-invoice.
- Leest artikelcodes uit de invoice.
- Zoekt per artikelcode de HS-code in `hs_mapping.csv`.
- Gebruikt geen HSN-/HS-code uit de invoice zelf.
- Excel-upload: behoudt de originele workbook en zet rechts een kolom `HS code (artikellijst)`.
- PDF-upload: maakt een nieuwe Excel met artikelcode, artikelnaam, HS-code, aantallen, waarde, totaalwaarde en PDF-controletext.
- Zet onder PDF-exports een rij `Factuur waarde` met de totale waarde.

## Projectbestanden

Voor deployment moeten deze bestanden in de GitHub-repository staan:

- `app.py`
- `fill_hs_codes.py`
- `hs_mapping.csv`
- `requirements.txt`
- `.streamlit/config.toml`

## Lokaal starten

Gebruik Python 3.12.

```powershell
cd "C:\Users\KevinWiersmaTeamFrei\OneDrive - TFF & Logistics\Documenten\HS Code brands of q\hs_code_tool"
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Open daarna:

```text
http://localhost:8501
```

## Streamlit Community Cloud deployment

1. Push de laatste versie naar GitHub.
2. Ga naar `https://share.streamlit.io`.
3. Klik op `Create app`.
4. Kies `Yup, I have an app`.
5. Selecteer repository:

```text
wiersmakevin7-oss/HS-CODES
```

6. Selecteer branch:

```text
main
```

7. Selecteer main file path:

```text
app.py
```

8. Open `Advanced settings`.
9. Kies Python versie:

```text
3.12
```

10. Secrets zijn niet nodig voor deze app.
11. Klik op `Deploy`.

Streamlit Community Cloud leest automatisch `requirements.txt` en installeert de packages. De app krijgt daarna een link op `streamlit.app` die je met collega's kunt delen.

## Privacy

Facturen worden bij gebruik van Streamlit Community Cloud geupload naar een externe cloudomgeving. Gebruik dit alleen als dat intern is toegestaan voor jullie invoice-, klant- en prijsinformatie.

## Artikeloverzicht bijwerken

Vervang `hs_mapping.csv` door een nieuwe CSV met minimaal:

```csv
artikelnummer,omschrijving,hs_code
7300 ZW 38,Rijbroek voorbeeld,64039110
```

## Veelvoorkomende problemen

### `ModuleNotFoundError`

Controleer of `requirements.txt` is geinstalleerd:

```powershell
python -m pip install -r requirements.txt
```

### PDF wordt niet goed gelezen

De tool ondersteunt meerdere invoice-layouts. Als een nieuwe leverancier niet goed wordt gelezen, voeg een voorbeeld-PDF toe en breid de parser in `fill_hs_codes.py` uit.

### OCR/PDF dependencies

`rapidocr_onnxruntime` wordt gebruikt voor gescande PDF's. Gebruik Python 3.12, omdat nieuwere Python-versies niet altijd door alle OCR-dependencies worden ondersteund.

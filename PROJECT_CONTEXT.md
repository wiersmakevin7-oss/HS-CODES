# Project Context - HS-code invoice tool

Gebruik dit bestand als overdracht naar een nieuwe Codex-chat of een ander account.

## Doel

Deze Streamlit-app vult HS-codes aan op basis van artikelnummer.

De app kan:

- Excel-invoices uploaden.
- PDF-invoices uploaden.
- Artikelcodes uitlezen.
- HS-codes zoeken in `hs_mapping.csv`.
- Artikelnaam toevoegen vanuit `hs_mapping.csv`.
- Bij PDF-exports ook aantallen, waarde per stuk en totaalwaarde toevoegen wanneer de PDF-layout dit ondersteunt.
- Een nieuwe Excel downloaden.

Belangrijk: gebruik geen HSN-/HS-code uit de invoice zelf als bron. De lookup moet altijd via de artikellijst/mapping gaan.

## Projectmap

Lokale map:

```text
C:\Users\KevinWiersmaTeamFrei\OneDrive - TFF & Logistics\Documenten\HS Code brands of q\hs_code_tool
```

GitHub:

```text
https://github.com/wiersmakevin7-oss/HS-CODES.git
```

Branch:

```text
main
```

## Belangrijke bestanden

- `app.py` - Streamlit UI.
- `fill_hs_codes.py` - alle Excel/PDF parsing en exportlogica.
- `hs_mapping.csv` - artikelnummer, omschrijving en HS-code.
- `requirements.txt` - Python dependencies.
- `.streamlit/config.toml` - Streamlit config.
- `README.md` - lokale start- en Streamlit Community Cloud instructies.

## Lokaal starten

```powershell
cd "C:\Users\KevinWiersmaTeamFrei\OneDrive - TFF & Logistics\Documenten\HS Code brands of q\hs_code_tool"
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Open:

```text
http://localhost:8501
```

Er staat ook een bureaubladknop:

```text
Start HS Code Tool.bat
```

## Streamlit Community Cloud

De app is voorbereid voor Streamlit Community Cloud.

Gebruik:

```text
Repository: wiersmakevin7-oss/HS-CODES
Branch: main
Main file path: app.py
Python version: 3.12
Secrets: niet nodig
```

Let op:

- Bij een public repo kan iedereen de code en `hs_mapping.csv` zien.
- Alleen mensen met schrijfrechten kunnen de repo aanpassen.
- Geuploade facturen komen niet in GitHub, maar worden wel verwerkt via Streamlit Cloud.

## Requirements

Belangrijke packages:

- `streamlit`
- `openpyxl`
- `pypdf`
- `pdfplumber`
- `pypdfium2`
- `rapidocr_onnxruntime`
- `pandas`

Gebruik Python 3.12. OCR-dependencies kunnen problemen geven met nieuwere Python-versies.

## Excel-logica

Voor Excel-invoices:

- Als er een blad `INVOICE` bestaat, verwerk alleen dat blad.
- Sla `PACKING LIST`-achtige bladen over wanneer van toepassing.
- Zoek artikelkolommen zoals:
  - `Article No.`
  - `Art No.`
  - `Buyer Art No.`
  - `Buyer Article No.`
- Voeg de HS-code helemaal rechts toe als `HS code (artikellijst)`.
- Behoud de originele workbook-layout en formules zoveel mogelijk.
- Vul per regel, niet samengevoegd per uniek artikel.

## PDF-output

PDF-export maakt een nieuwe Excel met kolommen:

```text
Page
Article No.
Article name
HS code
Aantallen
Waarde
Totaal waarde
PDF text
```

Onderaan staat:

```text
Factuur waarde
```

met de som van alle waarden in `Totaal waarde`.

PDF-rijen worden per basisartikelgroep gekleurd zodat bijbehorende regels en waardes visueel bij elkaar horen.

## Belangrijke PDF-formaten die al ondersteund zijn

De parser is gaandeweg uitgebreid voor meerdere leveranciers/layouts, onder andere:

- Brading invoice: article no kolom over meerdere pagina's.
- Maharaja invoice: OCR/scanned PDF.
- PO 502486INV PL: packing list pagina overslaan.
- Kartiyeka: `Our Product Code`.
- MRS Excel: `Buyer Art No`.
- Mehra Shoes: `Party's Code`.
- Guts en Glory: `Art.No`.
- Shifa: `Buyer Product Code`.
- Matrix-formaat links naast `ORD.NO.`.
- Mark Equestrian:
  - Artikelcode links naast `ORD.NO.`
  - Rechter kolommen voor aantallen/waarde/totaalwaarde
  - Variantcodes zoals `8687 DSR` worden herkend uit mapping-varianten zoals `8687 DSR 104`.
- Equi Style:
  - Kolommen `Quantity Pcs/Pairs`, `Per Piece Euro`, `Total Amount Euro`.
  - 73 regels getest.
- Sea shipment invoice:
  - Kolommen `QUANTITY (PCS)`, `UNIT PRICE`, `AMOUNT`.
  - Gebruikt `pdfplumber` woordposities.
  - 45 regels getest.
  - Factuurwaarde getest op `33,134.50`.

## Belangrijke functies in `fill_hs_codes.py`

- `load_catalog()` - laadt HS mapping plus artikelnaam en maakt aliases.
- `load_mapping()` - backward-compatible mapping loader.
- `extract_pdf_articles()` - hoofdroute voor PDF parsing.
- `fill_pdf_invoice()` - schrijft PDF-resultaten naar Excel.
- `fill_invoice()` - verwerkt Excel-invoices.
- `lookup_hs()` - zoekt HS-code op artikel en eventueel maat.
- `extract_pdf_articles_from_shipment_columns()` - parser voor sea shipment layout met quantity/unit price/amount.
- `extract_pdf_articles_from_layout()` - algemene layout parser.
- `extract_pdf_articles_from_ocr()` - OCR fallback.

## Testcommando's

Syntaxcheck:

```powershell
python -m py_compile app.py fill_hs_codes.py
```

Imports checken:

```powershell
python -c "import streamlit, openpyxl, pypdf, pypdfium2, pdfplumber; print('imports ok')"
```

Git status:

```powershell
git status --short
```

Laatste commit:

```powershell
git log -1 --oneline
```

Push naar GitHub:

```powershell
git add .
git commit -m "Beschrijving van wijziging"
git push origin main
```

## Recente Git commits

Belangrijke recente commits:

```text
df22501 Improve PDF value extraction and export formatting
cbe1ce4 Prepare Streamlit Cloud deployment
```

## Aandachtspunten voor volgende Codex-chat

- Werk altijd in de map `hs_code_tool`.
- Gebruik `apply_patch` voor codewijzigingen.
- Revert geen bestaande wijzigingen zonder expliciete vraag.
- Test nieuwe PDF-formaten met concrete voorbeeldbestanden in `Downloads`.
- Als een PDF eerst goed leek maar later regels mist, controleer altijd hoeveel regels oude fallback-parsers vinden.
- Bij waarde-extractie: bedragen kunnen als `US$1,234.50`, `USD6.84`, `USD.77` of gewone `4.45` in PDF staan.
- Houd `requirements.txt` in sync met imports. `pdfplumber` is nodig voor sommige PDF-layouts.
- Na wijzigingen die op Streamlit Cloud moeten komen: commit en push naar `origin/main`.

## Prompt voor nieuwe chat

Gebruik bijvoorbeeld:

```text
Lees PROJECT_CONTEXT.md in de projectmap. Ga verder met de HS-code Streamlit app. Werk in hs_code_tool en houd rekening met de bestaande PDF/Excel parsers.
```

# Codex Handoff - HS-code invoice tool

Nieuwe Codex-chat: lees dit bestand eerst en werk in deze map:

```text
C:\Users\KevinWiersmaTeamFrei\OneDrive - TFF & Logistics\Documenten\HS Code brands of q\hs_code_tool
```

## Korte opdracht

Onderhoud en verbeter een Streamlit-app die Excel- en PDF-invoices verwerkt.

De app moet:

- Artikelcodes uit invoices lezen.
- HS-codes opzoeken in `hs_mapping.csv`.
- Geen HSN-/HS-code uit de invoice zelf gebruiken.
- Bij Excel de originele workbook behouden en rechts een HS-code kolom toevoegen.
- Bij PDF een nieuwe Excel maken met:
  - `Page`
  - `Article No.`
  - `Article name`
  - `HS code`
  - `Aantallen`
  - `Waarde`
  - `Totaal waarde`
  - `PDF text`
- Bij PDF onderaan `Factuur waarde` zetten met de som van `Totaal waarde`.

## Belangrijke bestanden

```text
app.py
fill_hs_codes.py
hs_mapping.csv
requirements.txt
README.md
PROJECT_CONTEXT.md
.streamlit/config.toml
```

## GitHub

```text
https://github.com/wiersmakevin7-oss/HS-CODES.git
branch: main
```

Na wijzigingen die live moeten:

```powershell
git status --short
git add .
git commit -m "Beschrijf wijziging"
git push origin main
```

## Streamlit Cloud instellingen

```text
Repository: wiersmakevin7-oss/HS-CODES
Branch: main
Main file path: app.py
Python version: 3.12
Secrets: niet nodig
```

## Lokaal starten

```powershell
cd "C:\Users\KevinWiersmaTeamFrei\OneDrive - TFF & Logistics\Documenten\HS Code brands of q\hs_code_tool"
python -m streamlit run app.py
```

URL:

```text
http://localhost:8501
```

Als `streamlit` ontbreekt:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Testcommando's

```powershell
python -m py_compile app.py fill_hs_codes.py
python -c "import streamlit, openpyxl, pypdf, pypdfium2, pdfplumber; print('imports ok')"
```

## Belangrijke parserregels

- `fill_invoice()` verwerkt Excel.
- `fill_pdf_invoice()` schrijft PDF-resultaten naar Excel.
- `extract_pdf_articles()` bepaalt welke PDF-parser gebruikt wordt.
- `load_catalog()` laadt artikelnummer, artikelnaam en HS-code.
- `lookup_hs()` zoekt HS-code op artikelcode en eventueel maat.
- `extract_pdf_articles_from_shipment_columns()` gebruikt `pdfplumber` voor kolommen `QUANTITY (PCS)`, `UNIT PRICE`, `AMOUNT`.
- `extract_pdf_articles_from_layout()` is belangrijke fallback en vond o.a. 73 regels in `Invoice equi style.pdf`.
- `extract_pdf_articles_from_ocr()` is fallback voor scans.

## Bekende ondersteunde PDF-layouts

- `Article No.`
- `Art.No`
- `Buyer Product Code`
- `Our Product Code`
- `Party's Code`
- Artikel links naast `ORD.NO.`
- Matrixfacturen met rechter kolommen voor aantallen/waarde/totaal.
- Equi Style factuur met `Quantity Pcs/Pairs`, `Per Piece Euro`, `Total Amount Euro`.
- Sea shipment factuur met `QUANTITY (PCS)`, `UNIT PRICE`, `AMOUNT`.

## Recente belangrijke testresultaten

`Invoice equi style.pdf`

```text
73 regels
0 missende HS-codes
0 missende waarden
```

`invo sea shipment-105277 502511 105357 501923.pdf`

```text
45 regels
0 missende HS-codes
0 missende waarden
Factuur waarde: 33,134.50
```

`MARK EQ CI (2).pdf`

```text
16 regels
0 missende HS-codes
variantcodes zoals 8687 DSR worden herkend
```

## Let op

- Nieuwe PDF-layouts altijd testen op aantal gevonden regels.
- Als een nieuwe parser minder regels vindt dan een fallback, niet te vroeg returnen.
- Amounts kunnen formats hebben zoals:
  - `US$684.00`
  - `US$1,026.00`
  - `USD6.84`
  - `USD.77`
  - `4.45`
- `parse_pdf_number()` moet bedragen als echte Excel-getallen opslaan.
- `requirements.txt` moet alle imports bevatten, waaronder `pdfplumber`.
- Gebruik `apply_patch` voor codewijzigingen.
- Revert geen user changes zonder expliciete toestemming.

## Prompt voor nieuwe Codex-chat

```text
Lees CODEX_HANDOFF.md en PROJECT_CONTEXT.md. Ga verder met de HS-code Streamlit app in de map hs_code_tool. Houd de bestaande Excel/PDF parsers intact en test nieuwe invoice-layouts met concrete voorbeeldbestanden.
```

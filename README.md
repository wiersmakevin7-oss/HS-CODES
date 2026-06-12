kevin is een homo

HS-code invoice tool
Deze tool vult automatisch HS-codes in Excel-invoices.
Wat doet de tool?
Zoekt de kolom met `Article No.` / `Artikelnummer`.
Zoekt of maakt de kolom `HS code`.
Koppelt artikelnummer aan HS-code via `hs_mapping.csv`.
Gebruikt geen HSN-codes uit de invoice.
Maakt een nieuwe Excel-output; het originele bestand wordt niet overschreven.
Kan ook PDF-invoices verwerken wanneer de PDF/OCR-dependencies correct zijn geïnstalleerd.
Vereisten
Gebruik Python 3.10
> Let op: Python 3.13 of nieuwer wordt voor deze app niet aangeraden, omdat `rapidocr_onnxruntime` momenteel Python `<3.13` vereist. Met Python 3.13+ kan `pip install -r requirements.txt` daarom mislukken.
Controleer je Python-versies in PowerShell:
```powershell
py --list
python --version
```
Als Python 3.10 nog niet geïnstalleerd is, installeer die dan vanaf python.org en vink tijdens installatie Add python.exe to PATH aan.
Eerste installatie op Windows / PowerShell
Open PowerShell in deze projectmap en voer daarna uit:
```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```
Start daarna de app met:
```powershell
python -m streamlit run app.py
```
Gebruik bij voorkeur `python -m streamlit` in plaats van alleen `streamlit`, omdat dit altijd dezelfde Python-omgeving gebruikt waarin de packages zijn geïnstalleerd.
Als PowerShell de virtual environment blokkeert
Krijg je een melding dat scripts niet mogen worden uitgevoerd? Voer dan één keer uit:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```
Activeer daarna opnieuw:
```powershell
.\.venv\Scripts\Activate.ps1
```
Dagelijks starten
Als de installatie al eerder is gedaan, hoef je meestal alleen dit te doen:
```powershell
cd "C:\Users\EelcoHansmaTeamFreig\OneDrive - TFF & Logistics\Documenten\HS-CODES"
.\.venv\Scripts\Activate.ps1
python -m streamlit run app.py
```
Daarna opent Streamlit meestal automatisch in je browser. Zo niet, kopieer dan de lokale URL uit de terminal, meestal iets zoals:
```text
http://localhost:8501
```
Gebruiken als upload-tool
Start de app met `python -m streamlit run app.py`.
Upload je invoice.
Download de versie met HS-codes.
Gebruiken via command line
Voor Excel-invoices kun je de tool ook direct via de command line gebruiken:
```powershell
python fill_hs_codes.py "invoice.xlsx" "invoice_met_HS_codes.xlsx"
```
Artikeloverzicht bijwerken
Vervang `hs_mapping.csv` door een nieuwe mapping met minimaal deze kolommen:
```csv
artikelnummer,hs_code
7300 ZW 38,64039110
```
De meegeleverde mapping is gemaakt op basis van het bestand `Artikeloverzicht compleet HS codes 12-5-2026`.
Problemen oplossen
`streamlit` is not recognized
Start Streamlit via Python:
```powershell
python -m streamlit run app.py
```
Als dat ook niet werkt, installeer de dependencies opnieuw in de actieve virtual environment:
```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```
`ModuleNotFoundError: No module named 'openpyxl'` of `No module named 'pypdf'`
De dependencies zijn niet geïnstalleerd in de Python-omgeving waarmee je de app start. Activeer de virtual environment en installeer opnieuw:
```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```
Fout bij `rapidocr_onnxruntime`
Controleer eerst je Python-versie:
```powershell
python --version
```
Gebruik Python 3.12. Verwijder daarna eventueel de oude virtual environment en maak die opnieuw:
```powershell
Remove-Item -Recurse -Force .venv
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```
Controleren welke Python wordt gebruikt
```powershell
python -c "import sys; print(sys.executable)"
python -m pip show openpyxl pypdf streamlit rapidocr_onnxruntime
```
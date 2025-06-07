# BP Pricing Automation

Automation of pricing workflows for BP service stations using Python, Selenium, OCR and Excel.

---

## 🚀 Features

- **Daily Price Update** (`price_update_daily.py`):  
  Detects price changes by background-color in the Excel file and submits them via Selenium to the web portal.

- **Weekly Ratification** (`ratificacion_semanal.py`):  
  Every Friday runs the full cycle (login, price capture, digital signature, PDF export) for two users.

- **Folio Manager GUI** (`gestor_folios.py`):  
  PySide6 interface that reads folios from signed PDFs via OCR (Google Cloud Vision) and injects them back into the Excel sheets.

- **Data Extraction** (`email_extractor.py` / `email_extractor2.py` + `data_extractor.py` / `data_Extractor2.py`):  
  Pulls the price table from Outlook, pastes it into Excel and/or flags changed cells by color using xlwings.

- **Excel Template Generator** (`Excels.py`):  
  Creates a base workbook copy for each station.

---

## 📁 Repository Structure

bp-pricing-automation/
├── .env # Sensitive credentials & paths (gitignored)
├── .gitignore
├── README.md
├── requirements.txt
├── Actualización de precios OPE.xlsx # Input Excel (do NOT commit)
├── data_extractor.py
├── data_Extractor2.py
├── email_extractor.py
├── email_extractor2.py
├── Excels.py
├── gestor_folios.py
├── price_update_daily.py
└── ratificacion_semanal.py


---

## 📋 Prerequisites

- **Python 3.10+**  
- Install dependencies:

  ```bash
  pip install -r requirements.txt


🔧 Configuration
Environment Variables
Create a file named .env in the project root with your credentials and paths:

# User A
ARTURO_EMAIL=yourA@example.com
ARTURO_PASS=YourSecretPassA
ARTURO_CER_PATH=C:/hidden/path/arturo.cer
ARTURO_KEY_PATH=C:/hidden/path/arturo.key
ARTURO_KEY_PWD=YourKeyPasswordA

# User B
CARLOS_EMAIL=yourB@example.com
CARLOS_PASS=YourSecretPassB
CARLOS_CER_PATH=C:/hidden/path/carlos.cer
CARLOS_KEY_PATH=C:/hidden/path/carlos.key
CARLOS_KEY_PWD=YourKeyPasswordB

# Paths & options
ACUSES_BASE_PATH=C:/Users/you/Documents/Acuses
EXCEL_OPE_PATH=./Actualización de precios OPE.xlsx
OMIT_STATIONS=BP Taxqueña,BP Viveros,BP Ermita


Input Excel

Filename: Actualización de precios OPE.xlsx

Sheet: Precios

Header row: row 2

Key columns:

C = Station name

E/F/G = Regular, Premium, Diesel prices (cell background color signals change)

H = Group (e.g. “Arturo Aceves” or “Carlos Rodriguez”)

I = Company

Do NOT commit this file to the repo. It must remain local/in .gitignore.

⚙️ Usage
1. Daily price update
bash
Copiar
Editar
python price_update_daily.py
2. Weekly ratification (run on Fridays)
bash
Copiar
Editar
python ratificacion_semanal.py
3. Folio manager GUI
bash
Copiar
Editar
python gestor_folios.py
📦 Dependencies (requirements.txt)
text
Copiar
Editar
pandas==1.5.3
xlwings
pyperclip
beautifulsoup4==4.13.4
openpyxl==3.0.10
PyMuPDF==1.21.1
google-cloud-vision
PySide6
python-dotenv==1.1.0
webdriver-manager
selenium
pywin32
🔒 Security & Git
Never commit API keys, certificates or .env files.

Excel input (*.xlsx) is deliberately ignored by .gitignore.

📄 License
Internal use & educational purposes only. Do not redistribute without permission.

If you have questions or issues, please open an Issue in this repo.

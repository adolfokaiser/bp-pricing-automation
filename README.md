# BP Pricing Automation

Automation of pricing workflows for BP service stations using Python, Selenium, OCR, and Excel.

---

## 🚀 Features

* **Daily Price Update** (`price_update_daily.py`):
  Detects price changes in the Excel sheet based on background-color and submits them via Selenium to the web portal.

* **Weekly Ratification** (`ratificacion_semanal.py`):
  Automates the full cycle (login, price capture, digital signature, PDF export) every Friday for two user profiles.

* **Folio Manager GUI** (`gestor_folios.py`):
  PySide6-based interface that reads folios from signed PDFs using Google Cloud Vision OCR and inserts them into Excel.

* **Data Extraction** (`email_extractor.py`, `email_extractor2.py`, `data_extractor.py`, `data_Extractor2.py`):
  Extracts pricing tables from Outlook, pastes them into Excel, and highlights changes using `xlwings`.

* **Excel Template Generator** (`Excels.py`):
  Creates a personalized Excel workbook per service station.

* **National/BP Price Aggregator** (`merge_pricing_excels.py`):  
  Waits for all `.xlsx` pricing files (6 for BP, 15 for national), combines them into a single Excel workbook, normalizes the column names, and sends the result via Outlook to a predefined recipient.

* **CRE Price Scraper** (`cre_price_scraper.py`):  
  Uses Selenium and 2Captcha to scrape official fuel prices from the CRE website for selected municipalities (BP or national), iterating page by page and exporting the complete dataset to Excel.
  

---

## 📁 Repository Structure

```
bp-pricing-automation/
├── .env                        # Sensitive credentials & paths (ignored by Git)
├── .gitignore
├── README.md
├── requirements.txt
├── Actualización de precios OPE.xlsx  # Input Excel (do NOT commit)
├── merge_pricing_excels.py 
├── cre_price_scraper.py 
├── data_extractor.py
├── data_Extractor2.py
├── email_extractor.py
├── email_extractor2.py
├── Excels.py
├── gestor_folios.py
├── price_update_daily.py
└── ratificacion_semanal.py
```

---

## 📋 Prerequisites

* **Python 3.10+**
* Install dependencies:

```bash
pip install -r requirements.txt
```

---

## 🔧 Configuration

### Environment Variables

Create a `.env` file in the root folder with your credentials and paths:

```
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
```

### Input Excel

* **Filename**: `Actualización de precios OPE.xlsx`
* **Sheet**: `Precios`
* **Header row**: row 2
* **Key columns**:

  * Column C: Station name
  * Columns E/F/G: Regular, Premium, Diesel prices (background color signals change)
  * Column H: Group (e.g., "Arturo Aceves" or "Carlos Rodriguez")
  * Column I: Company

> **Note:** This file must remain local and is listed in `.gitignore`.

---

## ⚙️ Usage

1. **Run daily update**

   ```bash
   python price_update_daily.py
   ```

2. **Run weekly ratification (Fridays)**

   ```bash
   python ratificacion_semanal.py
   ```

3. **Launch GUI for folio management**

   ```bash
   python gestor_folios.py
   ```
4. **Run national/BP Excel merge and send email**

   ```bash
   python merge_pricing_excels.py bp
---

## 📦 Dependencies (`requirements.txt`)

```
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
send2trash
requests
urllib3
```

---

## 🔐 Security & Git

* Never commit `.env` files or certificates.
* `.gitignore` is configured to avoid pushing sensitive data and Excel files.

---

## 📄 License

This repository is intended for internal and educational use only. Do not redistribute without explicit permission.

If you have questions or run into issues, feel free to open an Issue in this repo.

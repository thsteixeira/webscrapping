# Webscrapping - Legal Process Automation

**Note: This is an archived project that is no longer in active development or deployment.**

Python automation scripts for scraping and managing legal process information from Brazilian court systems (Jurisconsult and PJE).

## Features

- **integra.py**: Automates client data retrieval from INTEGRA legal management system
- **incluir_push_selenium.py**: Automates process registration in Jurisconsult PUSH system
- **pesquisar_processos.py**: Scrapes legal process information from Jurisconsult and PJE courts
- **relatorio_processos.py**: Generates reports from scraped process data
- **solve_captcha.py**: CAPTCHA solving utilities for automated scraping

## Technology Stack

- Python 3
- Selenium WebDriver (Chrome)
- BeautifulSoup4 (HTML parsing)
- SQLite3 (local database)
- openpyxl (Excel file reading)
- Pillow + pytesseract (CAPTCHA solving)

## Setup

If you want to run this project locally for reference:

1. Clone the repository
2. Create a virtual environment: `python -m venv venv`
3. Activate it: `venv\Scripts\activate` (Windows) or `source venv/bin/activate` (Unix)
4. Install dependencies: `pip install -r requirements.txt`
5. Download ChromeDriver and place `chromedriver.exe` in the project root
6. Copy `.env.example` to `.env` and fill in your own credentials
7. Prepare Excel files with process/client lists in the `excel/` folder
8. Run the desired script: `python pesquisar_processos.py`

## Environment Variables

See `.env.example` for the required credentials if you want to use the INTEGRA or Jurisconsult PUSH integrations.

## Usage Notes

These scripts were designed for specific use cases with Brazilian court systems:
- **Jurisconsult (TJMA)**: Maranhão state court system
- **PJE**: Electronic judicial process system
- **INTEGRA**: Commercial legal practice management software

## License

This project is provided as-is for reference purposes only.

# Document Data Extractor
Python + Flask tool that extracts text from a PDF (up to 10 MB) or a website link and exports it as structured XML and Excel (.xlsx).

**Skills:** Python, Flask, PDF text extraction (pdfplumber), web scraping (requests + BeautifulSoup), XML generation, Excel export (openpyxl), input validation, error handling.

## Run locally
    pip install -r requirements.txt
    python app.py   # open http://127.0.0.1:5000

## Deploy (Render)
Build: `pip install -r requirements.txt` | Start: `gunicorn app:app`

## Limits
Scanned (image) PDFs need OCR and are not supported yet.

## 🌐 Live Demo

🔗 **Live :** [Click Here](https://doc-extractor-a8e7.onrender.com/)

"""Document Data Extractor: PDF (up to 10 MB) or website link -> clean XML."""
import io
import xml.etree.ElementTree as ET
from urllib.parse import urlparse

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import requests
from bs4 import BeautifulSoup
from flask import Flask, render_template, request, Response
from werkzeug.exceptions import RequestEntityTooLarge

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024  # a little above 10 MB
MAX_PDF_BYTES = 10 * 1024 * 1024


def pdf_to_sections(file_obj):
    """Each PDF page becomes one section; each text block becomes a paragraph."""
    sections = []
    with pdfplumber.open(file_obj) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            paras = [p.strip() for p in text.split("\n\n") if p.strip()]
            if not paras and text.strip():
                paras = [text.strip()]
            sections.append({"title": f"Page {i}", "paragraphs": paras})
    return sections


def web_to_sections(url):
    """Headings (h1-h3) start a section; paragraphs go under the latest heading."""
    r = requests.get(url, timeout=10, headers={"User-Agent": "DocExtractor/1.0"})
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    sections, current = [], {"title": "Introduction", "paragraphs": []}
    for tag in soup.find_all(["h1", "h2", "h3", "p"]):
        text = tag.get_text(" ", strip=True)
        if not text:
            continue
        if tag.name == "p":
            current["paragraphs"].append(text)
        else:
            if current["paragraphs"]:
                sections.append(current)
            current = {"title": text, "paragraphs": []}
    if current["paragraphs"]:
        sections.append(current)
    return sections


def build_xml(source, kind, sections):
    root = ET.Element("document", source=source, type=kind)
    for n, sec in enumerate(sections, start=1):
        s = ET.SubElement(root, "section", number=str(n))
        ET.SubElement(s, "title").text = sec["title"]
        for p in sec["paragraphs"]:
            ET.SubElement(s, "paragraph").text = p
    ET.indent(root)
    return ET.tostring(root, encoding="unicode", xml_declaration=False)


def excel_rows(sections):
    """Same rows that go into the Excel file, used for the on-page preview."""
    return [(n, sec["title"], k, p)
            for n, sec in enumerate(sections, start=1)
            for k, p in enumerate(sec["paragraphs"], start=1)]


@app.route("/", methods=["GET", "POST"])
def home():
    ctx = {"error": None, "xml": None, "sections": None, "source": None, "rows": []}
    if request.method == "POST":
        try:
            pdf = request.files.get("pdf")
            url = request.form.get("url", "").strip()
            if pdf and pdf.filename:
                data = pdf.read()
                if len(data) > MAX_PDF_BYTES:
                    raise ValueError("PDF is larger than 10 MB.")
                if not pdf.filename.lower().endswith(".pdf"):
                    raise ValueError("Please upload a .pdf file.")
                sections = pdf_to_sections(io.BytesIO(data))
                kind, source = "pdf", pdf.filename
            elif url:
                if urlparse(url).scheme not in ("http", "https"):
                    raise ValueError("Link must start with http:// or https://")
                sections = web_to_sections(url)
                kind, source = "web", url
            else:
                raise ValueError("Upload a PDF or paste a website link.")
            if not sections:
                raise ValueError("No readable text found (scanned PDFs need OCR).")
            ctx.update(sections=sections, source=source, rows=excel_rows(sections),
                       xml='<?xml version="1.0" encoding="UTF-8"?>\n' + build_xml(source, kind, sections))
        except RequestEntityTooLarge:
            raise
        except requests.RequestException as e:
            ctx["error"] = f"Could not open the link: {e}"
        except Exception as e:
            ctx["error"] = str(e)
    return render_template("index.html", **ctx)


@app.route("/download", methods=["POST"])
def download():
    xml = request.form.get("xml", "")
    return Response(xml, mimetype="application/xml",
                    headers={"Content-Disposition": "attachment; filename=extracted.xml"})


@app.route("/download-excel", methods=["POST"])
def download_excel():
    """Rebuild the Excel file from the XML, so the server stays stateless."""
    root = ET.fromstring(request.form.get("xml", "").encode("utf-8"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"
    ws.append(["Section No", "Section Title", "Paragraph No", "Text"])
    for sec in root.findall("section"):
        title = sec.findtext("title", "")
        for n, p in enumerate(sec.findall("paragraph"), start=1):
            ws.append([int(sec.get("number")), title, n, p.text or ""])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
    for col, width in zip("ABCD", (12, 30, 14, 100)):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2):
        row[3].alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return Response(buf.getvalue(),
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment; filename=extracted.xlsx"})


@app.errorhandler(413)
def too_big(_):
    return render_template("index.html", error="File too large (max 10 MB).", xml=None, sections=None, source=None, rows=[]), 413


if __name__ == "__main__":
    app.run(debug=True)

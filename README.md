JSGRZQQKIHUIWMSIM7NNUZLTM7P5AP5D

"""
UK Companies House - SH01 / RP04SH01 / RP01SH01 extractor (text-layer PDFs)
---------------------------------------------------------------------------
Run  : python ch_sh01_extract.py
Input: Company number (asked at runtime)
Key  : set env var CH_API_KEY, otherwise the script asks for it (hidden input)
Output: <CompanyName>_<CompanyNumber>_test.xlsx in OUTPUT_DIR

Electronic SH01(ef) PDFs are parsed here. Scanned/paper PDFs are FLAGGED
(red row + console report). Use ch_sh01_extract_ocr.py to also OCR them.

pip install requests pdfplumber openpyxl
"""
import os
import re
import io
import time
import getpass
from datetime import datetime

import requests
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.worksheet.table import Table, TableStyleInfo

# ----------------------------- CONFIG ---------------------------------------
OUTPUT_DIR = r"lll"
API_BASE = "https://api.company-information.service.gov.uk"
TARGET_TYPES = {"SH01", "RP04SH01", "RP01SH01"}   # SH02 ignored for testing
VERIFY_SSL = True        # if Zscaler SSL error: set to path of corporate CA .pem
REQUEST_PAUSE = 0.6      # seconds between calls (limit = 600 req / 5 min)

HEADERS = ["Date of Filling", "Investment Date", "Share Class", "Number of Shares",
           "Share price", "Currency", "Amount sold", "File", "Filing Type", "PDF Link"]

RED = PatternFill("solid", fgColor="FFC7CE")      # filing missed
YELLOW = PatternFill("solid", fgColor="FFF2CC")   # OCR row - verify
FOOTER_RE = re.compile(r"Electronically filed document for Company Number:.*?Page:\s*\d+", re.I | re.S)


class ScannedPDF(Exception):
    """PDF has no usable text layer."""


class ParseError(Exception):
    """PDF structure not recognised by the parser."""


# ----------------------------- HELPERS --------------------------------------
def normalise_company_no(raw):
    s = re.sub(r"\s+", "", raw).upper()
    if s.isdigit():
        return s.zfill(8)
    m = re.fullmatch(r"([A-Z]{2})(\d+)", s)
    if m:
        return m.group(1) + m.group(2).zfill(6)
    return s


def parse_int(txt):
    digits = re.sub(r"[^\d]", "", txt or "")
    return int(digits) if digits else None


def parse_dec(txt):
    s = re.sub(r"[^\d.,]", "", txt or "")          # strips £ $ € and spaces
    if not s:
        return None
    if "," in s and "." not in s and not re.fullmatch(r"\d{1,3}(,\d{3})+", s):
        s = s.replace(",", ".")                     # comma used as decimal
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def to_date(ddmmyyyy):
    try:
        return datetime.strptime(ddmmyyyy, "%d/%m/%Y").date()
    except (ValueError, TypeError):
        return None


def long_date(iso):
    d = datetime.strptime(iso, "%Y-%m-%d")
    return f"{d.day} {d:%B %Y}"


DESC_TEMPLATES = {
    "capital-allotment-shares": "Statement of capital following an allotment of shares on {date}",
}


def render_description(item):
    key = item.get("description", "") or ""
    vals = item.get("description_values") or {}
    if vals.get("description"):                     # legacy / free-text filings
        return vals["description"]
    dtxt = long_date(vals["date"]) if vals.get("date") else ""
    if key in DESC_TEMPLATES:
        return DESC_TEMPLATES[key].format(date=dtxt).strip()
    text = key.replace("-", " ").capitalize()
    return f"{text} on {dtxt}" if dtxt else text


# ----------------------------- API CLIENT -----------------------------------
class CHClient:
    def __init__(self, api_key):
        self.s = requests.Session()
        self.s.auth = (api_key, "")
        self.s.verify = VERIFY_SSL

    def get(self, url, params=None, accept="application/json", redirects=True, retries=3):
        r = None
        for attempt in range(retries + 1):
            r = self.s.get(url, params=params, headers={"Accept": accept},
                           timeout=60, allow_redirects=redirects)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 60))
                print(f"   Rate limit hit, waiting {wait}s ...")
                time.sleep(wait)
                continue
            if r.status_code >= 500 and attempt < retries:
                time.sleep(2 * (attempt + 1))
                continue
            break
        time.sleep(REQUEST_PAUSE)
        return r

    def company(self, no):
        r = self.get(f"{API_BASE}/company/{no}")
        if r.status_code == 401:
            raise SystemExit("401 Unauthorised - check the API key.")
        if r.status_code == 404:
            raise SystemExit(f"Company {no} not found.")
        r.raise_for_status()
        return r.json()

    def filing_history(self, no):
        items, start = [], 0
        while True:
            r = self.get(f"{API_BASE}/company/{no}/filing-history",
                         params={"items_per_page": 100, "start_index": start})
            r.raise_for_status()
            data = r.json()
            batch = data.get("items", [])
            items.extend(batch)
            start += len(batch)
            if not batch or start >= data.get("total_count", 0):
                break
        return items

    def document_pdf(self, metadata_url):
        meta = self.get(metadata_url)
        meta.raise_for_status()
        content_url = meta.json()["links"]["document"]
        # Content URL redirects to a signed S3 URL - follow it WITHOUT auth
        r = self.get(content_url, accept="application/pdf", redirects=False)
        if r.status_code in (301, 302, 303, 307, 308):
            r = requests.get(r.headers["Location"], timeout=120, verify=VERIFY_SSL)
        r.raise_for_status()
        return r.content


# ----------------------------- ELECTRONIC PARSER ----------------------------
LEFT_RE = re.compile(r"Class of Shares:\s*(.*?)\s*Currency:\s*([A-Z]{3})", re.S)
RIGHT_RE = re.compile(r"Number allotted\s*([\d,]+)\s*Nominal value of each share\s*([\d.,]+)"
                      r"\s*Amount paid:\s*(\S+)", re.S)
LINEAR_RE = re.compile(r"Class of Shares:\s*(.*?)\s*Currency:\s*([A-Z]{3})\s*Number allotted\s*([\d,]+)"
                       r"\s*Nominal value of each share\s*([\d.,]+)\s*Amount paid:\s*(\S+)", re.S)


def _find_phrase(words, phrase):
    """Return first word dict that starts the given phrase (exact, case-sensitive)."""
    parts = phrase.split()
    for i in range(len(words) - len(parts) + 1):
        if all(words[i + j]["text"] == p for j, p in enumerate(parts)):
            return words[i]
    return None


def _from_date(full_text):
    seg = full_text.split("From", 1)
    if len(seg) < 2:
        return None
    m = re.search(r"\d{2}/\d{2}/\d{4}", seg[1])     # first date after "From" = From date
    return to_date(m.group(0)) if m else None


def parse_electronic(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full = "\n".join((p.extract_text() or "") for p in pdf.pages)
        if len(full.strip()) < 50:
            raise ScannedPDF("No text layer (scanned/paper filing)")
        if "Class of Shares" not in full or "Number allotted" not in full:
            raise ParseError("Text PDF but not the SH01(ef) layout")

        inv_date = _from_date(full)
        left_txt, right_txt = [], []
        split_x, started = None, False

        # Two-column layout: left = class/currency, right = numbers.
        # Crop each page between "Shares Allotted" header and "Statement of Capital",
        # excluding the "Electronically filed document..." footer.
        for page in pdf.pages:
            words = page.extract_words()
            top, bottom = 0, page.height

            hdr = _find_phrase(words, "Shares Allotted")
            if hdr:
                top, started = hdr["bottom"], True
            if not started:
                continue

            foot = _find_phrase(words, "Electronically filed document")
            if foot:
                bottom = min(bottom, foot["top"])
            stop = _find_phrase(words, "Statement of Capital")
            if stop:
                bottom = min(bottom, stop["top"])

            na = _find_phrase(words, "Number allotted")
            if na and na["top"] < bottom and split_x is None:
                split_x = na["x0"] - 2
            bottom -= 1                              # crop() includes objects touching the edge
            if split_x and bottom > top:
                left_txt.append(page.crop((0, top, split_x, bottom)).extract_text() or "")
                right_txt.append(page.crop((split_x, top, page.width, bottom)).extract_text() or "")
            if stop:
                break

        rows = []
        lefts = LEFT_RE.findall("\n".join(left_txt))
        rights = RIGHT_RE.findall("\n".join(right_txt))
        if lefts and len(lefts) == len(rights):
            for (cls, cur), (num, nom, paid) in zip(lefts, rights):
                rows.append((cls, cur, num, nom, paid))
        else:
            # Fallback: linear text (footer removed, Statement of Capital cut off)
            body = FOOTER_RE.sub(" ", full).split("Statement of Capital")[0]
            rows = LINEAR_RE.findall(body)
            if not rows:
                raise ParseError(f"Block mismatch: {len(lefts)} share classes vs "
                                 f"{len(rights)} number blocks")

    out = []
    for cls, cur, num, nom, paid in rows:
        rec = {"inv_date": inv_date, "share_class": " ".join(cls.split()),
               "number": parse_int(num), "price": parse_dec(paid), "currency": cur}
        if rec["number"] is None or rec["price"] is None:
            raise ParseError(f"Could not read number/price for class '{rec['share_class']}'")
        out.append(rec)
    if inv_date is None:
        raise ParseError("Allotment From date not found")
    return out


# ----------------------------- EXCEL ----------------------------------------
def write_excel(path, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "SH01"
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)

    for i, rec in enumerate(records, start=2):
        ws.cell(i, 1, rec["filing_date"]).number_format = "DD-MM-YYYY"
        ws.cell(i, 8, rec["file"])
        ws.cell(i, 9, rec["type"])
        link = ws.cell(i, 10, "View PDF")
        link.hyperlink = rec["link"]
        link.style = "Hyperlink"

        if rec.get("number") is not None:
            ws.cell(i, 2, rec["inv_date"]).number_format = "DD-MM-YYYY"
            ws.cell(i, 3, rec["share_class"])
            ws.cell(i, 4, rec["number"]).number_format = "#,##0"
            ws.cell(i, 5, rec["price"])
            ws.cell(i, 6, rec["currency"])
            ws.cell(i, 7, f"=D{i}*E{i}").number_format = "#,##0.00"

        fill = RED if rec["status"] in ("FAILED", "SCANNED") else YELLOW if rec["status"] == "OCR" else None
        if fill:
            for col in range(1, 11):
                ws.cell(i, col).fill = fill
            ws.cell(i, 1).comment = Comment(f"{rec['status']}: {rec['reason']}", "SH01 script")

    if len(records):
        tbl = Table(displayName="SH01Data", ref=f"A1:J{len(records) + 1}")
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium3", showRowStripes=False)
        ws.add_table(tbl)
    for col, w in zip("ABCDEFGHIJ", [14, 15, 30, 17, 14, 10, 16, 70, 12, 11]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top")

    try:
        wb.save(path)
    except PermissionError:                          # file open in Excel
        path = path.replace(".xlsx", f"_{datetime.now():%H%M%S}.xlsx")
        wb.save(path)
    return path


# ----------------------------- MAIN -----------------------------------------
def main(ocr_parser=None):
    api_key = os.environ.get("CH_API_KEY") or getpass.getpass("Companies House API key: ").strip()
    company_no = normalise_company_no(input("Enter Company Number: "))
    ch = CHClient(api_key)

    profile = ch.company(company_no)
    name = profile.get("company_name", company_no)
    print(f"\n{name} ({company_no})")

    filings = [f for f in ch.filing_history(company_no) if f.get("type") in TARGET_TYPES]
    filings.sort(key=lambda f: f.get("date", ""), reverse=True)
    print(f"Found {len(filings)} filing(s) of type {', '.join(sorted(TARGET_TYPES))}\n")

    records, report = [], []
    for f in filings:
        base = {
            "filing_date": datetime.strptime(f["date"], "%Y-%m-%d").date(),
            "type": f["type"],
            "file": render_description(f),
            "link": (f"https://find-and-update.company-information.service.gov.uk/company/"
                     f"{company_no}/filing-history/{f.get('transaction_id')}/document?format=pdf&download=0"),
        }
        status, reason, rows = "OK", "", []
        try:
            meta_url = (f.get("links") or {}).get("document_metadata")
            if not meta_url:
                raise ParseError("No document available via API")
            pdf_bytes = ch.document_pdf(meta_url)
            try:
                rows = parse_electronic(pdf_bytes)
            except ScannedPDF as e:
                if ocr_parser is None:
                    status, reason = "SCANNED", f"{e} - run ch_sh01_extract_ocr.py"
                else:
                    rows = ocr_parser(pdf_bytes)
                    status, reason = "OCR", "Read by OCR - verify against PDF"
        except ParseError as e:
            status, reason = "FAILED", str(e)
        except Exception as e:                       # network, corrupt PDF, etc.
            status, reason = "FAILED", f"{type(e).__name__}: {e}"

        if rows:
            records += [{**base, **r, "status": status, "reason": reason} for r in rows]
        else:
            records.append({**base, "status": status if status != "OK" else "FAILED",
                            "reason": reason or "No allotment rows found"})
        report.append((base["filing_date"], f["type"], status, len(rows), reason))

    out_name = re.sub(r'[\\/:*?"<>|]', "", f"{name}_{company_no}_test.xlsx")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = write_excel(os.path.join(OUTPUT_DIR, out_name), records)

    print(f"{'Filed':<12}{'Type':<10}{'Status':<9}{'Rows':<6}Reason")
    for d, t, s, n, why in report:
        print(f"{d:%d-%m-%Y}  {t:<10}{s:<9}{n:<6}{why}")
    missed = [r for r in report if r[2] in ("FAILED", "SCANNED")]
    if missed:
        print(f"\n!! {len(missed)} filing(s) MISSED - highlighted red in Excel (see cell comment in column A)")
    print(f"\nSaved: {path}")


if __name__ == "__main__":
    main()

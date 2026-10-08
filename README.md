"""
UK Companies House - SH01 extractor WITH OCR for scanned/paper SH01 forms
-------------------------------------------------------------------------
Keep this file in the SAME folder as ch_sh01_extract.py (it reuses it).
Electronic PDFs -> text parser. Scanned PDFs -> RapidOCR (rows highlighted yellow = verify).
If OCR cannot read the structure, the filing is flagged red as MISSED.

pip install requests pdfplumber openpyxl pymupdf rapidocr_onnxruntime numpy
"""
import re
from datetime import date

import fitz                      # PyMuPDF - renders PDF pages to images (no poppler needed)
import numpy as np

import ch_sh01_extract as base
from ch_sh01_extract import ParseError, parse_int, parse_dec

try:                             # rapidocr_onnxruntime (v1.x)
    from rapidocr_onnxruntime import RapidOCR
    _NEW_API = False
except ImportError:              # rapidocr (v2+/v3)
    from rapidocr import RapidOCR
    _NEW_API = True

DPI = 300
MAX_PAGES = 4                    # section 3 is normally on page 1-2
LOW_CONF = 0.80
_ENGINE = None


def _engine():
    global _ENGINE
    if _ENGINE is None:
        _ENGINE = RapidOCR()
    return _ENGINE


def _ocr(img):
    """Return list of boxes: dict(text, score, x0, x1, y0, y1, cx, cy)."""
    if _NEW_API:
        out = _engine()(img)
        raw = [] if out.boxes is None else list(zip(out.boxes, out.txts, out.scores))
    else:
        res, _ = _engine()(img)
        raw = res or []
    boxes = []
    for pts, txt, score in raw:
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        boxes.append({"text": str(txt).strip(), "score": float(score),
                      "x0": min(xs), "x1": max(xs), "y0": min(ys), "y1": max(ys),
                      "cx": (min(xs) + max(xs)) / 2, "cy": (min(ys) + max(ys)) / 2})
    return boxes


def _render(page):
    pix = page.get_pixmap(dpi=DPI)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    return img[:, :, :3].copy() if pix.n >= 3 else np.stack([img[:, :, 0]] * 3, axis=-1)


def _first(boxes, pattern, below=0):
    rx = re.compile(pattern, re.I)
    hits = [b for b in boxes if rx.search(b["text"]) and b["cy"] > below]
    return min(hits, key=lambda b: b["cy"]) if hits else None


def _read_from_date(boxes, page_w):
    anchor = _first(boxes, r"^from\s*date|^from$")
    if not anchor:
        raise ParseError("OCR: 'From Date' label not found")
    note = _first(boxes, r"allotment date", below=anchor["y0"] - 200)
    right_limit = note["x0"] if note and note["x0"] > anchor["x1"] else page_w * 0.70
    tol = (anchor["y1"] - anchor["y0"]) * 0.9
    same_line = sorted((b for b in boxes if abs(b["cy"] - anchor["cy"]) <= tol
                        and anchor["x1"] < b["cx"] < right_limit), key=lambda b: b["x0"])
    digits = re.sub(r"\D", "", "".join(b["text"] for b in same_line))
    # boxed digits + d/m/y superscripts -> look for ddmm20yy
    for m in re.finditer(r"(?=(\d{2})(\d{2})(20\d{2}))", digits):
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
    raise ParseError(f"OCR: could not read From date (got '{digits}')")


def _read_table(boxes, page_w, page_h, title):
    hdr = {
        "class": _first(boxes, r"class of shares", below=title["cy"]),
        "number": _first(boxes, r"number of shares", below=title["cy"]),
        "nominal": _first(boxes, r"nominal value", below=title["cy"]),
        "paid": _first(boxes, r"amount paid", below=title["cy"]),
        "unpaid": _first(boxes, r"amount \(if any\)|if any|unpaid", below=title["cy"]),
    }
    missing = [k for k in ("class", "number", "paid") if not hdr[k]]
    if missing:
        raise ParseError(f"OCR: table header(s) not found: {', '.join(missing)}")

    # column left edges (currency column = everything left of class column)
    pad = 15
    order = [("currency", 0)]
    for k in ("class", "number", "nominal", "paid", "unpaid"):
        if hdr[k]:
            order.append((k, hdr[k]["x0"] - pad))
    order.sort(key=lambda kv: kv[1])
    note_col = _first(boxes, r"^currency$|if currency details", below=title["cy"])
    right_edge = page_w
    if note_col and hdr["unpaid"] and note_col["x0"] > hdr["unpaid"]["x1"]:
        right_edge = note_col["x0"]

    def col_of(b):
        name = None
        for k, edge in order:
            if b["cx"] >= edge:
                name = k
        return name

    header_top = min(h["y0"] for h in hdr.values() if h)
    stop = _first(boxes, r"non-cash|statement of capital|details of", below=header_top + 50)
    bottom = stop["y0"] if stop else page_h

    cand = sorted((b for b in boxes if header_top < b["cy"] < bottom and b["cx"] < right_edge),
                  key=lambda b: b["cy"])
    # group boxes into rows by vertical position
    rows, tol = [], DPI * 0.06
    for b in cand:
        if rows and abs(b["cy"] - rows[-1][-1]["cy"]) <= tol:
            rows[-1].append(b)
        else:
            rows.append([b])

    out = []
    for r in rows:
        cols = {}
        for b in sorted(r, key=lambda b: b["x0"]):
            cols.setdefault(col_of(b), []).append(b)
        txt = {k: " ".join(b["text"] for b in v) for k, v in cols.items()}
        number = parse_int(txt.get("number"))
        price = parse_dec(txt.get("paid"))
        cls = (txt.get("class") or "").strip()
        if number is None or price is None or not cls:
            continue                                   # header text / blank form lines
        cur = re.sub(r"[^A-Z]", "", (txt.get("currency") or "").upper())[:3] or "GBP"  # form default = GBP
        low = min(b["score"] for b in r) < LOW_CONF
        out.append({"share_class": " ".join(cls.split()), "number": number,
                    "price": price, "currency": cur, "low_conf": low})
    return out


def parse_scanned(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for pno in range(min(MAX_PAGES, doc.page_count)):
        img = _render(doc[pno])
        h, w = img.shape[:2]
        boxes = _ocr(img)
        title = _first(boxes, r"^shares allotted")
        if not title or not _first(boxes, r"class of shares"):
            continue
        inv_date = _read_from_date(boxes, w)
        rows = _read_table(boxes, w, h, title)
        if not rows:
            raise ParseError(f"OCR: page {pno + 1} has section 3 but no readable allotment rows")
        if any([r.pop("low_conf") for r in rows]):
            print("   (low OCR confidence on some cells - check the yellow rows)")
        return [{**r, "inv_date": inv_date} for r in rows]
    raise ParseError(f"OCR: 'Shares allotted' section not found in first {MAX_PAGES} pages "
                     "(unknown paper layout)")


if __name__ == "__main__":
    base.main(ocr_parser=parse_scanned)

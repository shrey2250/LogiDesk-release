"""
Bill Generator service.

CRITICAL DESIGN NOTE (read before editing):
The transaction table in the generated bill is a genuine embedded Microsoft Excel
OLE object living inside the selected .docx template (word/embeddings/*.xlsx),
NOT a python-docx Word table. This module edits that embedded workbook in place
and regenerates its on-page preview image, while leaving every other part of the
template (layout, fonts, spacing, headers, footers, MSME table, etc.) untouched.

Do not replace this with a python-docx `doc.add_table(...)` implementation -- that
was tried before and explicitly rejected: the user wants the real OLE object that
opens in Excel when double-clicked in Word, not a visually similar Word table.
"""

import copy as copy_module
import io
import math
import os
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from docx import Document

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    etree = None

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except Exception:  # pragma: no cover
    openpyxl = None
    Alignment = None
    Border = None
    Font = None
    PatternFill = None
    Side = None

try:
    import win32com.client as win32  # Windows + Office only
    import win32clipboard
    import ctypes
except Exception:  # pragma: no cover
    win32 = None
    win32clipboard = None
    ctypes = None

try:
    from PIL import Image, ImageChops
except Exception:  # pragma: no cover
    Image = None
    ImageChops = None


NS_MAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'v': 'urn:schemas-microsoft-com:vml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
for prefix, uri in NS_MAP.items():
    ET.register_namespace(prefix, uri)

for prefix, uri in [
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('w15', 'http://schemas.microsoft.com/office/word/2012/wordml'),
    ('w16', 'http://schemas.microsoft.com/office/word/2018/wordml'),
    ('w16cid', 'http://schemas.microsoft.com/office/word/2016/wordml/cid'),
    ('w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex'),
    ('w16sdtdh', 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash'),
    ('w16sdtfl', 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock'),
    ('w16du', 'http://schemas.microsoft.com/office/word/2023/wordml/word16du'),
    ('w16cex', 'http://schemas.microsoft.com/office/word/2018/wordml/cex'),
    ('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'),
    ('cx', 'http://schemas.microsoft.com/office/drawing/2014/chartex'),
    ('cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'),
    ('cx2', 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex'),
    ('cx3', 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex'),
    ('cx4', 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex'),
    ('cx5', 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex'),
    ('cx6', 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex'),
    ('cx7', 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex'),
    ('cx8', 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex'),
    ('aink', 'http://schemas.microsoft.com/office/drawing/2016/ink'),
    ('am3d', 'http://schemas.microsoft.com/office/drawing/2017/model3d'),
    ('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'),
    ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('w10', 'urn:schemas-microsoft-com:office:word'),
    ('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'),
    ('wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk'),
    ('wne', 'http://schemas.microsoft.com/office/word/2006/wordml'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
    ('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math'),
    ('oel', 'http://schemas.microsoft.com/office/2019/extlst'),
]:
    ET.register_namespace(prefix, uri)


DEFAULT_BILL_SAC_CODE = "996601"


# Standard 15-character GSTIN pattern, used as a fallback to pull a GSTIN out
# of a mapping file's Address text when there is no dedicated GSTIN column.
_GSTIN_RE = re.compile(r"\b(\d{2}[A-Za-z]{5}\d{4}[A-Za-z]\d[Zz][A-Za-z0-9])\b")
_GSTIN_LABEL_RE = re.compile(r"GSTIN\s*[:\-]?\s*\d{2}[A-Za-z]{5}\d{4}[A-Za-z]\d[Zz][A-Za-z0-9]", re.IGNORECASE)
BILL_MARKER_OPTIONS = ["+", "*", "Blank"]


# ──────────────────────────────────────────────────────────────────────────────
# UTILITY FUNCTIONS
# ──────────────────────────────────────────────────────────────────────────────

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def parse_date(value: Any) -> "pd.Timestamp | None":
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, (int, float)):
        try:
            number = float(value)
            if not math.isnan(number) and 20000 <= number <= 100000:
                return pd.Timestamp("1899-12-30") + pd.to_timedelta(number, unit="D")
        except Exception:
            pass

    if isinstance(value, str):
        text = value.strip()
        if text:
            if re.fullmatch(r"\d+(?:\.\d+)?", text):
                try:
                    serial = float(text)
                    if 20000 <= serial <= 100000:
                        return pd.Timestamp("1899-12-30") + pd.to_timedelta(serial, unit="D")
                except Exception:
                    pass
            for fmt in [
                "%Y-%m-%d", "%Y/%m/%d",
                "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
                "%d.%m.%y", "%d.%m.%Y",
                "%d-%m-%y", "%d-%m-%Y",
                "%d/%m/%y", "%d/%m/%Y",
            ]:
                try:
                    return pd.Timestamp(pd.to_datetime(text, format=fmt))
                except Exception:
                    pass

    try:
        result = pd.to_datetime(value, errors="coerce")
        if pd.isna(result):
            return None
        return pd.Timestamp(result)
    except Exception:
        return None


def parse_amount(value: Any) -> "float | None":
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, (int, float)):
        try:
            number = float(value)
            return None if math.isnan(number) else number
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("₹", "").replace("Rs.", "").replace("Rs", "").replace("INR", "").strip()
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text)
    except Exception:
        return None


def format_lr(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, (int, float)):
        try:
            number = float(value)
            if math.isnan(number):
                return ""
            return str(int(number)) if number.is_integer() else str(number)
        except Exception:
            pass
    text = str(value).strip()
    if not text:
        return ""
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except Exception:
        pass
    return text


def format_indian_currency(amount: Any) -> str:
    """
    Format a number using the Indian numbering system:
      1000 -> 1,000
      10000 -> 10,000
      100000 -> 1,00,000
      931000 -> 9,31,000
      1250000 -> 12,50,000
      10000000 -> 1,00,00,000
    """
    if amount is None:
        return "0"
    try:
        f = float(str(amount).replace(",", "").strip())
    except (ValueError, TypeError):
        return str(amount)

    is_negative = f < 0
    f = abs(f)

    whole_part = int(f)
    has_decimals = (f != whole_part)

    s = str(whole_part)
    if len(s) <= 3:
        formatted_whole = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        groups = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        formatted_whole = ",".join(groups) + "," + last3

    if has_decimals:
        formatted_dec = f"{f:.2f}".split(".")[1]
        result = f"{formatted_whole}.{formatted_dec}"
    else:
        result = formatted_whole

    return f"-{result}" if is_negative else result


def format_wt_value(h_val: Any, i_val: Any) -> str:
    """
    WT extraction rules:
    IF H has a value:
        WT = H exactly as written.
    IF H is empty:
        Look at column I.
        IF I == 1:
            WT = "FIX"
        ELSE IF I contains a numeric value other than 1:
            WT = numeric value + "Mt" (e.g. 27 -> "27Mt", 8 -> "8Mt")
        ELSE:
            WT = blank
    """
    h_str = ""
    if h_val is not None:
        try:
            if not pd.isna(h_val):
                h_str = str(h_val).strip()
        except Exception:
            pass
    if h_str.lower() in {"nan", "none", "null"}:
        h_str = ""

    if h_str != "":
        try:
            f = float(h_str)
            if f.is_integer():
                return str(int(f))
            return str(f)
        except ValueError:
            return h_str

    i_str = ""
    if i_val is not None:
        try:
            if not pd.isna(i_val):
                i_str = str(i_val).strip()
        except Exception:
            pass
    if i_str.lower() in {"nan", "none", "null"}:
        i_str = ""

    if not i_str:
        return ""

    try:
        f = float(i_str)
        if f == 1.0:
            return "FIX"
        val = int(f) if f.is_integer() else f
        return f"{val}Mt"
    except ValueError:
        if i_str == "1":
            return "FIX"
        m = re.match(r"^(\d+(?:\.\d+)?)\s*(?:mt)?$", i_str, re.IGNORECASE)
        if m:
            num = float(m.group(1))
            val = int(num) if num.is_integer() else num
            return f"{val}Mt"
        return i_str


def format_rate_value(rate_val: Any) -> str:
    """
    Format rate value rounded to nearest 0.50 rupee:
      rounded_rate = round(rate * 2) / 2
    Examples:
      1481.48 -> '1481.50'
      1481.5 -> '1481.50'
      56000 -> '56000'
    """
    if rate_val is None:
        return ""
    try:
        if pd.isna(rate_val):
            return ""
    except Exception:
        pass
    rate_str = str(rate_val).strip()
    if rate_str.lower() in {"nan", "none", "null", ""}:
        return ""
    try:
        f = float(rate_str.replace(",", ""))
        rounded = round(f * 2) / 2
        if rounded.is_integer():
            return str(int(rounded))
        return f"{rounded:.2f}"
    except ValueError:
        return rate_str


def amount_to_words_inr(amount: float) -> str:
    amount = round(float(amount or 0), 2)
    whole = int(amount)
    if whole == 0:
        return "Rupee Zero Only."

    ones  = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
    teens = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen",
             "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens  = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def under_1000(n: int) -> str:
        if n < 10:   return ones[n]
        if n < 20:   return teens[n - 10]
        if n < 100:  return tens[n // 10] + (" " + ones[n % 10] if n % 10 else "")
        h, r = n // 100, n % 100
        return ones[h] + " Hundred" + (" " + under_1000(r) if r else "")

    lakh     = whole // 100000
    thousand = (whole % 100000) // 1000
    rest     = whole % 1000

    parts = []
    if lakh:     parts.append(under_1000(lakh) + " Lakh")
    if thousand: parts.append(under_1000(thousand) + " Thousand")
    if rest:     parts.append(under_1000(rest))
    return f"Rupee {' '.join(parts)} Only."


# ──────────────────────────────────────────────────────────────────────────────
# TEXT PLACEHOLDER REPLACEMENT (paragraphs / table cells outside the OLE object)
# ──────────────────────────────────────────────────────────────────────────────

def _replace_in_paragraph(paragraph, replacements: Dict[str, str]) -> None:
    """
    Replace [PLACEHOLDER] tokens in a paragraph while preserving formatting.
    Placeholders may be split across multiple runs (Word does this routinely);
    this merges run text, applies replacements, and writes the result back into
    the first run (keeping its formatting), clearing the rest.
    """
    if not paragraph.runs:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    if not any(key in full_text for key in replacements):
        return
    new_text = full_text
    for key, value in replacements.items():
        new_text = new_text.replace(key, value)
    paragraph.runs[0].text = new_text
    for run in paragraph.runs[1:]:
        run.text = ""


def _set_run_cambria(run_elem: Any) -> None:
    """Ensure run font is explicitly set to Cambria."""
    if etree is None:
        return
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    rPr = run_elem.find(f"{{{w_ns}}}rPr")
    if rPr is None:
        rPr = etree.Element(f"{{{w_ns}}}rPr")
        run_elem.insert(0, rPr)
    rFonts = rPr.find(f"{{{w_ns}}}rFonts")
    if rFonts is None:
        rFonts = etree.SubElement(rPr, f"{{{w_ns}}}rFonts")
    rFonts.set(f"{{{w_ns}}}ascii", "Cambria")
    rFonts.set(f"{{{w_ns}}}hAnsi", "Cambria")
    rFonts.set(f"{{{w_ns}}}cs", "Cambria")


def _replace_placeholders_in_lxml(tree: Any, replacements: Dict[str, str]) -> None:
    """
    Replace [PLACEHOLDER] tokens in all paragraphs and table cells inside an lxml
    parsed document tree while preserving formatting, and ensure generated/replaced
    text (To, Address, GSTIN, etc.) uses Cambria font.
    """
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    nsmap = {"w": w_ns}
    for p in list(tree.xpath(".//w:p", namespaces=nsmap)):
        runs = p.xpath("./w:r", namespaces=nsmap)
        if not runs:
            continue
        full_text = "".join("".join(r.xpath(".//w:t/text()", namespaces=nsmap)) for r in runs)
        
        # If paragraph is solely a placeholder whose replacement is empty, remove it to avoid blank lines
        if full_text.strip() in {"[COMPANY NAME]", "[ADDRESS LINE 2]", "[CITY]"} and replacements.get(full_text.strip(), "") == "":
            parent = p.getparent()
            if parent is not None:
                parent.remove(p)
                continue

        is_replaced = any(k in full_text for k in replacements)
        is_to_header = full_text.strip().startswith("To,") or full_text.strip() == "To"
        
        if is_replaced or is_to_header:
            new_text = full_text
            if is_replaced:
                for k, v in replacements.items():
                    new_text = new_text.replace(k, v)
                if "[ADDRESS LINE 1]" in full_text and replacements.get("[ADDRESS LINE 2]", "") == "":
                    new_text = new_text.rstrip(", ")
            first_t = runs[0].xpath(".//w:t", namespaces=nsmap)
            if first_t:
                first_t[0].text = new_text
                first_t[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            _set_run_cambria(runs[0])
            for r in runs[1:]:
                for t in r.xpath(".//w:t", namespaces=nsmap):
                    t.text = ""



# ──────────────────────────────────────────────────────────────────────────────
# EMBEDDED OLE SPREADSHEET DISCOVERY + EDITING
# ──────────────────────────────────────────────────────────────────────────────

class TemplateOleInfo:
    """Describes the embedded Excel OLE object found inside a .docx template."""

    def __init__(
        self,
        xlsx_part: str,
        sheet_name: str,
        header_row: int,
        header_map: Dict[str, str],   # normalized_header -> column_letter
        first_data_row: int,
        template_data_rows: int,      # number of blank data rows in the template (e.g. 3)
        total_row: Optional[int],
        total_cell: Optional[str],    # e.g. "G5" holding "[TOTAL]" or similar
        ole_size_ref: Optional[str],  # e.g. "A1:J5"
        image_rel_id: str,            # rId for the preview image (v:imagedata r:id)
        image_part: str,              # e.g. "word/media/image1.emf"
        shape_width_pt: Optional[float],
        shape_height_pt: Optional[float],
    ):
        self.xlsx_part = xlsx_part
        self.sheet_name = sheet_name
        self.header_row = header_row
        self.header_map = header_map
        self.first_data_row = first_data_row
        self.template_data_rows = template_data_rows
        self.total_row = total_row
        self.total_cell = total_cell
        self.ole_size_ref = ole_size_ref
        self.image_rel_id = image_rel_id
        self.image_part = image_part
        self.shape_width_pt = shape_width_pt
        self.shape_height_pt = shape_height_pt


# Column names we look for inside the embedded sheet's header row, and the row
# field each maps to. We NEVER invent a header that isn't already in the
# template -- we only fill data under whatever headers already exist.
# Header tokens recognized in the SOURCE Excel (not the embedded OLE
# workbook -- see _COLUMN_ALIASES below for that). Used both to locate the
# header row and, in _extract_headers_data, to absorb a column label that a
# workbook has placed on the row directly beneath the main header row.
_SOURCE_HEADER_TOKENS = {
    "senunit", "bill", "billc", "date", "lrno", "truckno", "town", "state",
    "wtmt", "remarks", "n", "total", "pay",
}

_COLUMN_ALIASES: Dict[str, List[str]] = {
    "sr_no":     ["srno", "sno", "sr"],
    "date":      ["date"],
    "lr_no":     ["lrno", "lrnumber"],
    "truck_no":  ["truckno", "trucknumber", "vehicleno"],
    "wt":        ["wt", "wtmt", "weight"],
    "rate":      ["rate"],
    "amount":    ["amount", "total"],
    "sac_code":  ["saccode", "sac"],
    "remarks":   ["remarks", "remark", "boe", "notes"],
}


def _col_letter(col_idx: int) -> str:
    letters = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _col_idx(col_letter: str) -> int:
    idx = 0
    for char in col_letter.upper():
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx



def find_ole_spreadsheet(docx_path: str) -> Optional[TemplateOleInfo]:
    """
    Inspect the .docx and locate its embedded Excel OLE object (if any), along
    with the header row / column layout that already exists inside it, and the
    preview-image relationship that renders it on the page.

    Returns None if the template has no embedded Excel OLE object.
    """
    if openpyxl is None:
        raise RuntimeError(
            "openpyxl is not installed. Install it with: pip install openpyxl"
        )

    with zipfile.ZipFile(docx_path, "r") as z:
        names = z.namelist()

        xlsx_parts = [n for n in names if n.startswith("word/embeddings/") and n.lower().endswith(".xlsx")]
        if not xlsx_parts:
            return None
        xlsx_part = xlsx_parts[0]

        xlsx_bytes = z.read(xlsx_part)

        # Find the relationship id + target image for the v:imagedata preview
        document_xml = z.read("word/document.xml").decode("utf-8")
        rels_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")

        # v:imagedata r:id="rIdX" appears right before the <o:OLEObject .../> tag
        m = re.search(r'<v:imagedata\s+r:id="(rId\d+)"[^/]*/>\s*</v:shape>\s*<o:OLEObject', document_xml)
        image_rel_id = m.group(1) if m else None

        image_part = None
        if image_rel_id:
            m2 = re.search(
                rf'<Relationship\s+Id="{image_rel_id}"[^>]*Target="([^"]+)"',
                rels_xml,
            )
            if m2:
                target = m2.group(1)
                image_part = "word/" + target if not target.startswith("word/") else target

        # Shape display size, e.g. style="width:478.5pt;height:75.75pt"
        shape_w = shape_h = None
        m3 = re.search(r'<v:shape[^>]*style="([^"]*)"', document_xml)
        if m3:
            style = m3.group(1)
            mw = re.search(r"width:([\d.]+)pt", style)
            mh = re.search(r"height:([\d.]+)pt", style)
            if mw:
                shape_w = float(mw.group(1))
            if mh:
                shape_h = float(mh.group(1))

    # Load the embedded workbook to discover header row / columns / total row
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    header_row = None
    header_map: Dict[str, str] = {}
    for r in range(1, min(20, ws.max_row) + 1):
        row_headers: Dict[str, str] = {}
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is None:
                continue
            norm = normalize_header(str(cell.value))
            for field, aliases in _COLUMN_ALIASES.items():
                if norm in aliases and field not in row_headers:
                    row_headers[field] = _col_letter(c)
        # A real header row should identify at least date + amount + one more
        if "date" in row_headers and "amount" in row_headers and len(row_headers) >= 3:
            header_row = r
            header_map = row_headers
            break

    if header_row is None:
        return None

    # Find first data row (row directly below header) and how many template
    # data rows exist before a TOTAL row / blank gap
    first_data_row = header_row + 1
    total_row = None
    total_cell = None
    r = first_data_row
    template_data_rows = 0
    while r <= ws.max_row + 1:
        # A "TOTAL" label anywhere in this row marks the total row
        row_texts = [clean_text(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if any(t.upper() == "TOTAL" for t in row_texts if t):
            total_row = r
            if "amount" in header_map:
                total_cell = f"{header_map['amount']}{r}"
            break
        # Stop scanning after a reasonable number of blank rows (template padding)
        if r - first_data_row > 50:
            break
        template_data_rows += 1
        r += 1

    # oleSize ref from xl/workbook.xml inside the embedded package
    ole_size_ref = None
    try:
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zx:
            wb_xml = zx.read("xl/workbook.xml").decode("utf-8")
            m4 = re.search(r'<oleSize\s+ref="([^"]+)"', wb_xml)
            if m4:
                ole_size_ref = m4.group(1)
    except Exception:
        pass

    return TemplateOleInfo(
        xlsx_part=xlsx_part,
        sheet_name=sheet_name,
        header_row=header_row,
        header_map=header_map,
        first_data_row=first_data_row,
        template_data_rows=max(template_data_rows, 1),
        total_row=total_row,
        total_cell=total_cell,
        ole_size_ref=ole_size_ref,
        image_rel_id=image_rel_id,
        image_part=image_part,
        shape_width_pt=shape_w,
        shape_height_pt=shape_h,
    )


def build_inr_words_formula(cell_ref: str) -> str:
    """
    Construct a pure native Excel formula that dynamically translates any numeric
    cell value into Indian numbering system words ("Rupees ... Only.").
    Supports numbers up to 99,99,99,999 (99 Crores) with no VBA macros required.
    """
    def n2(v: str) -> str:
        units = (
            'CHOOSE(' + v + ',"One","Two","Three","Four","Five","Six","Seven",'
            '"Eight","Nine","Ten","Eleven","Twelve","Thirteen","Fourteen",'
            '"Fifteen","Sixteen","Seventeen","Eighteen","Nineteen")'
        )
        tens = (
            'CHOOSE(INT(' + v + '/10)-1,"Twenty","Thirty","Forty","Fifty",'
            '"Sixty","Seventy","Eighty","Ninety")'
        )
        rem = (
            'IF(MOD(' + v + ',10)=0,""," "&CHOOSE(MOD(' + v + ',10),'
            '"One","Two","Three","Four","Five","Six","Seven","Eight","Nine"))'
        )
        return 'IF(' + v + '=0,"",IF(' + v + '<20,' + units + ',' + tens + '&' + rem + '))'

    rnd = 'ROUND(' + cell_ref + ',0)'
    cr = 'INT(MOD(' + rnd + '/10000000,100))'
    lk = 'INT(MOD(' + rnd + '/100000,100))'
    th = 'INT(MOD(' + rnd + '/1000,100))'
    hd = 'INT(MOD(' + rnd + '/100,10))'
    rt = 'INT(MOD(' + rnd + ',100))'

    f_cr = 'IF(' + cr + '>0,' + n2(cr) + '&" Crore ","")'
    f_lk = 'IF(' + lk + '>0,' + n2(lk) + '&" Lakh ","")'
    f_th = 'IF(' + th + '>0,' + n2(th) + '&" Thousand ","")'
    f_hd = 'IF(' + hd + '>0,CHOOSE(' + hd + ',"One","Two","Three","Four","Five","Six","Seven","Eight","Nine")&" Hundred ","")'
    f_rt = 'IF(' + rt + '>0,' + n2(rt) + ',"")'

    f_body = 'TRIM(' + f_cr + '&' + f_lk + '&' + f_th + '&' + f_hd + '&' + f_rt + ')'
    return '=IF(' + rnd + '=0,"Rupees Zero Only.","Rupees "&' + f_body + '&" Only.")'


def _fill_embedded_workbook(
    xlsx_bytes: bytes,
    ole_info: TemplateOleInfo,
    rows: List[Dict[str, Any]],
    subtotal: float,
) -> Tuple[bytes, str, int]:
    """
    Fill the embedded workbook's existing columns with the filtered transaction
    rows for a single route.

    Every route table has the exact structure:
      Header
      Data row 1
      Data row 2
      ...
      Data row N
      EMPTY ROW
      TOTAL ROW
      RUPEES IN WORDS ROW (Dynamic Formula)

    Returns (new_xlsx_bytes, new_ole_size_ref, total_needed_rows).
    """
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb[ole_info.sheet_name]

    hm = ole_info.header_map
    n_rows = len(rows)

    # 1. Capture data row styling from row 2
    data_styles = {}
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=ole_info.first_data_row, column=c)
        data_styles[c] = {
            "font": copy_module.copy(cell.font),
            "border": copy_module.copy(cell.border),
            "alignment": copy_module.copy(cell.alignment),
            "number_format": cell.number_format,
            "fill": copy_module.copy(cell.fill),
        }

    # 2. Capture total row styling & defaults from template total row (row 5)
    t_row_idx = ole_info.total_row or 5
    total_styles = {}
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=t_row_idx, column=c)
        total_styles[c] = {
            "font": copy_module.copy(cell.font),
            "border": copy_module.copy(cell.border),
            "alignment": copy_module.copy(cell.alignment),
            "number_format": cell.number_format,
            "fill": copy_module.copy(cell.fill),
            "value": cell.value,
        }

    max_c = max(max(data_styles.keys()), max(total_styles.keys()), 9)

    # Row layout:
    # Row 1: Header
    # Rows 2 .. 1 + n_rows: Data rows
    # Row 1 + n_rows + 1: Empty row
    # Row 1 + n_rows + 2: Total row
    # Row 1 + n_rows + 3: Rupees in Words row
    empty_row_idx = 1 + n_rows + 1
    total_row_idx = 1 + n_rows + 2
    words_row_idx = 1 + n_rows + 3
    total_needed_rows = words_row_idx

    # If ws currently has more rows than total_needed_rows, delete extra rows
    if ws.max_row > total_needed_rows:
        ws.delete_rows(total_needed_rows + 1, ws.max_row - total_needed_rows)

    # Populate Data rows
    for i in range(n_rows):
        r = ole_info.first_data_row + i
        row_data = rows[i]
        for c in range(1, max_c + 1):
            cell = ws.cell(row=r, column=c)
            st = data_styles.get(c, data_styles.get(1))
            if st:
                cell.font = copy_module.copy(st["font"])
                cell.border = copy_module.copy(st["border"])
                cell.alignment = copy_module.copy(st["alignment"])
                cell.number_format = st["number_format"]
                cell.fill = copy_module.copy(st["fill"])
            cell.value = None

        if "sr_no" in hm:
            ws[f"{hm['sr_no']}{r}"] = i + 1
        if "date" in hm:
            ws[f"{hm['date']}{r}"] = row_data.get("date_text", "")
        if "lr_no" in hm:
            ws[f"{hm['lr_no']}{r}"] = row_data.get("lr_no", "")
        if "truck_no" in hm:
            ws[f"{hm['truck_no']}{r}"] = row_data.get("truck_no", "")
        if "wt" in hm:
            cell_wt = ws[f"{hm['wt']}{r}"]
            cell_wt.value = row_data.get("wt_value", "")
            if Alignment is not None:
                vert = cell_wt.alignment.vertical if cell_wt.alignment else None
                cell_wt.alignment = Alignment(horizontal="center", vertical=vert)
        if "rate" in hm:
            cell_rate = ws[f"{hm['rate']}{r}"]
            rate_val = row_data.get("rate", "")
            if isinstance(rate_val, (int, float)):
                cell_rate.value = int(rate_val) if float(rate_val).is_integer() else rate_val
            elif isinstance(rate_val, str) and rate_val.strip().replace(".", "", 1).isdigit():
                try:
                    num_r = float(rate_val.strip())
                    cell_rate.value = int(num_r) if num_r.is_integer() else num_r
                except Exception:
                    cell_rate.value = rate_val
            else:
                cell_rate.value = rate_val
            if Alignment is not None:
                vert = cell_rate.alignment.vertical if cell_rate.alignment else None
                cell_rate.alignment = Alignment(horizontal="center", vertical=vert)
        if "amount" in hm:
            cell_amt = ws[f"{hm['amount']}{r}"]
            amt = row_data.get("amount", "")
            if isinstance(amt, (int, float)):
                cell_amt.value = int(amt) if float(amt).is_integer() else amt
            elif isinstance(amt, str) and amt.strip().replace(".", "", 1).isdigit():
                try:
                    num_a = float(amt.strip())
                    cell_amt.value = int(num_a) if num_a.is_integer() else num_a
                except Exception:
                    cell_amt.value = amt
            else:
                cell_amt.value = amt
            if Alignment is not None:
                vert = cell_amt.alignment.vertical if cell_amt.alignment else None
                cell_amt.alignment = Alignment(horizontal="center", vertical=vert)
        if "sac_code" in hm:
            ws[f"{hm['sac_code']}{r}"] = row_data.get("sac_code", DEFAULT_BILL_SAC_CODE)
        if "remarks" in hm:
            ws[f"{hm['remarks']}{r}"] = row_data.get("remarks", "")

    # Populate Empty Row (completely blank cells, but with data row table borders)
    for c in range(1, max_c + 1):
        cell = ws.cell(row=empty_row_idx, column=c)
        st = data_styles.get(c, data_styles.get(1))
        if st:
            cell.font = copy_module.copy(st["font"])
            cell.border = copy_module.copy(st["border"])
            cell.alignment = copy_module.copy(st["alignment"])
            cell.number_format = st["number_format"]
            cell.fill = copy_module.copy(st["fill"])
        cell.value = None

    # Populate Total Row
    st_val = subtotal if subtotal is not None else 0
    val_to_set = 0 if st_val == 0 else (int(st_val) if float(st_val).is_integer() else round(st_val, 2))

    for c in range(1, max_c + 1):
        cell = ws.cell(row=total_row_idx, column=c)
        t_st = total_styles.get(c, total_styles.get(1))
        if t_st:
            cell.font = copy_module.copy(t_st["font"])
            cell.border = copy_module.copy(t_st["border"])
            cell.alignment = copy_module.copy(t_st["alignment"])
            cell.number_format = t_st["number_format"]
            cell.fill = copy_module.copy(t_st["fill"])
            cell.value = t_st["value"]
        else:
            cell.value = None

    wt_col = hm.get("wt", "E")
    cell_tot_wt = ws[f"{wt_col}{total_row_idx}"]
    cell_tot_wt.value = "TOTAL"
    if Alignment is not None:
        vert = cell_tot_wt.alignment.vertical if cell_tot_wt.alignment else None
        cell_tot_wt.alignment = Alignment(horizontal="center", vertical=vert)

    rate_col = hm.get("rate", "F")
    if rate_col:
        cell_tot_rate = ws[f"{rate_col}{total_row_idx}"]
        if Alignment is not None:
            vert = cell_tot_rate.alignment.vertical if cell_tot_rate.alignment else None
            cell_tot_rate.alignment = Alignment(horizontal="center", vertical=vert)

    amt_col = hm.get("amount", "G")
    cell_tot_amt = ws[f"{amt_col}{total_row_idx}"]
    sum_formula = f"=SUM({amt_col}{ole_info.first_data_row}:{amt_col}{total_row_idx - 1})"
    cell_tot_amt.value = sum_formula
    if Alignment is not None:
        vert = cell_tot_amt.alignment.vertical if cell_tot_amt.alignment else None
        cell_tot_amt.alignment = Alignment(horizontal="center", vertical=vert)

    for c in range(1, max_c + 1):
        cell = ws.cell(row=total_row_idx, column=c)
        if cell.value is not None and str(cell.value).strip() == "[TOTAL]":
            cell.value = sum_formula

    # Determine active columns from header map
    max_header_idx = max(_col_idx(col_let) for col_let in hm.values()) if hm else 9
    last_col_letter = _col_letter(max_header_idx)

    # Populate Rupees in Words row (Merged across columns A..last_col_letter)
    words_formula = build_inr_words_formula(f"{amt_col}{total_row_idx}")

    first_cell_border = data_styles.get(1, {}).get("border")
    for c in range(1, max_c + 1):
        cell = ws.cell(row=words_row_idx, column=c)
        if first_cell_border:
            cell.border = copy_module.copy(first_cell_border)
        cell.value = None

    cell_words = ws.cell(row=words_row_idx, column=1)
    cell_words.value = words_formula
    if Font is not None:
        cell_words.font = Font(name="Cambria", size=9.5, bold=False)
    if Alignment is not None:
        cell_words.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.merge_cells(start_row=words_row_idx, start_column=1, end_row=words_row_idx, end_column=max_header_idx)

    # Explicit row heights so Excel OLE server matches Word's exact layout
    ws.row_dimensions[1].height = 17.25
    for r_idx in range(ole_info.first_data_row, total_row_idx + 1):
        ws.row_dimensions[r_idx].height = 14.5
    ws.row_dimensions[words_row_idx].height = 16.5

    # Clean up any unused/hidden phantom columns beyond last_col_letter (e.g. template column J)
    for c_idx in list(ws.column_dimensions.keys()):
        if _col_idx(str(c_idx)) > max_header_idx:
            del ws.column_dimensions[c_idx]

    # Ensure sheet view has standard 100% zoom
    if hasattr(ws, "sheet_view") and ws.sheet_view:
        ws.sheet_view.zoomScale = 100
        ws.sheet_view.zoomScaleNormal = 100

    new_ole_size_ref = f"A1:{last_col_letter}{total_needed_rows}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    with zipfile.ZipFile(buf, "r") as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}

    wb_xml = parts["xl/workbook.xml"].decode("utf-8")
    if "<oleSize" in wb_xml:
        wb_xml = re.sub(
            r'<oleSize\s+ref="[^"]*"\s*/>',
            f'<oleSize ref="{new_ole_size_ref}"/>',
            wb_xml,
        )
    else:
        wb_xml = wb_xml.replace("</workbook>", f'<oleSize ref="{new_ole_size_ref}"/></workbook>')
    parts["xl/workbook.xml"] = wb_xml.encode("utf-8")

    # Suppress all green triangle error indicators ("Number stored as text", evalError, etc.)
    for s_name in [k for k in parts if k.startswith("xl/worksheets/sheet") and k.endswith(".xml")]:
        s_xml = parts[s_name].decode("utf-8")
        if "<ignoredErrors>" not in s_xml:
            ignored_block = (
                f'<ignoredErrors>'
                f'<ignoredError sqref="A1:{last_col_letter}{total_needed_rows + 50}" '
                f'numberStoredAsText="1" evalError="1" formula="1" twoDigitTextYear="1"/>'
                f'</ignoredErrors>'
            )
            if "</worksheet>" in s_xml:
                insert_match = re.search(r'(</worksheet>)', s_xml)
                if insert_match:
                    s_xml = s_xml[:insert_match.start()] + ignored_block + s_xml[insert_match.start():]
            parts[s_name] = s_xml.encode("utf-8")

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)

    return out_buf.getvalue(), new_ole_size_ref, total_needed_rows


# ──────────────────────────────────────────────────────────────────────────────
# PREVIEW IMAGE REGENERATION (what Word displays for the OLE object on the page)
# ──────────────────────────────────────────────────────────────────────────────

def _render_preview_via_win32(
    xlsx_path: str,
    sheet_name: str,
    cell_range: str,
    out_emf_path: str
) -> Tuple[bool, Optional[float], Optional[float]]:
    """
    Primary path (Windows + Microsoft Office installed): use Excel COM automation
    to copy the exact cell range as an Enhanced Metafile (EMF) via the Windows
    clipboard, saving the native vector EMF image to out_emf_path.
    Returns (success, measured_width_pt, measured_height_pt).
    """
    if win32 is None or win32clipboard is None:
        return False, None, None
    excel = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        try:
            ws = wb.Sheets(sheet_name)
            rng = ws.Range(cell_range)
            measured_w = float(rng.Width)
            measured_h = float(rng.Height)

            # xlScreen=1, xlPicture=-4147 (copies high-fidelity vector EMF to clipboard)
            rng.CopyPicture(Appearance=1, Format=-4147)

            win32clipboard.OpenClipboard()
            try:
                # CF_ENHMETAFILE = 14
                hemf = win32clipboard.GetClipboardData(14)
                if isinstance(hemf, bytes):
                    emf_bytes = hemf
                else:
                    gdi32 = ctypes.windll.gdi32
                    buf_size = gdi32.GetEnhMetaFileBits(hemf, 0, None)
                    if buf_size <= 0:
                        return False, None, None
                    buf = (ctypes.c_char * buf_size)()
                    gdi32.GetEnhMetaFileBits(hemf, buf_size, buf)
                    emf_bytes = bytes(buf)

                if not emf_bytes or len(emf_bytes) == 0:
                    return False, None, None

                with open(out_emf_path, "wb") as f:
                    f.write(emf_bytes)

                ok = os.path.exists(out_emf_path) and os.path.getsize(out_emf_path) > 0
                return (ok, measured_w, measured_h) if ok else (False, None, None)
            finally:
                win32clipboard.CloseClipboard()
        finally:
            wb.Close(SaveChanges=False)
    except Exception:
        return False, None, None
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            except Exception:
                pass


def _render_preview_via_libreoffice(xlsx_path: str, cell_range: str, out_png_path: str) -> bool:
    """
    Fallback path (no Microsoft Office / non-Windows): use headless LibreOffice
    to render the workbook's print area to PDF, then rasterize + auto-crop that
    page to the content bounding box as a PNG preview image.
    Returns True on success, False if LibreOffice/Pillow are unavailable or the
    conversion failed.
    """
    if Image is None:
        return False

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        return False

    with tempfile.TemporaryDirectory() as tmp:
        # Set print area to the exact oleSize range, hide gridlines, save a temp copy
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb[wb.sheetnames[0]]
        ws.print_area = cell_range
        ws.sheet_view.showGridLines = False
        ws.print_options.gridLines = False
        tmp_xlsx = os.path.join(tmp, "preview_source.xlsx")
        wb.save(tmp_xlsx)

        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, tmp_xlsx],
                capture_output=True, timeout=60, check=True,
            )
        except Exception:
            return False

        pdf_path = os.path.join(tmp, "preview_source.pdf")
        if not os.path.exists(pdf_path):
            return False

        pdftocairo = shutil.which("pdftocairo")
        raw_png_prefix = os.path.join(tmp, "raw")
        if pdftocairo:
            try:
                subprocess.run(
                    [pdftocairo, "-png", "-r", "200", "-singlefile", pdf_path, raw_png_prefix],
                    capture_output=True, timeout=30, check=True,
                )
            except Exception:
                pass

        raw_png = raw_png_prefix + ".png"
        if not os.path.exists(raw_png):
            # Fall back to LibreOffice's own PNG export of the whole sheet
            try:
                subprocess.run(
                    [soffice, "--headless", "--convert-to", "png", "--outdir", tmp, tmp_xlsx],
                    capture_output=True, timeout=60, check=True,
                )
                raw_png = os.path.join(tmp, "preview_source.png")
            except Exception:
                return False

        if not os.path.exists(raw_png):
            return False

        img = Image.open(raw_png)
        gray = img.convert("L")
        bg = Image.new("L", gray.size, 255)
        diff = ImageChops.difference(gray, bg)
        bbox = diff.getbbox()
        if bbox:
            pad = 5
            x0, y0, x1, y1 = bbox
            x0 = max(0, x0 - pad)
            y0 = max(0, y0 - pad)
            x1 = min(img.width, x1 + pad)
            y1 = min(img.height, y1 + pad)
            img = img.crop((x0, y0, x1, y1))
        img.save(out_png_path)
        return True


def _splice_docx_parts(docx_path: str, replacements: Dict[str, bytes]) -> None:
    """
    Replace / add parts inside a .docx zip archive in place. `replacements` maps
    zip member name -> new bytes. If a member doesn't already exist it is added.
    """
    with zipfile.ZipFile(docx_path, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}

    parts.update(replacements)

    tmp_path = docx_path + ".tmp"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)

    os.replace(tmp_path, docx_path)


def _grow_ole_display_box(docx_path: str, ole_info: TemplateOleInfo, extra_rows: int) -> None:
    """
    When extra rows were inserted into the embedded spreadsheet (more
    transactions than the template's original capacity), grow the OLE
    object's on-page display box height proportionally, so the newly added
    rows are visible instead of being cropped/squashed into the template's
    original fixed-height box. Width and everything else about the template
    layout is left untouched -- only the height of this one object's frame
    changes, and only when it must.

    This edits two places in word/document.xml that both encode the same
    display size and must stay in sync:
      - <w:object w:dxaOrig="..." w:dyaOrig="...">   (twentieths of a point)
      - <v:shape style="width:...pt;height:...pt">   (points)
    """
    if extra_rows <= 0 or not ole_info.shape_height_pt:
        return

    n_template_rows = max(ole_info.template_data_rows, 1)
    total_rows_before = n_template_rows + 1  # + header row region already in the box
    total_rows_after = total_rows_before + extra_rows
    growth_factor = total_rows_after / total_rows_before

    new_height_pt = ole_info.shape_height_pt * growth_factor

    with zipfile.ZipFile(docx_path, "r") as z:
        document_xml = z.read("word/document.xml").decode("utf-8")

    # Update the v:shape style height (keep width unchanged)
    def _bump_style(m: "re.Match[str]") -> str:
        style = re.sub(r"height:[\d.]+pt", f"height:{new_height_pt:.2f}pt", m.group(1))
        return f'style="{style}"'

    document_xml_new = re.sub(
        r'style="([^"]*)"(?=[^>]*/>\s*<v:imagedata|[^>]*>\s*<v:imagedata)',
        _bump_style,
        document_xml,
        count=1,
    )
    if document_xml_new == document_xml:
        # Fallback: the shape tag's style may not be immediately followed by
        # v:imagedata in every template variant -- just target the first
        # v:shape style attribute in the document.
        document_xml_new = re.sub(
            r'(<v:shape[^>]*\bstyle=")([^"]*)(")',
            lambda m: m.group(1) + re.sub(r"height:[\d.]+pt", f"height:{new_height_pt:.2f}pt", m.group(2)) + m.group(3),
            document_xml,
            count=1,
        )

    # Update w:dyaOrig (dxaOrig/dyaOrig are in twentieths of a point = pt * 20)
    if ole_info.shape_height_pt:
        new_dya = int(round(new_height_pt * 20))
        document_xml_new = re.sub(
            r'(w:dyaOrig=")\d+(")',
            rf"\g<1>{new_dya}\g<2>",
            document_xml_new,
            count=1,
        )

    if document_xml_new != document_xml:
        _splice_docx_parts(docx_path, {"word/document.xml": document_xml_new.encode("utf-8")})


def _update_ole_preview_image(
    docx_path: str,
    ole_info: TemplateOleInfo,
    new_xlsx_bytes: bytes,
    new_ole_size_ref: str,
    extra_rows: int = 0,
) -> None:
    """
    Regenerate the on-page preview image for the OLE object so what's displayed
    in Word matches the new data, without altering the object's display box size
    on the page (so surrounding layout/spacing is preserved per requirement #8),
    unless more rows were added than the template originally had room for -- in
    that case the display height grows proportionally (via _grow_ole_display_box)
    so text isn't squashed, matching what Excel/Word do natively when an OLE
    object's range grows.
    """
    if extra_rows > 0:
        _grow_ole_display_box(docx_path, ole_info, extra_rows)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_xlsx = os.path.join(tmp, "embedded.xlsx")
        with open(tmp_xlsx, "wb") as f:
            f.write(new_xlsx_bytes)

        replacements: Dict[str, bytes] = {
            ole_info.xlsx_part: new_xlsx_bytes,
        }

        produced_preview = False
        preview_target_part = ole_info.image_part

        # Primary: native Windows/Excel COM rendering (pixel-perfect EMF)
        if win32 is not None and ole_info.image_part and ole_info.image_part.lower().endswith(".emf"):
            tmp_emf = os.path.join(tmp, "preview.emf")
            ok, _, _ = _render_preview_via_win32(tmp_xlsx, ole_info.sheet_name, new_ole_size_ref, tmp_emf)
            if ok:
                with open(tmp_emf, "rb") as f:
                    replacements[ole_info.image_part] = f.read()
                produced_preview = True

        # Fallback: LibreOffice + raster crop, saved as PNG (swap the image part's
        # extension/content-type registration rather than faking an EMF).
        if not produced_preview:
            tmp_png = os.path.join(tmp, "preview.png")
            ok = _render_preview_via_libreoffice(tmp_xlsx, new_ole_size_ref, tmp_png)
            if ok:
                with open(tmp_png, "rb") as f:
                    png_bytes = f.read()

                if ole_info.image_part:
                    new_image_part = re.sub(r"\.(emf|wmf|png|jpg|jpeg)$", ".png", ole_info.image_part, flags=re.I)
                    if new_image_part == ole_info.image_part and not new_image_part.lower().endswith(".png"):
                        new_image_part = ole_info.image_part + ".png"

                    if new_image_part != ole_info.image_part:
                        # Remove old part, add new one, and repoint the relationship
                        with zipfile.ZipFile(docx_path, "r") as zin:
                            rels_xml = zin.read("word/_rels/document.xml.rels").decode("utf-8")
                            ct_xml = zin.read("[Content_Types].xml").decode("utf-8")

                        old_target = ole_info.image_part[len("word/"):] if ole_info.image_part.startswith("word/") else ole_info.image_part
                        new_target = new_image_part[len("word/"):] if new_image_part.startswith("word/") else new_image_part
                        rels_xml = rels_xml.replace(f'Target="{old_target}"', f'Target="{new_target}"')

                        if 'Extension="png"' not in ct_xml:
                            ct_xml = ct_xml.replace(
                                "<Default Extension=",
                                '<Default Extension="png" ContentType="image/png"/><Default Extension=',
                                1,
                            )

                        replacements["word/_rels/document.xml.rels"] = rels_xml.encode("utf-8")
                        replacements["[Content_Types].xml"] = ct_xml.encode("utf-8")
                        # Drop the old image part, add the new one
                        with zipfile.ZipFile(docx_path, "r") as zin:
                            existing_names = set(zin.namelist())
                        if ole_info.image_part in existing_names:
                            # Mark old part for removal by rebuilding without it
                            _remove_docx_part(docx_path, ole_info.image_part)
                        replacements[new_image_part] = png_bytes
                        preview_target_part = new_image_part
                    else:
                        replacements[ole_info.image_part] = png_bytes
                produced_preview = True

        if not produced_preview:
            # Last resort: keep the old preview image as-is (data is still correct
            # inside the OLE object; only the on-page picture won't reflect it
            # until the user double-clicks / updates the object in Word).
            pass

        _splice_docx_parts(docx_path, replacements)


def _remove_docx_part(docx_path: str, part_name: str) -> None:
    with zipfile.ZipFile(docx_path, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist() if name != part_name}
    tmp_path = docx_path + ".tmp"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)
    os.replace(tmp_path, docx_path)


# ──────────────────────────────────────────────────────────────────────────────
# SERVICE
# ──────────────────────────────────────────────────────────────────────────────

class BillGeneratorService:
    def __init__(self):
        self.company_mapping: Dict[str, Dict[str, Any]] = {}
        self.special_routes: Dict[str, Tuple[str, str]] = {}
        self.default_mapping_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "company_mapping.xlsx"
        )

    @staticmethod
    def check_word_dependencies() -> None:
        missing = []
        try:
            import docx  # noqa: F401
        except Exception:
            missing.append("python-docx")
        try:
            import openpyxl  # noqa: F401
        except Exception:
            missing.append("openpyxl")
        if missing:
            raise RuntimeError(
                "Missing required packages: " + ", ".join(missing) +
                ". Install with: pip install " + " ".join(missing)
            )

    @staticmethod
    def normalize_company_name(value: Any) -> str:
        return re.sub(r"\s+", " ", clean_text(value)).casefold()

    @staticmethod
    def normalize_route_key(value: Any) -> str:
        text = clean_text(value).strip().casefold()
        if not text:
            return ""
        # Normalize multiple whitespace to single space
        text = re.sub(r"\s+", " ", text)
        # Remove whitespace around common punctuation: -, +, ,, /, . without altering punctuation
        text = re.sub(r"\s*([,\-+/.])\s*", r"\1", text)
        return text.strip()

    @staticmethod
    def normalize_bill_marker(marker: str) -> str:
        value = clean_text(marker)
        if value.lower() == "blank":  return "Blank"
        if value in {"+", "*"}:       return value
        if value.strip() in {"+", "*"}: return value.strip()
        return "Blank"

    @staticmethod
    def normalize_bill_value(value: Any) -> str:
        return clean_text(value).strip()

    @staticmethod
    def matches_bill_marker(value: Any, requested_marker: str) -> bool:
        norm = BillGeneratorService.normalize_bill_marker(requested_marker)
        actual = BillGeneratorService.normalize_bill_value(value)
        compact = re.sub(r"\s+", "", actual)
        if norm == "Blank":
            return actual == "" or actual.lower() in {"blank", "null", "n/a", "na"} or compact == ""
        if actual in {"", "Blank", "blank"} or compact == "":
            return False
        if norm in {"+", "*"}:
            return compact == norm
        return actual.casefold() == norm.casefold() or compact.casefold() == norm.casefold()

    @staticmethod
    def resolve_header(headers: List[str], *aliases: str) -> "str | None":
        alias_set = {normalize_header(a) for a in aliases}
        for h in headers:
            if normalize_header(h) in alias_set:
                return h
        return None

    def ensure_default_mapping(self) -> str:
        if os.path.exists(self.default_mapping_path):
            return self.default_mapping_path
        pd.DataFrame([{
            "Company Name": "F S INTERNATIONAL",
            "Address": "Plot No.341,\nUdhyog Vihar, Phase VI,\nGURGAON",
            "GSTIN": "06AAEFF0548A1Z2",
            "Origin(s)": "MUNDRA (GJ)",
            "SAC Code": DEFAULT_BILL_SAC_CODE,
        }]).to_excel(self.default_mapping_path, index=False)
        return self.default_mapping_path

    def load_mapping_file(self, mapping_path: str) -> Dict[str, Dict[str, Any]]:
        """
        Load mapping Excel/CSV. Supports:
          - Sheet 1: Company mapping (single 'Address' column or legacy 3-column)
          - Sheet 2: Special Route mapping (Special Route, Route Origin, Route Destination)
        """
        self.special_routes = {}
        if not mapping_path or not os.path.exists(mapping_path):
            if os.path.exists(self.default_mapping_path):
                mapping_path = self.default_mapping_path
            else:
                return {}

        is_csv = mapping_path.lower().endswith(".csv")
        sheet_names = []
        if not is_csv:
            try:
                with pd.ExcelFile(mapping_path) as wb:
                    sheet_names = wb.sheet_names
            except Exception:
                sheet_names = []

        if is_csv or not sheet_names:
            df = pd.read_csv(mapping_path, dtype=object) if is_csv \
                 else pd.read_excel(mapping_path, dtype=object)
        else:
            df = pd.read_excel(mapping_path, sheet_name=sheet_names[0], dtype=object)
            if len(sheet_names) > 1:
                spec_sheet = None
                for s in sheet_names[1:]:
                    s_low = s.strip().lower()
                    if any(k in s_low for k in ["special", "route", "sheet2", "sheet 2"]):
                        spec_sheet = s
                        break
                if spec_sheet is None:
                    spec_sheet = sheet_names[1]

                try:
                    df_spec = pd.read_excel(mapping_path, sheet_name=spec_sheet, dtype=object)
                except Exception:
                    df_spec = pd.DataFrame()

                if df_spec is not None and not df_spec.empty:
                    spec_cols = [clean_text(col) for col in df_spec.columns]
                    df_spec.columns = spec_cols
                    spec_norm = {normalize_header(col): col for col in spec_cols}

                    def pick_spec(*names: str) -> "str | None":
                        for name in names:
                            if normalize_header(name) in spec_norm:
                                return spec_norm[normalize_header(name)]
                        return None

                    spec_col = pick_spec("Special Route", "Special Routes", "SpecialRoute", "Special_Route", "Town", "Special", "Route Name", "Route")
                    orig_col = pick_spec("Route Origin", "RouteOrigin", "Route_Origin", "Origin", "From", "Source", "From City", "Origin City")
                    dest_col = pick_spec("Route Destination", "RouteDestination", "Route_Destination", "Destination", "Dest", "To", "To City", "City", "Town")

                    c0 = spec_cols[0] if len(spec_cols) > 0 else None
                    c1 = spec_cols[1] if len(spec_cols) > 1 else None
                    c2 = spec_cols[2] if len(spec_cols) > 2 else None

                    use_spec = spec_col or c0
                    use_orig = orig_col or c1
                    use_dest = dest_col or c2

                    spec_map: Dict[str, Tuple[str, str]] = {}
                    for _, srow in df_spec.iterrows():
                        raw_spec = clean_text(srow.get(use_spec, "")) if use_spec else ""
                        raw_orig = clean_text(srow.get(use_orig, "")) if use_orig else ""
                        raw_dest = clean_text(srow.get(use_dest, "")) if use_dest else ""
                        if not raw_spec:
                            continue
                        k = self.normalize_route_key(raw_spec)
                        if k:
                            spec_map[k] = (raw_orig, raw_dest)
                    self.special_routes = spec_map

        if df.empty:
            return {}

        columns = [clean_text(col) for col in df.columns]
        df.columns = columns
        normalized = {normalize_header(col): col for col in columns}

        def pick(*names: str) -> "str | None":
            for name in names:
                if normalize_header(name) in normalized:
                    return normalized[normalize_header(name)]
            return None

        company_col = pick("Company Name", "Company", "SEN.UNIT", "Customer")
        addr_col    = pick("Address", "Full Address", "Address Full")
        addr1_col   = pick("Address Line 1", "Address1", "Address 1", "AddressLine1")
        addr2_col   = pick("Address Line 2", "Address2", "Address 2", "AddressLine2")
        city_col    = pick("City", "Town")
        gstin_col   = pick("GSTIN", "GST Number", "GSTIN No")
        origin_col  = pick("Origin(s)", "Origins", "Origin", "Origin(s) / Places")
        sac_col     = pick("SAC Code", "SAC", "Default SAC")

        result: Dict[str, Dict[str, Any]] = {}
        for _, row in df.iterrows():
            company_name = clean_text(row.get(company_col)) if company_col else ""
            if not company_name:
                continue

            if addr_col:
                raw = row.get(addr_col)
                if raw is None or (isinstance(raw, float) and math.isnan(raw)):
                    address_full = ""
                else:
                    address_full = str(raw).strip()
                lines = [l.strip() for l in address_full.split("\n") if l.strip()]
                addr1 = lines[0] if len(lines) > 0 else ""
                addr2 = lines[1] if len(lines) > 1 else ""
                city  = lines[-1] if len(lines) > 2 else (lines[1] if len(lines) == 2 else "")
            else:
                addr1 = clean_text(row.get(addr1_col)) if addr1_col else ""
                addr2 = clean_text(row.get(addr2_col)) if addr2_col else ""
                city  = clean_text(row.get(city_col))  if city_col  else ""
                address_full = "\n".join(p for p in [addr1, addr2, city] if p)

            origins_raw = clean_text(row.get(origin_col)) if origin_col else ""
            origins = [p.strip() for p in re.split(r"[,;|]+", origins_raw) if p.strip()] \
                      if origins_raw else []

            gstin = clean_text(row.get(gstin_col)) if gstin_col else ""
            if not gstin:
                # The mapping file may have no dedicated GSTIN column and
                # instead carry the GSTIN inline inside the Address text
                # (confirmed against the uploaded mapping file). Extract it
                # with the standard 15-character GSTIN pattern and strip the
                # "GSTIN:" label (and the code itself) back out of the
                # address text so it isn't duplicated on the bill.
                m = _GSTIN_RE.search(address_full)
                if m:
                    gstin = m.group(1).upper()
                    address_full = _GSTIN_LABEL_RE.sub("", address_full).strip(" ,\n")
                    lines = [l.strip() for l in address_full.split("\n") if l.strip()]
                    addr1 = lines[0] if len(lines) > 0 else addr1
                    addr2 = lines[1] if len(lines) > 1 else addr2
                    city  = lines[-1] if len(lines) > 2 else (lines[1] if len(lines) == 2 else city)

            entry = {
                "company":        company_name,
                "address":        address_full,
                "address_line_1": addr1,
                "address_line_2": addr2,
                "city":           city,
                "gstin":          gstin,
                "origins":        origins,
                "sac_code":       clean_text(row.get(sac_col)) if sac_col else DEFAULT_BILL_SAC_CODE,
            }
            key = self.normalize_company_name(company_name)
            if key not in result:
                result[key] = entry
            else:
                ex = result[key]
                for f in ("address", "address_line_1", "address_line_2", "city", "gstin", "sac_code"):
                    if not ex.get(f) and entry.get(f):
                        ex[f] = entry[f]
                merged = list(ex.get("origins", [])) + list(entry.get("origins", []))
                seen: set = set()
                ex["origins"] = [o for o in (clean_text(x) for x in merged)
                                 if o and not (o.casefold() in seen or seen.add(o.casefold()))]
                ex["company"] = company_name
        self.company_mapping = result
        return result

    @staticmethod
    def is_valid_company_name(value: Any) -> bool:
        text = clean_text(value)
        if not text: return False
        norm = normalize_header(text)
        if norm in {"nan", "n", "na", "null", "total", "grandtotal", "summary", "-"}: return False
        if text.lower().startswith(("total", "summary")): return False
        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", text.replace(",", "").strip()): return False
        return True

    def get_company_names_for_period(
        self,
        excel_path: str,
        start_date: Any,
        end_date: Any,
        bill_marker: str,
    ) -> List[str]:
        """
        Dynamically derive the Company / SEN.UNIT dropdown from the SOURCE Excel
        (never from the mapping file), restricted to rows whose DATE falls in
        [start_date, end_date] and whose BILL marker matches. Per requirement #12.
        """
        if not excel_path or not os.path.exists(excel_path):
            return []
        start_dt = parse_date(start_date)
        end_dt   = parse_date(end_date)
        if start_dt is None or end_dt is None:
            return []

        marker = self.normalize_bill_marker(bill_marker)
        names_by_key: Dict[str, str] = {}
        try:
            wb = pd.ExcelFile(excel_path)
        except Exception:
            return []

        for sheet in wb.sheet_names:
            try:
                raw = pd.read_excel(excel_path, sheet_name=sheet, header=None, dtype=object)
            except Exception:
                continue
            if raw is None or raw.empty:
                continue
            hdr_idx = self.find_header_row(raw)
            if hdr_idx is None:
                continue
            headers, data = self._extract_headers_data(raw, hdr_idx)
            lookup = {
                "SEN.UNIT": self.resolve_header(headers, "SEN.UNIT", "SEN UNIT", "Company", "Company Name", "Customer"),
                "DATE":     self.resolve_header(headers, "DATE", "Date", "Transaction Date"),
                "BILL":     self.resolve_header(headers, "BILL", "Bill (C)", "Bill(C)", "Bill", "BILL C"),
            }
            sc, dc, bc = lookup["SEN.UNIT"], lookup["DATE"], lookup["BILL"]
            if not sc or not dc or not bc:
                continue
            for _, row in data.iterrows():
                dv = parse_date(row.get(dc, ""))
                if dv is None or not (start_dt <= dv <= end_dt):
                    continue
                if not self.matches_bill_marker(clean_text(row.get(bc, "")), marker):
                    continue
                cv = clean_text(row.get(sc, ""))
                if not self.is_valid_company_name(cv):
                    continue
                ck = self.normalize_company_name(cv)
                if ck and ck not in names_by_key:
                    names_by_key[ck] = cv

        return sorted(names_by_key.values(), key=str.casefold)

    def get_company_names_from_workbook(self, excel_path: str) -> List[str]:
        return self.get_company_names_for_period(excel_path, "01.01.2000", "31.12.2100", "+")

    def find_header_row(self, raw: pd.DataFrame) -> "int | None":
        if raw is None or raw.empty:
            return None
        for i in range(min(30, len(raw))):
            norms = [normalize_header(v) for v in raw.iloc[i].tolist()]
            if sum(1 for v in norms if v in _SOURCE_HEADER_TOKENS) >= 5:
                return i
        return None

    @staticmethod
    def _extract_headers_data(raw: pd.DataFrame, hdr_idx: int):
        """
        Build the header list starting from row `hdr_idx`.

        Some source workbooks (confirmed against the actual uploaded source
        Excel) split a column's real label across two physical rows: the
        cell in the main header row is blank or holds a stray summary value
        (e.g. a row count), while the true label ("Wt.MT.") sits in the row
        directly beneath it. If we only ever read `hdr_idx`, that column's
        header is lost and the field silently never gets populated. So: for
        any column where the primary header row doesn't already give us a
        recognized field name, check the very next row for one of the known
        header tokens and adopt it if found. If we adopt anything from that
        row, it is a header continuation, not a data row, so data starts one
        row later than usual.
        """
        primary = raw.iloc[hdr_idx].tolist()
        next_row = raw.iloc[hdr_idx + 1].tolist() if hdr_idx + 1 < len(raw) else []

        combined = list(primary)
        absorbed_any = False
        if next_row:
            for idx, val in enumerate(next_row):
                norm = normalize_header(val)
                if not norm or norm not in _SOURCE_HEADER_TOKENS:
                    continue
                existing_norm = normalize_header(primary[idx]) if idx < len(primary) else ""
                if existing_norm in _SOURCE_HEADER_TOKENS:
                    continue  # primary row already has a usable label here
                if idx < len(combined):
                    combined[idx] = val
                else:
                    combined.append(val)
                absorbed_any = True

        headers, counts = [], {}
        for idx, val in enumerate(combined):
            name = clean_text(val) or f"Unnamed_{idx+1}"
            if name in counts:
                counts[name] += 1; name = f"{name}_{counts[name]}"
            else:
                counts[name] = 1
            headers.append(name)

        data_start = hdr_idx + (2 if absorbed_any else 1)
        data = raw.iloc[data_start:].copy()
        data.columns = headers
        return headers, data

    def get_bill_rows(
        self,
        excel_path: str,
        company_name: str,
        start_date: Any,
        end_date: Any,
        bill_marker: str,
        return_debug: bool = False,
    ):
        """
        Return the SOURCE Excel transaction rows for the selected company, date
        range, and BILL marker. This is the ONLY source of transaction data for
        the generated bill -- mapping rows are never used as transactions.
        """
        empty = ([], {"company_matches": 0, "date_matches": 0, "bill_matches": 0}) if return_debug else []
        if not excel_path or not os.path.exists(excel_path):
            return empty

        start_dt = parse_date(start_date)
        end_dt   = parse_date(end_date)
        if start_dt is None or end_dt is None:
            return empty

        marker      = self.normalize_bill_marker(bill_marker)
        company_key = self.normalize_company_name(company_name)
        matches: List[Dict[str, Any]] = []
        debug = {"company_matches": 0, "date_matches": 0, "bill_matches": 0}

        try:
            wb = pd.ExcelFile(excel_path)
        except Exception:
            return ([], debug) if return_debug else []

        for sheet in wb.sheet_names:
            try:
                raw = pd.read_excel(excel_path, sheet_name=sheet, header=None, dtype=object)
            except Exception:
                continue
            if raw is None or raw.empty:
                continue
            hdr_idx = self.find_header_row(raw)
            if hdr_idx is None:
                continue
            headers, data = self._extract_headers_data(raw, hdr_idx)

            lk = {
                "SEN.UNIT": self.resolve_header(headers, "SEN.UNIT", "SEN UNIT", "Company", "Company Name", "Customer"),
                "BILL":     self.resolve_header(headers, "BILL", "Bill (C)", "Bill(C)", "Bill", "BILL C"),
                "DATE":     self.resolve_header(headers, "DATE", "Date", "Transaction Date"),
                "L.R.NO":   self.resolve_header(headers, "L.R.NO", "LR No", "L.R. No", "LRNO", "Lr No", "LR No.", "LR Number"),
                "TRUCK No.":self.resolve_header(headers, "TRUCK No.", "Truck No", "Truck No.", "Truck Number", "TRUCK NO"),
                "ORIGIN":   self.resolve_header(headers, "ORIGIN", "Origin", "From", "Source", "From City"),
                "TOWN":     self.resolve_header(headers, "TOWN", "Town", "Destination", "City", "To"),
                "STATE":    self.resolve_header(headers, "STATE", "State"),
                "WT. MT.":  self.resolve_header(headers, "WT. MT.", "WT MT", "WT. MT", "WT", "Weight"),
                "RATE":     self.resolve_header(headers, "RATE", "Rate"),
                "Remarks":  self.resolve_header(headers, "Remarks", "Remark"),
                # "PAY" is the actual per-row billable amount column in the
                # real source workbook (confirmed against the uploaded
                # source Excel); a literal "TOTAL"/"Amount" header, when
                # present, is tried too but PAY is listed first because in
                # the real file a later, unrelated "TOTAL" column exists
                # that is always 0 and must not be picked in its place.
                "N":        self.resolve_header(headers, "PAY", "Pay", "N", "TOTAL", "Total", "Amount", "Amount Total"),
            }

            for _, row in data.iterrows():
                cv = clean_text(row.get(lk.get("SEN.UNIT", ""), ""))
                dv = parse_date(row.get(lk.get("DATE", ""), ""))
                bv = self.normalize_bill_value(row.get(lk.get("BILL", ""), ""))

                if self.normalize_company_name(cv) == company_key: debug["company_matches"] += 1
                if dv is not None and start_dt <= dv <= end_dt:    debug["date_matches"] += 1
                if self.matches_bill_marker(bv, marker):           debug["bill_matches"] += 1

                if self.normalize_company_name(cv) != company_key: continue
                if dv is None or not (start_dt <= dv <= end_dt):   continue
                if not self.matches_bill_marker(bv, marker):       continue

                # Amount MUST come from Excel column N (col index 13)
                amount_raw = row.iloc[13] if len(row) > 13 else row.get(lk.get("N", ""), "")
                amount = parse_amount(amount_raw)
                if amount is None:
                    continue
                rounded_amount = int(round(amount))

                # Rate MUST come from Excel column M (col index 12)
                rate_raw = row.iloc[12] if len(row) > 12 else row.get(lk.get("RATE", ""), "")
                rate_str = format_rate_value(rate_raw)

                # WT from column H (index 7) first, fallback to I (index 8) with FIX logic
                h_raw = row.iloc[7] if len(row) > 7 else row.get(lk.get("WT. MT.", ""), "")
                i_raw = row.iloc[8] if len(row) > 8 else None
                wt_str = format_wt_value(h_raw, i_raw)

                matches.append({
                    "company":  cv,
                    "origin":   clean_text(row.get(lk.get("ORIGIN", ""), "")) if lk.get("ORIGIN") else "",
                    "date":     dv,
                    "date_text": dv.strftime("%d.%m.%y"),
                    "lr_no":    format_lr(row.get(lk.get("L.R.NO", ""), "")),
                    "truck_no": clean_text(row.get(lk.get("TRUCK No.", ""), "")),
                    "town":     clean_text(row.get(lk.get("TOWN", ""), "")),
                    "state":    clean_text(row.get(lk.get("STATE", ""), "")),
                    "wt_value": wt_str,
                    "rate":     rate_str,
                    "amount":   rounded_amount,
                    "remarks":  clean_text(row.get(lk.get("Remarks", ""), "")),
                    "sheet":    sheet,
                    "sac_code": DEFAULT_BILL_SAC_CODE,
                })

        matches.sort(key=lambda x: x["date"])
        return (matches, debug) if return_debug else matches

    def resolve_transaction_route(
        self,
        row: Dict[str, Any],
        mapping: Dict[str, Any],
        special_routes: "Dict[str, Tuple[str, str]] | None" = None,
    ) -> Tuple[str, str]:
        """
        Resolve the final Origin and Destination for a transaction row:
        1. First check if TOWN matches a Special Route in Sheet2 (case-insensitive & whitespace-tolerant).
        2. If matched: use Special Route Origin and Destination (ignores company normal origin).
        3. If not matched: use normal company mapping route logic.
        """
        if special_routes is None:
            special_routes = getattr(self, "special_routes", {}) or {}

        town_val = clean_text(row.get("town", ""))
        norm_town = self.normalize_route_key(town_val)

        # 1. First check whether the TOWN matches a Special Route in Sheet2
        if norm_town and norm_town in special_routes:
            spec_orig, spec_dest = special_routes[norm_town]
            final_orig = clean_text(spec_orig) or "MAIN ORIGIN"
            final_dest = clean_text(spec_dest) or town_val or "UNKNOWN"
            return final_orig, final_dest

        # 2. Normal company mapping route logic
        origins_from_map = [clean_text(o) for o in (mapping.get("origins") or []) if clean_text(o)]
        default_origin = origins_from_map[0] if origins_from_map else "MAIN ORIGIN"
        origin = clean_text(row.get("origin", ""))
        if not origin:
            town_candidate = town_val
            origin = next((o for o in origins_from_map if town_candidate.casefold() == o.casefold()), default_origin)
        final_orig = origin or default_origin
        final_dest = town_val or "UNKNOWN"
        return final_orig, final_dest

    def build_route_groups(
        self,
        rows: List[Dict[str, Any]],
        mapping: Dict[str, Any],
        special_routes: "Dict[str, Tuple[str, str]] | None" = None,
    ) -> Dict[Tuple[str, str], List[Dict[str, Any]]]:
        if special_routes is None:
            special_routes = getattr(self, "special_routes", {}) or {}

        # Canonical key: (normalized origin, normalized destination)
        # Value: ((display_origin, display_dest), list_of_rows)
        groups_map: Dict[Tuple[str, str], Tuple[Tuple[str, str], List[Dict[str, Any]]]] = {}

        for row in rows:
            final_orig, final_dest = self.resolve_transaction_route(row, mapping, special_routes)
            norm_key = (
                re.sub(r"\s+", " ", clean_text(final_orig)).strip().casefold(),
                re.sub(r"\s+", " ", clean_text(final_dest)).strip().casefold(),
            )
            if norm_key not in groups_map:
                groups_map[norm_key] = ((final_orig.strip(), final_dest.strip()), [row])
            else:
                groups_map[norm_key][1].append(row)

        result: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
        for (disp_orig, disp_dest), group_rows in groups_map.values():
            result[(disp_orig, disp_dest)] = group_rows
        return result

    def ensure_template_exists(self, template_path: "str | None") -> str:
        if not template_path:
            raise FileNotFoundError(
                "No Word template selected. Use the 'Choose Word Template...' button "
                "to select your .docx file before generating a bill."
            )
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Word template not found: {template_path}")
        return template_path

    def generate_bill_document(
        self,
        excel_path: str,
        company_name: str,
        start_date: Any,
        end_date: Any,
        bill_marker: str,
        bill_no: str,
        bill_date: Any,
        sac_code: str,
        mapping_path: str,
        output_path: str,
        template_path: "str | None" = None,
    ) -> Dict[str, Any]:
        template_path = self.ensure_template_exists(template_path)

        mapping        = self.load_mapping_file(mapping_path)
        company_record = mapping.get(self.normalize_company_name(company_name), {})
        if not company_record:
            raise ValueError(f"No company mapping found for '{company_name}'.")

        rows = self.get_bill_rows(excel_path, company_name, start_date, end_date, bill_marker)
        if not rows:
            raise ValueError("No matching bill rows found for the selected company and date range.")

        route_groups = self.build_route_groups(rows, company_record, self.special_routes)

        grand_total  = int(round(sum(r["amount"] for r in rows)))
        amount_words = amount_to_words_inr(grand_total)

        self._write_bill_docx(
            template_path=template_path,
            output_path=output_path,
            company_record=company_record,
            company_name=company_name,
            bill_no=bill_no,
            bill_date=bill_date,
            route_groups=route_groups,
            sac_code=(sac_code or DEFAULT_BILL_SAC_CODE),
            grand_total=grand_total,
            amount_words=amount_words,
        )
        return {
            "rows":         len(rows),
            "routes":       len(route_groups),
            "grand_total":  grand_total,
            "amount_words": amount_words,
            "output_path":  output_path,
        }

    @staticmethod
    def _amount_words_body(amount_words: str) -> str:
        text = clean_text(amount_words)
        if text.lower().startswith("rupee "): text = text[6:]
        if text.lower().endswith(" only."):   text = text[:-6].strip()
        return text

    def _write_bill_docx(
        self,
        template_path: str,
        output_path: str,
        company_record: Dict[str, Any],
        company_name: str,
        bill_no: str,
        bill_date: Any,
        route_groups: Dict[Tuple[str, str], List[Dict[str, Any]]],
        sac_code: str,
        grand_total: float,
        amount_words: str,
    ) -> None:
        """
        1. Copy the selected template byte-for-byte (preserves all layout, logo,
           styles, MSME table, bank details, footer, signatures, etc.).
        2. Replace text placeholders ([BILL NO], [DATE], [COMPANY NAME], etc.) in
           the surrounding Word content using lxml (preserving XML namespaces).
        3. For EACH unique route (Origin -> Destination), create one separate copy
           of the template's embedded Excel spreadsheet object, containing ONLY
           that route's rows and its independent subtotal in the TOTAL row.
        4. Dynamically set the route heading for each route: '1) [ORIGIN] TO [DESTINATION]'.
        5. If a route has more rows than the template originally provided for,
           increase ONLY that table object's HEIGHT proportionally, preserving its
           EXACT original width (478.5pt / A1:J{last_row}), columns, borders, fonts, and styling.
        6. Regenerate preview images (EMF via Excel COM on Windows / PNG fallback).
        """
        if etree is None:
            raise RuntimeError("lxml is required for generating bills. Install it with: pip install lxml")

        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        # ── Step 1: embedded OLE spreadsheet discovery from template ──
        ole_info = find_ole_spreadsheet(template_path)
        if ole_info is None:
            raise ValueError(
                "The selected template does not contain the required embedded "
                "Excel Spreadsheet. Please choose the correct Word template."
            )

        with zipfile.ZipFile(template_path, "r") as z:
            template_xlsx_bytes = z.read(ole_info.xlsx_part)
            template_img_bytes = z.read(ole_info.image_part) if (ole_info.image_part and ole_info.image_part in z.namelist()) else b""
            doc_xml_bytes = z.read("word/document.xml")
            rels_xml_bytes = z.read("word/_rels/document.xml.rels")
            ct_xml_bytes = z.read("[Content_Types].xml")

        # ── Step 2: Parse document.xml with lxml (preserves all prefixes & ignorable markup) ──
        parser = etree.XMLParser(recover=True, remove_blank_text=False)
        doc_tree = etree.fromstring(doc_xml_bytes, parser)

        # Text placeholder replacements
        addr1 = company_record.get("address_line_1", "").rstrip(", ")
        addr2 = company_record.get("address_line_2", "").rstrip(", ")
        city  = company_record.get("city", "").rstrip(", ")
        address_full = company_record.get("address", "") or "\n".join(p for p in [addr1, addr2, city] if p)

        bill_date_str = bill_date.strftime("%d.%m.%y") if hasattr(bill_date, "strftime") else str(bill_date).strip()

        gt = grand_total
        total_str = format_indian_currency(gt)

        replacements = {
            "[BILL NO]":         bill_no,
            "[DATE]":            bill_date_str,
            "[COMPANY NAME]":    company_name,
            "[ADDRESS]":         address_full,
            "[ADDRESS LINE 1]":  addr1,
            "[ADDRESS LINE 2]":  addr2,
            "[CITY]":            city,
            "[GSTIN]":           company_record.get("gstin", ""),
            "[TOTAL]":           total_str,
            "[AMOUNT IN WORDS]": self._amount_words_body(amount_words),
        }
        _replace_placeholders_in_lxml(doc_tree, replacements)

        # Also replace placeholders in all headers and footers so Bill No, Date, etc. repeat on every page
        header_replacements = {}
        with zipfile.ZipFile(template_path, "r") as z:
            for n in z.namelist():
                if (n.startswith("word/header") or n.startswith("word/footer")) and n.endswith(".xml"):
                    hdr_xml = z.read(n)
                    hdr_tree = etree.fromstring(hdr_xml, parser)
                    _replace_placeholders_in_lxml(hdr_tree, replacements)
                    header_replacements[n] = etree.tostring(hdr_tree, encoding="utf-8", xml_declaration=True, standalone=True)

        # ── Step 3: Locate route block in body ──
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        v_ns = "urn:schemas-microsoft-com:vml"
        o_ns = "urn:schemas-microsoft-com:office:office"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        w14_ns = "http://schemas.microsoft.com/office/word/2010/wordml"
        nsmap = {"w": w_ns, "v": v_ns, "o": o_ns, "r": r_ns, "w14": w14_ns}

        bodies = doc_tree.xpath("./w:body", namespaces=nsmap)
        if not bodies:
            raise ValueError("Invalid document.xml: <w:body> not found.")
        body = bodies[0]
        body_children = list(body)

        ole_elem = None
        ole_idx = None
        for idx, child in enumerate(body_children):
            if child.xpath(".//o:OLEObject", namespaces=nsmap) or child.xpath(".//w:object", namespaces=nsmap):
                ole_elem = child
                ole_idx = idx
                break

        if ole_elem is None:
            raise ValueError("Could not find embedded OLE object in word/document.xml")

        heading_elem = None
        heading_idx = None
        for idx in range(ole_idx - 1, -1, -1):
            child = body_children[idx]
            if child.tag.endswith("p"):
                text = "".join(child.xpath(".//w:t/text()", namespaces=nsmap))
                if "[ORIGIN]" in text or "[DESTINATION]" in text or "TO" in text or re.match(r"^\s*\d+[\).]", text):
                    heading_elem = child
                    heading_idx = idx
                    break

        if heading_elem is None:
            heading_idx = max(0, ole_idx - 1)
            heading_elem = body_children[heading_idx]

        spacing_idx = None
        spacing_elem = None
        if ole_idx + 1 < len(body_children):
            next_child = body_children[ole_idx + 1]
            if next_child.tag.endswith("p") and "".join(next_child.xpath(".//w:t/text()", namespaces=nsmap)).strip() == "":
                spacing_idx = ole_idx + 1
                spacing_elem = next_child

        replace_start = heading_idx
        replace_end = (spacing_idx + 1) if spacing_idx is not None else (ole_idx + 1)

        # Remove any static outside Grand Total paragraphs or Amount in Words tables from Word body
        while replace_end < len(body_children):
            cand = body_children[replace_end]
            cand_text = "".join(cand.xpath(".//w:t/text()", namespaces=nsmap)).strip()
            if (
                "GRAND TOTAL" in cand_text
                or "[TOTAL]" in cand_text
                or "[AMOUNT IN WORDS]" in cand_text
                or "Rupees" in cand_text
            ):
                replace_end += 1
            else:
                break

        # ── Step 4: Build each route's workbook, preview image, and XML elements ──
        new_route_elements = []
        docx_replacements: Dict[str, bytes] = {}
        rels_to_add: List[Tuple[str, str, str]] = []

        rels_tree = etree.fromstring(rels_xml_bytes, parser)
        existing_rids = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', rels_xml_bytes.decode("utf-8"))]
        next_rid_num = (max(existing_rids) if existing_rids else 20) + 1

        shape_width_pt = ole_info.shape_width_pt or 478.5
        shape_width_dxa = int(round(shape_width_pt * 20))

        with tempfile.TemporaryDirectory() as tmp:
            for route_num, ((origin, dest), route_rows) in enumerate(route_groups.items(), start=1):
                route_subtotal = int(round(sum(float(r.get("amount", 0) or 0) for r in route_rows)))

                # Fill route workbook with ONLY this route's data
                route_xlsx_bytes, route_ole_size_ref, route_total_rows = _fill_embedded_workbook(
                    xlsx_bytes=template_xlsx_bytes,
                    ole_info=ole_info,
                    rows=route_rows,
                    subtotal=route_subtotal,
                )

                route_xlsx_tmp = os.path.join(tmp, f"route_{route_num}.xlsx")
                with open(route_xlsx_tmp, "wb") as f:
                    f.write(route_xlsx_bytes)

                route_img_bytes = None
                route_img_ext = "emf"
                measured_w = None
                measured_h = None

                # Primary: native Windows Excel COM (vector EMF)
                if win32 is not None:
                    tmp_emf = os.path.join(tmp, f"preview_{route_num}.emf")
                    ok, m_w, m_h = _render_preview_via_win32(route_xlsx_tmp, ole_info.sheet_name, route_ole_size_ref, tmp_emf)
                    if ok:
                        with open(tmp_emf, "rb") as f:
                            route_img_bytes = f.read()
                        route_img_ext = "emf"
                        measured_w = m_w
                        measured_h = m_h

                # Fallback: LibreOffice PNG
                if route_img_bytes is None:
                    tmp_png = os.path.join(tmp, f"preview_{route_num}.png")
                    if _render_preview_via_libreoffice(route_xlsx_tmp, route_ole_size_ref, tmp_png):
                        with open(tmp_png, "rb") as f:
                            route_img_bytes = f.read()
                        route_img_ext = "png"

                # Fallback: template image
                if route_img_bytes is None:
                    route_img_bytes = template_img_bytes
                    route_img_ext = "emf" if (ole_info.image_part and ole_info.image_part.endswith(".emf")) else "png"

                # Height & Width calculation: match exact live Excel COM measurements
                # so the table is created at its full bigger size and stays the same size when clicked
                route_width_pt = measured_w if (measured_w is not None) else 487.2
                route_height_pt = measured_h if (measured_h is not None) else (17.25 + (route_total_rows - 2) * 14.5 + 16.5)
                shape_width_dxa = int(round(route_width_pt * 20))
                route_dya = int(round(route_height_pt * 20))

                img_rid = f"rId{next_rid_num}"
                next_rid_num += 1
                pkg_rid = f"rId{next_rid_num}"
                next_rid_num += 1

                shape_id = f"_x0000_i{1044 + route_num - 1}"
                object_id = f"_{1851433745 + route_num - 1}"

                pkg_filename = f"Microsoft_Excel_Worksheet{route_num}.xlsx"
                pkg_part_name = f"word/embeddings/{pkg_filename}"
                pkg_target = f"embeddings/{pkg_filename}"

                img_filename = f"image_ole_{route_num}.{route_img_ext}"
                img_part_name = f"word/media/{img_filename}"
                img_target = f"media/{img_filename}"

                docx_replacements[pkg_part_name] = route_xlsx_bytes
                docx_replacements[img_part_name] = route_img_bytes

                rels_to_add.append((img_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", img_target))
                rels_to_add.append((pkg_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", pkg_target))

                # Cloned heading
                cloned_heading = copy_module.deepcopy(heading_elem)
                if f"{{{w14_ns}}}paraId" in cloned_heading.attrib:
                    cloned_heading.attrib[f"{{{w14_ns}}}paraId"] = f"{0x11000000 + route_num * 2:08X}"

                # Keep heading together with the table (prevent page break between heading and table)
                p_prs = cloned_heading.xpath("./w:pPr", namespaces=nsmap)
                if not p_prs:
                    p_pr = etree.Element(f"{{{w_ns}}}pPr")
                    cloned_heading.insert(0, p_pr)
                else:
                    p_pr = p_prs[0]

                if not p_pr.xpath("./w:keepNext", namespaces=nsmap):
                    p_pr.append(etree.Element(f"{{{w_ns}}}keepNext"))
                if not p_pr.xpath("./w:keepLines", namespaces=nsmap):
                    p_pr.append(etree.Element(f"{{{w_ns}}}keepLines"))

                # Tight paragraph spacing so heading and table are visually and structurally bonded
                sp_nodes = p_pr.xpath("./w:spacing", namespaces=nsmap)
                if not sp_nodes:
                    sp_elem = etree.Element(f"{{{w_ns}}}spacing")
                    sp_elem.set(f"{{{w_ns}}}before", "100" if route_num > 1 else "0")
                    sp_elem.set(f"{{{w_ns}}}after", "20")
                    sp_elem.set(f"{{{w_ns}}}line", "240")
                    sp_elem.set(f"{{{w_ns}}}lineRule", "auto")
                    p_pr.append(sp_elem)
                else:
                    sp_nodes[0].set(f"{{{w_ns}}}after", "20")
                    if route_num > 1:
                        sp_nodes[0].set(f"{{{w_ns}}}before", "100")

                runs = cloned_heading.xpath("./w:r", namespaces=nsmap)
                heading_text = f"{route_num}) {origin.upper()} TO {dest.upper()}"
                if runs:
                    first_t = runs[0].xpath(".//w:t", namespaces=nsmap)
                    if first_t:
                        first_t[0].text = heading_text
                        first_t[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    _set_run_cambria(runs[0])
                    for r in runs[1:]:
                        cloned_heading.remove(r)
                new_route_elements.append(cloned_heading)

                # Cloned OLE paragraph
                cloned_ole = copy_module.deepcopy(ole_elem)
                if f"{{{w14_ns}}}paraId" in cloned_ole.attrib:
                    cloned_ole.attrib[f"{{{w14_ns}}}paraId"] = f"{0x11000001 + route_num * 2:08X}"

                ole_p_prs = cloned_ole.xpath("./w:pPr", namespaces=nsmap)
                if not ole_p_prs:
                    ole_p_pr = etree.Element(f"{{{w_ns}}}pPr")
                    cloned_ole.insert(0, ole_p_pr)
                else:
                    ole_p_pr = ole_p_prs[0]

                if not ole_p_pr.xpath("./w:keepLines", namespaces=nsmap):
                    ole_p_pr.append(etree.Element(f"{{{w_ns}}}keepLines"))

                ole_sp_nodes = ole_p_pr.xpath("./w:spacing", namespaces=nsmap)
                if not ole_sp_nodes:
                    ole_sp = etree.Element(f"{{{w_ns}}}spacing")
                    ole_sp.set(f"{{{w_ns}}}before", "0")
                    ole_sp.set(f"{{{w_ns}}}after", "40")
                    ole_sp.set(f"{{{w_ns}}}line", "240")
                    ole_sp.set(f"{{{w_ns}}}lineRule", "auto")
                    ole_p_pr.append(ole_sp)
                else:
                    ole_sp_nodes[0].set(f"{{{w_ns}}}before", "0")

                for obj in cloned_ole.xpath(".//w:object", namespaces=nsmap):
                    obj.set(f"{{{w_ns}}}dxaOrig", str(shape_width_dxa))
                    obj.set(f"{{{w_ns}}}dyaOrig", str(route_dya))
                    # Remove duplicated v:shapetype on subsequent routes to avoid ID collision
                    if route_num > 1:
                        for st in obj.xpath(".//v:shapetype", namespaces=nsmap):
                            obj.remove(st)

                for shape in cloned_ole.xpath(".//v:shape", namespaces=nsmap):
                    shape.set("id", shape_id)
                    shape.set("style", f"width:{route_width_pt:.2f}pt;height:{route_height_pt:.2f}pt")

                for img_data in cloned_ole.xpath(".//v:imagedata", namespaces=nsmap):
                    img_data.set(f"{{{r_ns}}}id", img_rid)

                for ole_obj in cloned_ole.xpath(".//o:OLEObject", namespaces=nsmap):
                    ole_obj.set("ShapeID", shape_id)
                    ole_obj.set("ObjectID", object_id)
                    ole_obj.set(f"{{{r_ns}}}id", pkg_rid)

                new_route_elements.append(cloned_ole)

                if spacing_elem is not None:
                    cloned_spacing = copy_module.deepcopy(spacing_elem)
                    if f"{{{w14_ns}}}paraId" in cloned_spacing.attrib:
                        cloned_spacing.attrib[f"{{{w14_ns}}}paraId"] = f"{0x11000002 + route_num * 2:08X}"
                    new_route_elements.append(cloned_spacing)

        # ── Step 5: Replace template route block in body with new elements ──
        for elem in body_children[replace_start:replace_end]:
            body.remove(elem)

        for idx_offset, elem in enumerate(new_route_elements):
            body.insert(replace_start + idx_offset, elem)

        # Serialize document.xml and include headers/footers
        docx_replacements["word/document.xml"] = etree.tostring(doc_tree, encoding="utf-8", xml_declaration=True, standalone=True)
        docx_replacements.update(header_replacements)

        # ── Step 6: Update document.xml.rels ──
        old_rids = {ole_info.image_rel_id} if ole_info.image_rel_id else set()
        pkg_rel_m = re.search(r'<Relationship[^>]*Target="[^"]*embeddings/[^"]*"[^>]*Id="([^"]+)"', rels_xml_bytes.decode("utf-8"))
        if not pkg_rel_m:
            pkg_rel_m = re.search(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="[^"]*embeddings/[^"]*"', rels_xml_bytes.decode("utf-8"))
        if pkg_rel_m:
            old_rids.add(pkg_rel_m.group(1))

        for rel_child in list(rels_tree):
            if rel_child.get("Id") in old_rids:
                rels_tree.remove(rel_child)

        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        for rid, rtype, target in rels_to_add:
            etree.SubElement(rels_tree, f"{{{rel_ns}}}Relationship", {
                "Id": rid,
                "Type": rtype,
                "Target": target,
            })

        docx_replacements["word/_rels/document.xml.rels"] = etree.tostring(rels_tree, encoding="utf-8", xml_declaration=True, standalone=True)

        # ── Step 7: Update [Content_Types].xml ──
        ct_tree = etree.fromstring(ct_xml_bytes, parser)
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        existing_exts = {child.get("Extension") for child in ct_tree if child.tag.endswith("Default")}
        if "xlsx" not in existing_exts:
            etree.SubElement(ct_tree, f"{{{ct_ns}}}Default", {
                "Extension": "xlsx",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            })
        if "emf" not in existing_exts:
            etree.SubElement(ct_tree, f"{{{ct_ns}}}Default", {
                "Extension": "emf",
                "ContentType": "image/x-emf",
            })
        if "png" not in existing_exts:
            etree.SubElement(ct_tree, f"{{{ct_ns}}}Default", {
                "Extension": "png",
                "ContentType": "image/png",
            })
        docx_replacements["[Content_Types].xml"] = etree.tostring(ct_tree, encoding="utf-8", xml_declaration=True, standalone=True)

        # ── Step 8: Assemble clean output ZIP directly from template ──
        with zipfile.ZipFile(template_path, "r") as zin:
            parts = {n: zin.read(n) for n in zin.namelist() if n not in [ole_info.xlsx_part, ole_info.image_part]}

        parts.update(docx_replacements)

        tmp_out = output_path + ".tmp"
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in parts.items():
                zout.writestr(name, data)

        os.replace(tmp_out, output_path)



if __name__ == "__main__":
    service = BillGeneratorService()
    print("Bill generator ready (OLE embedded-spreadsheet mode).")
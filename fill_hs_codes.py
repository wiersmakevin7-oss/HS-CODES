"""
Vult automatisch HS-codes in een Excel-invoice op basis van artikelnummer.

Gebruik CLI:
    python fill_hs_codes.py invoice.xlsx output.xlsx

De mapping staat standaard in hs_mapping.csv naast dit script.
"""
from __future__ import annotations

import csv
import re
import sys
from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader

ARTICLE_HEADERS = {
    "article no",
    "article no.",
    "art no",
    "art no.",
    "artikelnummer",
    "article number",
    "buyer art no",
    "buyer art no.",
    "buyer article no",
    "buyer article no.",
    "item no",
    "item no.",
}
HS_HEADERS = {"hs code (artikellijst)", "hscode (artikellijst)", "hs-code (artikellijst)"}
SIZE_HEADERS = {"size", "maat"}


def norm_text(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def norm_key(value) -> str:
    text = norm_text(value).upper()
    text = text.replace("\u00a0", " ")
    text = text.replace(".", "")
    text = re.sub(r"\s+", " ", text)
    return text


def compact_key(value) -> str:
    return re.sub(r"[^A-Z0-9]", "", norm_key(value))


def norm_header(value) -> str:
    return re.sub(r"\s+", " ", norm_text(value).lower().replace(":", "").replace(".", ""))


def norm_hs(value) -> str:
    text = norm_text(value)
    if not text:
        return ""
    # Excel may store HS codes as floats/scientific notation in source files.
    try:
        if re.fullmatch(r"[0-9]+(\.0+)?", text):
            text = str(int(float(text)))
    except Exception:
        pass
    digits = re.sub(r"\D", "", text)
    return digits


def norm_size(value) -> str:
    text = norm_key(value)
    text = re.sub(r"\b(CM|MTR|METER)\b", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def load_mapping(mapping_csv: Path) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    with mapping_csv.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            article = norm_key(row.get("artikelnummer") or row.get("article") or row.get("article_no"))
            hs = norm_hs(row.get("hs_code") or row.get("HS code") or row.get("hs"))
            if article and hs:
                mapping[article] = hs
                compact_article = compact_key(article)
                if compact_article:
                    mapping[compact_article] = hs
    return mapping


def article_pattern(article: str) -> re.Pattern:
    escaped = re.escape(article).replace(r"\ ", r"\s+")
    return re.compile(rf"(?<![A-Z0-9]){escaped}(?![A-Z0-9])")


def norm_pdf_token(value: str) -> str:
    return norm_text(value).lower().replace(":", "").rstrip(".")


def get_pdf_page_text(page) -> str:
    try:
        return page.extract_text(extraction_mode="layout") or ""
    except TypeError:
        return page.extract_text() or ""


def is_packing_list_page(text: str) -> bool:
    normalized = re.sub(r"\s+", " ", text or "").strip().lower()
    compact = re.sub(r"[^a-z]", "", normalized)
    return "packing list" in normalized or "packlist" in compact or "装箱单" in (text or "")


def find_headers(ws) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int]]:
    """Return row, article_col, hs_col, size_col. 1-based indexes."""
    best = (None, None, None, None)
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 80)):
        article_col = None
        hs_col = None
        size_col = None
        for cell in row:
            h = norm_header(cell.value)
            if h in ARTICLE_HEADERS:
                article_col = cell.column
            if h in HS_HEADERS:
                hs_col = cell.column
            if h in SIZE_HEADERS:
                size_col = cell.column
        if article_col:
            best = (row[0].row, article_col, hs_col, size_col)
            break
    return best


def lookup_hs(mapping: Dict[str, str], article_value, size_value=None) -> Optional[str]:
    article = norm_key(article_value)
    if not article:
        return None

    candidates = [article]
    size = norm_size(size_value)
    if size:
        candidates.insert(0, norm_key(f"{article} {size}"))

    for candidate in candidates:
        hs = mapping.get(candidate) or mapping.get(compact_key(candidate))
        if hs:
            return hs
    return None


def is_invoice_detail_row(ws, row_idx: int, article_col: int, size_col: Optional[int]) -> bool:
    size_value = ws.cell(row_idx, size_col).value if size_col else None
    color_col = size_col - 1 if size_col and size_col > article_col + 1 else None
    color_value = ws.cell(row_idx, color_col).value if color_col else None
    has_variant = bool(norm_text(size_value) or norm_text(color_value))
    has_number = any(
        isinstance(ws.cell(row_idx, col).value, (int, float))
        for col in range(article_col + 1, ws.max_column + 1)
    )
    return has_variant and has_number


def copy_cell_style(src, dst) -> None:
    if src.has_style:
        dst._style = copy(src._style)
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format


def find_or_create_hs_col(ws, header_row: int, article_col: int, preferred_col: Optional[int]) -> int:
    if preferred_col:
        return preferred_col

    col = ws.max_column + 1
    ws.cell(header_row, col).value = "HS code (artikellijst)"

    for row_idx in range(1, ws.max_row + 1):
        copy_cell_style(ws.cell(row_idx, col - 1), ws.cell(row_idx, col))

    article_letter = ws.cell(header_row, article_col).column_letter
    hs_letter = ws.cell(header_row, col).column_letter
    if article_letter in ws.column_dimensions:
        ws.column_dimensions[hs_letter].width = ws.column_dimensions[article_letter].width
    return col


def fill_invoice(input_xlsx: Path, output_xlsx: Path, mapping_csv: Path) -> dict:
    mapping = load_mapping(mapping_csv)
    wb = load_workbook(input_xlsx)
    total_filled = 0
    unmatched = []

    invoice_sheets = [ws for ws in wb.worksheets if ws.title.strip().lower() == "invoice"]
    worksheets = invoice_sheets or wb.worksheets

    for ws in worksheets:
        header_row, article_col, hs_col, size_col = find_headers(ws)
        if not header_row or not article_col:
            continue
        created_hs_col = hs_col is None
        hs_col = find_or_create_hs_col(ws, header_row, article_col, hs_col)
        if created_hs_col and size_col and size_col >= hs_col:
            size_col += 1

        current_article_value = None

        for r in range(header_row + 1, ws.max_row + 1):
            article_value = ws.cell(r, article_col).value
            article = norm_key(article_value)
            if norm_header(article_value) in ARTICLE_HEADERS:
                current_article_value = None
                continue

            if article and article.startswith(("SUB TOTAL", "TOTAL", "EURO ", "N W", "NW", "G W", "GW", "NET WEIGHT", "GROSS WEIGHT")):
                current_article_value = None
                continue

            if article and not article.startswith("="):
                current_article_value = article_value
            elif not is_invoice_detail_row(ws, r, article_col, size_col):
                continue

            lookup_article = current_article_value
            if not norm_key(lookup_article):
                continue

            size_value = ws.cell(r, size_col).value if size_col else None
            hs = lookup_hs(mapping, lookup_article, size_value)
            if hs:
                ws.cell(r, hs_col).value = hs
                total_filled += 1
            else:
                unmatched.append({"sheet": ws.title, "row": r, "article": lookup_article})

    wb.save(output_xlsx)
    return {"filled": total_filled, "unmatched": unmatched[:100], "unmatched_count": len(unmatched)}


def extract_pdf_words(page) -> List[dict]:
    words = []

    def visitor_text(text, _cm, tm, _font_dict, font_size):
        if not norm_text(text):
            return
        x = float(tm[4])
        y = float(tm[5])
        char_width = max(float(font_size or 10) * 0.5, 3)
        for match in re.finditer(r"\S+", text):
            words.append({"text": match.group(0), "x": x + (match.start() * char_width), "source_x": x, "y": y})

    page.extract_text(visitor_text=visitor_text)
    return words


def group_pdf_lines(words: List[dict], y_tolerance: float = 4.0) -> List[dict]:
    lines = []
    for word in sorted(words, key=lambda item: (-item["y"], item["x"])):
        for line in lines:
            if abs(line["y"] - word["y"]) <= y_tolerance:
                line["words"].append(word)
                line["y"] = (line["y"] + word["y"]) / 2
                break
        else:
            lines.append({"y": word["y"], "words": [word]})

    for line in lines:
        line["words"].sort(key=lambda item: item["x"])
        line["text"] = " ".join(word["text"] for word in line["words"])
    return lines


def find_pdf_article_column(lines: List[dict]) -> Optional[dict]:
    for line in lines:
        words = line["words"]
        tokens = [norm_pdf_token(word["text"]) for word in words]
        for idx, token in enumerate(tokens[:-1]):
            if token != "article" or tokens[idx + 1] not in {"no", "number"}:
                continue

            article_x = words[idx]["x"]
            header_end_x = words[idx + 1]["x"] + max(len(words[idx + 1]["text"]) * 6, 15)
            header_starts = [word["x"] for word in words]
            previous_starts = [x for x in header_starts if x < article_x - 10]
            next_starts = [x for x in header_starts if x > header_end_x + 10]
            left = ((max(previous_starts) + article_x) / 2) if previous_starts else article_x - 20
            right = ((header_end_x + min(next_starts)) / 2) if next_starts else header_end_x + 120
            return {"header_y": line["y"], "left": left, "right": right}
    return None


def find_article_in_text(value: str, mapping: Dict[str, str], patterns: List[Tuple[str, re.Pattern]]) -> Optional[str]:
    normalized = norm_key(value)
    if not normalized:
        return None
    if normalized in mapping:
        return normalized
    compact = compact_key(normalized)
    if compact in mapping:
        return compact

    for article, pattern in patterns:
        if pattern.search(normalized):
            return article
    return None


def find_article_prefix(value: str, mapping: Dict[str, str]) -> Optional[str]:
    compact = compact_key(value)
    if not compact:
        return None

    for article in sorted(mapping, key=len, reverse=True):
        if len(article) < 4 or not compact.startswith(article):
            continue
        if article.isdigit() and len(compact) > len(article) and compact[len(article)].isdigit():
            continue
        if re.fullmatch(r"\d{6,12}", compact):
            continue
        if re.fullmatch(r"\d+(\.\d+)?", norm_text(value)):
            continue
        return article
    return None


def extract_article_before_po(text: str, mapping: Dict[str, str]) -> Optional[str]:
    before_po = re.split(r"\bPO\s*#", text, flags=re.I)[0]
    candidates = re.findall(r"\b[A-Z]?\d{3,5}(?:\s*[A-Z]{2,6})?(?:\s*\d{1,3}(?:-\d{1,3})?)?\b", before_po, flags=re.I)
    for candidate in reversed(candidates):
        article = find_article_in_text(candidate, mapping, [])
        if article:
            return article
    return None


def remove_po_number(text: str) -> str:
    return norm_text(re.sub(r"\bPO\s*#\s*[A-Z0-9-]+", "", text, flags=re.I))


def extract_pdf_articles_from_po_lines(reader: PdfReader, mapping: Dict[str, str]) -> List[dict]:
    rows = []
    seen = set()

    for page_number, page in enumerate(reader.pages, start=1):
        text = get_pdf_page_text(page)
        if is_packing_list_page(text):
            continue

        previous_line = ""
        for line_number, line in enumerate(text.splitlines(), start=1):
            clean_line = norm_text(line)
            if not clean_line:
                continue

            if re.search(r"\bPO\s*#", clean_line, flags=re.I):
                combined = norm_text(f"{previous_line} {clean_line}")
                article = extract_article_before_po(combined, mapping)
                if article:
                    key = (page_number, article)
                    if key not in seen:
                        seen.add(key)
                        rows.append(
                            {
                                "page": page_number,
                                "line": line_number,
                                "article": article,
                                "hs_code": mapping[article],
                                "text": remove_po_number(combined),
                            }
                        )

            previous_line = clean_line

    return rows


def find_product_code_bounds(header_line: str) -> Optional[Tuple[int, int]]:
    header_lower = header_line.lower()
    start = header_lower.find("our product code")
    if start < 0:
        return None

    next_headers = [
        header_lower.find("party's code"),
        header_lower.find("party code"),
        header_lower.find("order no"),
        header_lower.find("colour"),
        header_lower.find("color"),
    ]
    next_headers = [idx for idx in next_headers if idx > start]
    end = min(next_headers) if next_headers else start + 24
    return start, max(end, start + 12)


def extract_pdf_articles_left_of_order_no(
    reader: PdfReader,
    mapping: Dict[str, str],
    patterns: List[Tuple[str, re.Pattern]],
) -> List[dict]:
    rows = []

    for page_number, page in enumerate(reader.pages, start=1):
        text = get_pdf_page_text(page)
        if is_packing_list_page(text):
            continue

        order_col: Optional[int] = None
        for line_number, line in enumerate(text.splitlines(), start=1):
            header_match = re.search(r"\bORD\.?\s*NO\.?", line, flags=re.I)
            if header_match and re.search(r"\bSIZE\b", line[: header_match.start()], flags=re.I):
                order_col = header_match.start()
                continue

            if order_col is None or not norm_text(line):
                continue

            normalized_line = norm_key(line)
            if normalized_line.startswith(("TOTAL", "AMOUNT", "DECLARATION", "THE EXPORTER", "NET WEIGHT", "GROSS WEIGHT")):
                continue

            left_of_order = line[:order_col].strip()
            article = find_article_in_text(left_of_order, mapping, patterns)
            if not article:
                continue

            rows.append(
                {
                    "page": page_number,
                    "line": line_number,
                    "article": article,
                    "hs_code": mapping[article],
                    "text": norm_text(line),
                }
            )

    return rows


def find_party_code_bounds(header_line: str) -> Optional[Tuple[int, int]]:
    header_lower = header_line.lower()
    match = re.search(r"party['’]s\s+code", header_lower)
    if not match:
        return None

    start = max(0, match.start() - 2)
    next_headers = [
        header_lower.find("order no", match.end()),
        header_lower.find("order number", match.end()),
        header_lower.find("colour", match.end()),
        header_lower.find("color", match.end()),
    ]
    next_headers = [idx for idx in next_headers if idx > match.start()]
    end = min(next_headers) if next_headers else match.end() + 20
    return start, max(end, start + 12)


def find_buyer_product_code_bounds(header_line: str, next_line: str = "") -> Optional[Tuple[int, int]]:
    buyer_match = re.search(r"\bbuyer\b", header_line, flags=re.I)
    if not buyer_match:
        return None

    product_code_match = re.search(r"product\s+code", next_line, flags=re.I)
    if not product_code_match or abs(product_code_match.start() - buyer_match.start()) > 10:
        return None

    start = max(0, min(buyer_match.start(), product_code_match.start()) - 4)
    qty_match = re.search(r"\bqty\.?\b", header_line[buyer_match.end() :], flags=re.I)
    if qty_match:
        end = buyer_match.end() + qty_match.start()
    else:
        description_match = re.search(r"\bproduct\b", header_line[buyer_match.end() :], flags=re.I)
        end = buyer_match.end() + description_match.start() if description_match else start + 24
    return start, max(end, start + 14)


def extract_pdf_articles_from_buyer_product_code_column(reader: PdfReader, mapping: Dict[str, str]) -> List[dict]:
    rows = []
    last_bounds: Optional[Tuple[int, int]] = None

    for page_number, page in enumerate(reader.pages, start=1):
        text = get_pdf_page_text(page)
        if is_packing_list_page(text):
            continue

        lines = text.splitlines()
        bounds = None
        header_index = None
        for idx, line in enumerate(lines):
            next_line = lines[idx + 1] if idx + 1 < len(lines) else ""
            bounds = find_buyer_product_code_bounds(line, next_line)
            if bounds:
                header_index = idx + 1
                last_bounds = bounds
                break

        if bounds is None or header_index is None:
            if last_bounds is None:
                continue
            bounds = last_bounds
            header_index = -1

        start, end = bounds
        for line_number, line in enumerate(lines[header_index + 1 :], start=header_index + 2):
            if not norm_text(line):
                continue

            normalized_line = norm_key(line)
            if normalized_line.startswith(("TOTAL", "GRAND TOTAL", "AMOUNT", "DECLARATION", "SIGNATURE", "NET WEIGHT", "GROSS WEIGHT")):
                continue

            cell = line[start:end].strip() if len(line) > start else ""
            article = find_article_in_text(cell, mapping, []) or find_article_prefix(cell, mapping)
            if not article:
                continue

            rows.append(
                {
                    "page": page_number,
                    "line": line_number,
                    "article": article,
                    "hs_code": mapping[article],
                    "text": norm_text(line),
                }
            )

    return rows


def extract_pdf_articles_from_party_code_column(reader: PdfReader, mapping: Dict[str, str]) -> List[dict]:
    rows = []
    last_bounds: Optional[Tuple[int, int]] = None

    for page_number, page in enumerate(reader.pages, start=1):
        text = get_pdf_page_text(page)
        if is_packing_list_page(text):
            continue

        lines = text.splitlines()
        bounds = None
        header_index = None
        for idx, line in enumerate(lines):
            bounds = find_party_code_bounds(line)
            if bounds:
                header_index = idx
                last_bounds = bounds
                break

        if bounds is None or header_index is None:
            if last_bounds is None:
                continue
            bounds = last_bounds
            header_index = -1

        start, end = bounds
        for line_number, line in enumerate(lines[header_index + 1 :], start=header_index + 2):
            if not norm_text(line):
                continue

            normalized_line = norm_key(line)
            if normalized_line.startswith(("TOTAL", "GRAND TOTAL", "AMOUNT", "DECLARATION", "SIGNATURE", "GSTIN", "PAN")):
                continue

            cell = line[start:end].strip() if len(line) > start else ""
            article = find_article_in_text(cell, mapping, []) or find_article_prefix(cell, mapping)
            if not article:
                continue

            rows.append(
                {
                    "page": page_number,
                    "line": line_number,
                    "article": article,
                    "hs_code": mapping[article],
                    "text": norm_text(line),
                }
            )

    return rows


def extract_pdf_articles_from_product_code_column(reader: PdfReader, mapping: Dict[str, str]) -> List[dict]:
    rows = []
    seen = set()
    last_bounds: Optional[Tuple[int, int]] = None

    for page_number, page in enumerate(reader.pages, start=1):
        text = get_pdf_page_text(page)
        if is_packing_list_page(text):
            continue

        lines = text.splitlines()
        bounds = None
        header_index = None
        for idx, line in enumerate(lines):
            bounds = find_product_code_bounds(line)
            if bounds:
                header_index = idx
                last_bounds = bounds
                break

        if bounds is None or header_index is None:
            if last_bounds is None:
                continue
            bounds = last_bounds
            header_index = -1

        start, end = bounds
        for line_number, line in enumerate(lines[header_index + 1 :], start=header_index + 2):
            if not norm_text(line):
                continue

            normalized_line = norm_key(line)
            if normalized_line.startswith(("TOTAL", "GRAND TOTAL", "AMOUNT", "DECLARATION", "SIGNATURE", "GSTIN", "PAN")):
                continue

            cell = line[start:end].strip() if len(line) > start else ""
            article = find_article_prefix(cell, mapping)
            if not article:
                continue

            key = (page_number, line_number, article)
            if key in seen:
                continue
            seen.add(key)
            rows.append(
                {
                    "page": page_number,
                    "line": line_number,
                    "article": article,
                    "hs_code": mapping[article],
                    "text": norm_text(line),
                }
            )

    return rows


def extract_pdf_articles_from_positions(
    reader: PdfReader,
    mapping: Dict[str, str],
    patterns: List[Tuple[str, re.Pattern]],
) -> List[dict]:
    rows = []

    for page_number, page in enumerate(reader.pages, start=1):
        if is_packing_list_page(get_pdf_page_text(page)):
            continue
        lines = group_pdf_lines(extract_pdf_words(page))
        article_column = find_pdf_article_column(lines)
        if not article_column:
            continue

        for line_number, line in enumerate(lines, start=1):
            if line["y"] >= article_column["header_y"] - 2:
                continue

            column_words = [
                word["text"]
                for word in line["words"]
                if article_column["left"] <= word["source_x"] <= article_column["right"]
            ]
            column_text = norm_text(" ".join(column_words))
            if norm_key(column_text).startswith(("SUB TOTAL", "TOTAL", "EURO ", "N W", "NW", "G W", "GW")):
                continue

            matched_article = find_article_in_text(column_text, mapping, patterns)
            if matched_article:
                rows.append(
                    {
                        "page": page_number,
                        "line": line_number,
                        "article": matched_article,
                        "hs_code": mapping[matched_article],
                        "text": line["text"],
                    }
                )

    return rows


def find_layout_article_bounds(header_line: str, next_line: str = "") -> Optional[Tuple[int, int]]:
    header_lower = header_line.lower()
    match = re.search(r"\b(?:article|art)\.?\s*(?:no\.?|number)\b", header_lower)
    if not match:
        article_match = re.search(r"\barticle\b", header_lower)
        if not article_match:
            return None

        next_lower = next_line.lower()
        number_matches = list(re.finditer(r"\bnumber\b", next_lower))
        aligned_number = None
        for number_match in number_matches:
            if abs(number_match.start() - article_match.start()) <= 8:
                aligned_number = number_match
                break
        if not aligned_number:
            return None

        match = article_match

    short_art_header = bool(re.fullmatch(r"art\.?\s*no\.?", match.group(0), flags=re.I))
    start = max(0, match.start() - (10 if short_art_header else 3))
    next_matches = [
        m.start()
        for m in re.finditer(
            r"\b(item|qty|quantity|hsn|hs\s*code|price|amount|total|cartons?|ctn|pcs|description|desc)\b",
            header_lower[match.end() :],
        )
    ]
    end = match.end() + min(next_matches) if next_matches else len(header_line)
    return start, max(end, start + 18)


def extract_pdf_articles_from_layout(
    reader: PdfReader,
    mapping: Dict[str, str],
    patterns: List[Tuple[str, re.Pattern]],
) -> List[dict]:
    rows = []
    last_bounds: Optional[Tuple[int, int]] = None

    for page_number, page in enumerate(reader.pages, start=1):
        text = get_pdf_page_text(page)
        if is_packing_list_page(text):
            continue

        lines = text.splitlines()
        bounds = None
        header_index = None
        for idx, line in enumerate(lines):
            next_line = lines[idx + 1] if idx + 1 < len(lines) else ""
            bounds = find_layout_article_bounds(line, next_line)
            if bounds:
                header_index = idx
                last_bounds = bounds
                break

        if bounds is None or header_index is None:
            if last_bounds is None:
                continue
            bounds = last_bounds
            header_index = -1

        start, end = bounds
        for line_number, line in enumerate(lines[header_index + 1 :], start=header_index + 2):
            if not norm_text(line):
                continue

            normalized_line = norm_key(line)
            if normalized_line.startswith(("SUB TOTAL", "TOTAL", "EURO ", "N W", "NW", "G W", "GW")):
                continue

            article_cell = line[start:end].strip() if len(line) > start else ""
            matched_article = find_article_in_text(article_cell, mapping, patterns)
            if matched_article:
                rows.append(
                    {
                        "page": page_number,
                        "line": line_number,
                        "article": matched_article,
                        "hs_code": mapping[matched_article],
                        "text": norm_text(line),
                    }
                )

    return rows


def extract_pdf_articles_from_ocr(input_pdf: Path, mapping: Dict[str, str]) -> List[dict]:
    try:
        import pypdfium2 as pdfium
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        return []

    rows = []
    seen = set()
    text_reader = PdfReader(BytesIO(input_pdf.read_bytes()))
    document = pdfium.PdfDocument(str(input_pdf))
    ocr = RapidOCR()

    try:
        for page_number in range(len(document)):
            if page_number < len(text_reader.pages) and is_packing_list_page(get_pdf_page_text(text_reader.pages[page_number])):
                continue

            page = document[page_number]
            try:
                image = page.render(scale=2).to_pil()
            finally:
                if hasattr(page, "close"):
                    page.close()

            width, height = image.size
            table_image = image.crop(
                (
                    max(0, int(width * 0.08)),
                    int(height * 0.40),
                    min(width, int(width * 0.96)),
                    min(height, int(height * 0.78)),
                )
            )
            result, _elapsed = ocr(table_image)

            for box, text, confidence in result or []:
                if confidence < 0.75:
                    continue

                xs = [point[0] for point in box]
                ys = [point[1] for point in box]
                x_min = min(xs)
                y_min = min(ys)
                raw_text = norm_text(text)
                normalized = norm_key(raw_text)

                if not raw_text:
                    continue
                if x_min < 20 or x_min > 650:
                    continue
                if normalized.startswith(("ORDER", "SUBTOTAL", "TOTAL", "GROSS", "AMOUNT", "MARKS", "DESCRIPTION")):
                    continue

                article = find_article_prefix(raw_text, mapping)
                if not article:
                    continue

                key = (page_number + 1, round(y_min / 8), article)
                if key in seen:
                    continue
                seen.add(key)
                rows.append(
                    {
                        "page": page_number + 1,
                        "line": int(y_min),
                        "article": article,
                        "hs_code": mapping[article],
                        "text": raw_text,
                    }
                )
    finally:
        document.close()

    return rows


def extract_pdf_articles(input_pdf: Path, mapping_csv: Path) -> List[dict]:
    mapping = load_mapping(mapping_csv)
    patterns = [(article, article_pattern(article)) for article in sorted(mapping, key=len, reverse=True)]
    reader = PdfReader(BytesIO(input_pdf.read_bytes()))
    buyer_product_rows = extract_pdf_articles_from_buyer_product_code_column(reader, mapping)
    party_rows = extract_pdf_articles_from_party_code_column(reader, mapping)
    product_rows = extract_pdf_articles_from_product_code_column(reader, mapping)
    preferred_rows = max((buyer_product_rows, party_rows, product_rows), key=len)
    if preferred_rows:
        return preferred_rows
    rows = extract_pdf_articles_left_of_order_no(reader, mapping, patterns)
    if rows:
        return rows
    rows = extract_pdf_articles_from_po_lines(reader, mapping)
    if rows:
        return rows
    rows = extract_pdf_articles_from_positions(reader, mapping, patterns)
    if rows:
        return rows
    rows = extract_pdf_articles_from_layout(reader, mapping, patterns)
    if rows:
        return rows
    return extract_pdf_articles_from_ocr(input_pdf, mapping)


def fill_pdf_invoice(input_pdf: Path, output_xlsx: Path, mapping_csv: Path) -> dict:
    rows = extract_pdf_articles(input_pdf, mapping_csv)
    wb = Workbook()
    ws = wb.active
    ws.title = "HS codes"
    ws.append(["Page", "Article No.", "HS code", "PDF text"])

    for row in rows:
        ws.append([row["page"], row["article"], row["hs_code"], row["text"]])

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 100
    wb.save(output_xlsx)
    return {"filled": len(rows), "unmatched": [], "unmatched_count": 0}


if __name__ == "__main__":
    if len(sys.argv) not in (3, 4):
        print("Gebruik: python fill_hs_codes.py invoice.xlsx output.xlsx [hs_mapping.csv]")
        raise SystemExit(2)
    input_file = Path(sys.argv[1])
    output_file = Path(sys.argv[2])
    mapping_file = Path(sys.argv[3]) if len(sys.argv) == 4 else Path(__file__).with_name("hs_mapping.csv")
    result = fill_invoice(input_file, output_file, mapping_file)
    print(f"Klaar: {result['filled']} HS-codes ingevuld. Output: {output_file}")
    if result["unmatched_count"]:
        print(f"Niet gevonden: {result['unmatched_count']} artikelen. Eerste voorbeelden:")
        for item in result["unmatched"][:20]:
            print(f"- {item['sheet']} rij {item['row']}: {item['article']}")

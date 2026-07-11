import sys
import threading
import re
import datetime
import time
import platform
import os
import uuid
from pathlib import Path
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Tuple, Set

import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfgen import canvas

VERSION = "0.2"
__version__ = VERSION

CS_NAMES: Set[str] = {"cs", "c/s", "case", "cases", "case qty", "case quantity"}
SL_NAMES: Set[str] = {"sl", "sl.", "s.no", "sno", "sr", "sr.", "no", "no.", "#", "item"}
TOTAL_KEYWORDS: Set[str] = {"total", "sub total", "subtotal", "grand total", "net payable amt"}

INVOICE_PATTERNS: List[re.Pattern] = [
    re.compile(r"Bill\s*(?:No|NO|Number)?\s*:?\s*([A-Z0-9/-]+)", re.I),
    re.compile(r"Invoice\s*(?:No|NO|Number)?\s*:?\s*([A-Z0-9/-]+)", re.I),
]

# ── Logger Section ───────────────────────────────────────────────────────────

class TimestampLogger:
    """A custom logger that prefixes every message with [HH:MM:SS] and accumulates logs."""
    def __init__(self) -> None:
        self.logs: List[str] = []
        self.gui_callback: Optional[Callable[[str], None]] = None

    def info(self, msg: str) -> None:
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        for line in str(msg).split("\n"):
            formatted = f"[{timestamp}] {line}"
            self.logs.append(formatted)
            if self.gui_callback:
                self.gui_callback(formatted)
            else:
                try:
                    print(formatted)
                except UnicodeEncodeError:
                    # Safe fallback printing for Windows CP1252 consoles
                    enc = sys.stdout.encoding or 'ascii'
                    print(formatted.encode(enc, errors='replace').decode(enc))

logger = TimestampLogger()

# ── Validation Section ────────────────────────────────────────────────────────

@dataclass
class ValidationResult:
    source_file: str
    invoice_number: str
    retailer_name: str
    pages: int
    tables_found: int
    header_row: Optional[int]
    product_column: str
    upc_column: str
    mrp_column: str
    cs_column: str
    rows_seen: int
    rows_accepted: int
    rows_skipped: int
    total_rows: int
    empty_rows: int
    total_rows_skipped: int
    cs_parse_failures: int
    invoice_cs_total: float
    extracted_cs_total: float
    difference: float
    invoice_products: int
    extracted_products: int
    confidence_score: float = 0.0
    risk_level: str = "HIGH"
    warnings: List[str] = field(default_factory=list)
    processing_log: List[str] = field(default_factory=list)
    started: str = ""
    finished: str = ""
    elapsed: float = 0.0

@dataclass
class GlobalStats:
    execution_time: float = 0.0
    pages: int = 0
    tables: int = 0
    rows_seen: int = 0
    rows_accepted: int = 0
    rows_skipped: int = 0
    skipped_totals: int = 0
    parse_errors: int = 0
    avg_confidence: float = 0.0
    highest_confidence: float = 0.0
    lowest_confidence: float = 0.0
    warnings: int = 0
    stated_cs_total: float = 0.0
    extracted_cs_total: float = 0.0
    cumulative_difference: float = 0.0

class InvoiceAccumulator:
    """Accumulates invoice-specific metrics during page-by-page extraction."""
    def __init__(self, invoice_number: str, retailer_name: str, source_file: str) -> None:
        self.invoice_number: str = invoice_number
        self.retailer_name: str = retailer_name
        self.source_file: str = source_file
        self.pages: int = 1
        self.tables_found: int = 0
        self.header_row: Optional[int] = None
        self.product_column: str = "not found"
        self.upc_column: str = "not found"
        self.mrp_column: str = "not found"
        self.cs_column: str = "not found"
        self.rows_seen: int = 0
        self.rows_accepted: int = 0
        self.rows_skipped: int = 0
        self.empty_rows: int = 0
        self.total_rows_skipped: int = 0
        self.cs_parse_failures: int = 0
        self.invoice_cs_total: float = 0.0
        self.extracted_cs_total: float = 0.0
        self.warnings: List[str] = []
        self.log_start_idx: int = len(logger.logs)
        self.started: str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.start_time: float = time.time()
        self.detected_text_totals: List[float] = []
        self.detected_table_totals: List[float] = []

def calculate_confidence(validation: ValidationResult) -> Tuple[float, str]:
    """Calculates confidence score out of 100 based on weighted requirements."""
    score = 0.0
    if validation.invoice_number:
        score += 10
    if validation.tables_found > 0:
        score += 10
    if validation.header_row is not None:
        score += 10
    if validation.product_column and validation.product_column != "not found":
        score += 10
    if validation.cs_column and validation.cs_column != "not found":
        score += 15
    if validation.upc_column and validation.upc_column != "not found":
        score += 5
    if validation.mrp_column and validation.mrp_column != "not found":
        score += 5
    
    # Rows Accepted criteria: legitimate empty invoices get full credit
    if validation.rows_accepted > 0 or (validation.invoice_cs_total == 0 and validation.extracted_cs_total == 0):
        score += 15
        
    # Totals Match criteria: difference < 0.01 matches perfectly (including zero-totals)
    if abs(validation.difference) < 0.01:
        score += 15
        
    if validation.cs_parse_failures == 0:
        score += 5

    if score >= 96:
        risk = "LOW"
    elif score >= 90:
        risk = "MEDIUM"
    else:
        risk = "HIGH"
    return score, risk

def get_validation_status(val: ValidationResult) -> str:
    """Determines the status of a validation result: PASS, WARNING, or FAIL."""
    if val.confidence_score >= 96 and abs(val.difference) < 0.01 and val.cs_parse_failures == 0:
        return "PASS"
    elif val.confidence_score < 90:
        return "FAIL"
    else:
        return "WARNING"

def get_overall_status(validations: List[ValidationResult]) -> str:
    """Aggregates the status of all validations into an overall status."""
    if not validations:
        return "FAIL"
    statuses = [get_validation_status(v) for v in validations]
    if "FAIL" in statuses:
        return "FAIL"
    elif "WARNING" in statuses:
        return "WARNING"
    else:
        return "PASS"

# ── Extraction Section ────────────────────────────────────────────────────────

def safe_str(val: Any) -> str:
    """Convert any cell value (including None) to a clean string."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "none" else s

def extract_invoice_number(text: str) -> str:
    """Scans text for invoice numbers using regex patterns."""
    for p in INVOICE_PATTERNS:
        m = p.search(text or "")
        if m:
            return m.group(1).strip()
    return ""

def extract_retailer_and_invoice(text: str) -> Tuple[str, str]:
    """Heuristic extraction of Retailer Name and Invoice Number from text."""
    m = re.search(r"Retailer\s*Name\s*:\s*([^:\n]+?)(?:\s+(?:Invoice|Bill)\s*NO|\n|$)", text, re.I)
    retailer = m.group(1).strip() if m else ""
    
    if not retailer:
        m2 = re.search(r"([A-Z0-9\s\-\.&,\']+?)\s+(?:Bill|Invoice)\s*(?:No|NO|Number)?\s*:?\s*([A-Z0-9/-]+)", text, re.I)
        if m2:
            potential = m2.group(1).strip()
            lines = potential.split("\n")
            if lines:
                potential = lines[-1].strip()
            if not any(x in potential.lower() for x in ["tax invoice", "distributor", "enterprises", "page", "state", "fssai"]):
                retailer = potential

    inv = extract_invoice_number(text)
    return retailer, inv

def extract_total_cs_from_text(text: str) -> float:
    """Looks for total CS values stated in the raw page text."""
    pats = [
        r"Total\s*(?:Cases|Qty|Quantity|CS|C/S)?\s*:\s*([0-9\.,]+)",
        r"(?:Cases|Qty|Quantity|CS|C/S)\s*Total\s*:\s*([0-9\.,]+)",
        r"Total\s+([0-9\.,]+)\s*(?:Cases|CS|C/S|Qty)",
    ]
    for pat in pats:
        m = re.search(pat, text, re.I)
        if m:
            try:
                val = float(m.group(1).replace(",", ""))
                if val > 0:
                    return val
            except ValueError:
                pass
    return 0.0

def find_cs(headers: List[str]) -> Optional[int]:
    """Returns the index of the CS/Case quantity column."""
    for i, h in enumerate(headers):
        n = h.lower().strip().replace(".", "").replace("_", " ")
        if n in CS_NAMES:
            return i
    return None

def find_sl(headers: List[str]) -> Optional[int]:
    """Returns the index of the Serial Number column."""
    for i, h in enumerate(headers):
        n = h.lower().strip().replace(".", "").replace("_", " ")
        if n in SL_NAMES:
            return i
    return None

def is_total_row(vals: List[str], sl_idx: Optional[int]) -> bool:
    """Identifies if a row is a total/summary row."""
    if sl_idx is not None and not vals[sl_idx].strip():
        return True
    for v in vals:
        if v.lower().strip() in TOTAL_KEYWORDS:
            return True
    return False

def parse_header_row(row: List[Any]) -> List[str]:
    """Build headers from a raw table row, safely handling None and duplicates."""
    headers = []
    seen: Dict[str, int] = {}
    for i, cell in enumerate(row):
        h = safe_str(cell) or f"Column_{i+1}"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 1
        headers.append(h)
    return headers

def parse_data_row(row: List[Any], n_cols: int) -> List[str]:
    """Convert a data row to strings, pad/trim to n_cols."""
    vals = [safe_str(c) for c in row]
    while len(vals) < n_cols:
        vals.append("")
    return vals[:n_cols]

def find_header_row(table: List[Any]) -> Optional[int]:
    """Scan rows top-down and return index of first row containing CS column."""
    for i, row in enumerate(table):
        headers = parse_header_row(row)
        if find_cs(headers) is not None:
            return i
    return None

def finalize_accumulator(acc: InvoiceAccumulator, all_validations: List[ValidationResult]) -> ValidationResult:
    """Calculates final values for an accumulator and generates a ValidationResult."""
    if acc.detected_table_totals:
        acc.invoice_cs_total = max(acc.detected_table_totals)
    elif acc.detected_text_totals:
        acc.invoice_cs_total = max(acc.detected_text_totals)
    else:
        acc.invoice_cs_total = 0.0

    acc.difference = acc.extracted_cs_total - acc.invoice_cs_total
    elapsed = time.time() - acc.start_time
    
    validation = ValidationResult(
        source_file=acc.source_file,
        invoice_number=acc.invoice_number,
        retailer_name=acc.retailer_name or "Unknown Retailer",
        pages=acc.pages,
        tables_found=acc.tables_found,
        header_row=acc.header_row,
        product_column=acc.product_column,
        upc_column=acc.upc_column,
        mrp_column=acc.mrp_column,
        cs_column=acc.cs_column,
        rows_seen=acc.rows_seen,
        rows_accepted=acc.rows_accepted,
        rows_skipped=acc.rows_skipped,
        total_rows=acc.rows_seen + acc.empty_rows + acc.total_rows_skipped + acc.cs_parse_failures,
        empty_rows=acc.empty_rows,
        total_rows_skipped=acc.total_rows_skipped,
        cs_parse_failures=acc.cs_parse_failures,
        invoice_cs_total=acc.invoice_cs_total,
        extracted_cs_total=acc.extracted_cs_total,
        difference=acc.difference,
        invoice_products=acc.rows_accepted + acc.cs_parse_failures + acc.rows_skipped,
        extracted_products=acc.rows_accepted,
        started=acc.started,
        finished=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        elapsed=elapsed,
        warnings=list(acc.warnings),
        processing_log=logger.logs[acc.log_start_idx:]
    )
    
    score, risk = calculate_confidence(validation)
    validation.confidence_score = score
    validation.risk_level = risk
    
    if score < 95:
        validation.warnings.append(f"Confidence score is low: {score:.1f}%")
    if abs(validation.difference) > 0.01:
        validation.warnings.append(f"Total CS mismatch: Stated={validation.invoice_cs_total}, Extracted={validation.extracted_cs_total}")
    if validation.cs_parse_failures > 0:
        validation.warnings.append(f"{validation.cs_parse_failures} CS parse failure(s) encountered.")
        
    all_validations.append(validation)
    return validation

def extract_from_pdf(
    pdf_path: Path,
    on_page: Optional[Callable[[int, int], None]] = None
) -> Tuple[List[Dict[str, Any]], List[ValidationResult], List[Dict[str, Any]]]:
    """Extracts rows from PDF and populates ValidationResult and SkippedRow collections."""
    rows_out: List[Dict[str, Any]] = []
    validation_results: List[ValidationResult] = []
    skipped_rows: List[Dict[str, Any]] = []

    filename = pdf_path.name
    display_name = filename[:15] + "...." if len(filename) > 15 else filename

    logger.info("Opening File")
    logger.info(display_name)

    current_retailer = ""
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        logger.info(f"Pages: {total_pages}")

        current_inv_num = ""
        acc: Optional[InvoiceAccumulator] = None

        for page_num, page in enumerate(pdf.pages, 1):
            if on_page:
                on_page(page_num, total_pages)

            page_text = page.extract_text() or ""
            ret, inv = extract_retailer_and_invoice(page_text)
            if ret:
                current_retailer = ret

            if inv and inv != current_inv_num:
                if acc is not None:
                    finalize_accumulator(acc, validation_results)
                current_inv_num = inv
                acc = InvoiceAccumulator(
                    invoice_number=inv,
                    retailer_name=ret or current_retailer or "Unknown Retailer",
                    source_file=pdf_path.name
                )
                logger.info("Invoice")
                logger.info(inv)
                logger.info(f"Page {page_num}/{total_pages}")
            elif not current_inv_num:
                current_inv_num = inv or pdf_path.stem
                acc = InvoiceAccumulator(
                    invoice_number=current_inv_num,
                    retailer_name=ret or "Unknown Retailer",
                    source_file=pdf_path.name
                )
                logger.info("Invoice")
                logger.info(current_inv_num)
                logger.info(f"Page {page_num}/{total_pages}")
            else:
                if acc is not None:
                    acc.pages += 1
                logger.info(f"Page {page_num}/{total_pages}")

            if ret and acc:
                acc.retailer_name = ret

            txt_total = extract_total_cs_from_text(page_text)
            if txt_total > 0 and acc:
                acc.detected_text_totals.append(txt_total)

            tables = page.extract_tables() or []
            if not tables:
                logger.info("No tables found")
                continue

            logger.info(f"{len(tables)} table(s) found")

            for t_idx, table in enumerate(tables, 1):
                if not table or len(table) < 2:
                    reason = f"Too few rows ({len(table) if table else 0})"
                    if acc:
                        acc.tables_found += 1
                        acc.total_rows_skipped += len(table) if table else 0
                    for r_num in range(len(table) if table else 0):
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": r_num + 1,
                            "Reason": reason,
                            "Raw Data": str(table[r_num]) if table else ""
                        })
                    continue

                header_row_idx = find_header_row(table)
                if header_row_idx is None:
                    reason = "No CS column found in table"
                    if acc:
                        acc.tables_found += 1
                        acc.total_rows_skipped += len(table)
                    for r_num in range(len(table)):
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": r_num + 1,
                            "Reason": reason,
                            "Raw Data": str(table[r_num])
                        })
                    continue

                if acc:
                    acc.tables_found += 1
                    if acc.header_row is None or header_row_idx > acc.header_row:
                        acc.header_row = header_row_idx

                headers = parse_header_row(table[header_row_idx])
                n_cols = len(headers)
                cs_idx = find_cs(headers)
                sl_idx = find_sl(headers)

                prod_idx = next((i for i, h in enumerate(headers)
                                 if "product" in h.lower() and "name" in h.lower()), None)
                upc_idx  = next((i for i, h in enumerate(headers)
                                 if h.lower().strip().replace(".", "")
                                 in {"upc", "upc code", "barcode"}), None)
                mrp_idx  = next((i for i, h in enumerate(headers)
                                 if h.lower().strip().replace(".", "").replace(" ", "")
                                 in {"mrp", "mrp rs", "mrprs"}), None)

                if acc:
                    if cs_idx is not None:
                        acc.cs_column = headers[cs_idx]
                    if prod_idx is not None:
                        acc.product_column = headers[prod_idx]
                    if upc_idx is not None:
                        acc.upc_column = headers[upc_idx]
                    if mrp_idx is not None:
                        acc.mrp_column = headers[mrp_idx]

                data_rows = table[header_row_idx + 1:]
                
                if header_row_idx > 0 and acc:
                    acc.total_rows_skipped += header_row_idx
                    for r_num in range(header_row_idx):
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": r_num + 1,
                            "Reason": "Preamble row skipped",
                            "Raw Data": str(table[r_num])
                        })

                for offset, raw_row in enumerate(data_rows, 1):
                    row_num = header_row_idx + 1 + offset
                    if not raw_row or not any(safe_str(c) for c in raw_row):
                        if acc:
                            acc.empty_rows += 1
                        continue

                    if acc:
                        acc.rows_seen += 1

                    vals = parse_data_row(raw_row, n_cols)

                    if is_total_row(vals, sl_idx):
                        if acc:
                            acc.total_rows_skipped += 1
                            if cs_idx is not None:
                                try:
                                    t_cs = float(vals[cs_idx].replace(",", ""))
                                    if t_cs > 0:
                                        acc.detected_table_totals.append(t_cs)
                                except ValueError:
                                    pass
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": row_num,
                            "Reason": "Total row skipped",
                            "Raw Data": str(raw_row)
                        })
                        continue

                    if cs_idx is None:
                        if acc:
                            acc.rows_skipped += 1
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": row_num,
                            "Reason": "CS column not defined",
                            "Raw Data": str(raw_row)
                        })
                        continue

                    raw_cs = vals[cs_idx]
                    raw_cs_strip = raw_cs.strip()
                    if not raw_cs_strip or raw_cs_strip in {"-", "—", "–", "."}:
                        if acc:
                            acc.rows_skipped += 1
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": row_num,
                            "Reason": "Blank CS value",
                            "Raw Data": str(raw_row)
                        })
                        continue

                    try:
                        cs = float(raw_cs.replace(",", ""))
                    except (ValueError, AttributeError):
                        if acc:
                            acc.cs_parse_failures += 1
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": row_num,
                            "Reason": f"CS parse failed on {raw_cs!r}",
                            "Raw Data": str(raw_row)
                        })
                        continue

                    if cs <= 0:
                        if acc:
                            acc.rows_skipped += 1
                        skipped_rows.append({
                            "Invoice": current_inv_num,
                            "Page": page_num,
                            "Table": t_idx,
                            "Row": row_num,
                            "Reason": "CS value <= 0",
                            "Raw Data": str(raw_row)
                        })
                        continue

                    product = vals[prod_idx] if prod_idx is not None else ""
                    upc     = vals[upc_idx]  if upc_idx  is not None else ""
                    mrp     = vals[mrp_idx]  if mrp_idx  is not None else ""

                    if not product and acc:
                        acc.warnings.append(f"Row {row_num}: Product Name empty (CS={cs}, UPC={upc})")

                    rows_out.append({
                        "Invoice Number": current_inv_num,
                        "Product Name":   product,
                        "UPC":            upc,
                        "MRP":            mrp,
                        "CS":             cs,
                    })
                    if acc:
                        acc.rows_accepted += 1
                        acc.extracted_cs_total += cs

        if acc is not None:
            finalize_accumulator(acc, validation_results)

    for val in validation_results:
        status = get_validation_status(val)
        total_skipped = val.rows_skipped + val.total_rows_skipped + val.cs_parse_failures
        
        if status == "PASS":
            logger.info(
                f"✓ [{status}] Invoice {val.invoice_number} | "
                f"Retailer: {val.retailer_name} | "
                f"Confidence: {int(val.confidence_score)}% | "
                f"Accepted: {val.rows_accepted} | "
                f"Skipped: {total_skipped}"
            )
        else:
            icon = "⚠" if status == "WARNING" else "❌"
            logger.info(
                f"{icon} [{status}] Invoice {val.invoice_number} | "
                f"Retailer: {val.retailer_name} | "
                f"Confidence: {int(val.confidence_score)}% (Risk: {val.risk_level})"
            )
            logger.info(
                f"  └─ Accepted: {val.rows_accepted} rows | "
                f"Skipped: {total_skipped} rows (including {val.cs_parse_failures} CS parse failures)"
            )
            
            # Log mapping mismatches
            mismatches = []
            if not val.invoice_number:
                mismatches.append("Missing Invoice Number")
            if val.header_row is None:
                mismatches.append("Missing Header Row")
            if val.product_column == "not found":
                mismatches.append("Product Column not found")
            if val.cs_column == "not found":
                mismatches.append("CS Column not found")
            if val.upc_column == "not found":
                mismatches.append("UPC Column not found")
            if val.mrp_column == "not found":
                mismatches.append("MRP Column not found")
            
            if mismatches:
                logger.info(f"  └─ Element Mapping Mismatches: {', '.join(mismatches)}")
                
            if abs(val.difference) >= 0.01:
                logger.info(
                    f"  └─ Totals Mismatch! Stated CS Total: {val.invoice_cs_total:.2f} | "
                    f"Extracted CS Total: {val.extracted_cs_total:.2f} (Diff: {val.difference:.2f})"
                )
                
            for w in val.warnings:
                if "confidence" not in w.lower():
                    logger.info(f"  └─ Alert: {w}")

    logger.info(f"Done: {len(rows_out)} CS rows extracted")
    return rows_out, validation_results, skipped_rows

# ── Statistics Section ────────────────────────────────────────────────────────

def process(
    pdfs: List[Path],
    log: Any = None, # Signature compatibility
    on_progress: Optional[Callable[[float, str], None]] = None,
    save_path: Optional[str] = None,
    fmt: str = "xlsx",
    gui_update_callback: Optional[Callable[..., None]] = None,
    detailed_pdf: bool = False
) -> Optional[Path]:
    """Main pipeline execution for single or multiple PDFs."""
    logger.logs.clear()
    start_time = datetime.datetime.now()

    all_rows: List[Dict[str, Any]] = []
    all_validations: List[ValidationResult] = []
    all_skipped_rows: List[Dict[str, Any]] = []

    total_files = len(pdfs)
    for idx, pdf in enumerate(pdfs):
        logger.info(f"ℹ Loading {pdf.name}")
        try:
            def on_page_progress(page_num: int, total_pages: int, _idx=idx, _total=total_files):
                if on_progress:
                    pct = ((_idx + page_num / total_pages) / _total) * 100
                    on_progress(
                        pct,
                        f"Processing file {_idx+1}/{_total} | {pdf.name} | Page {page_num}/{total_pages}"
                    )
            
            rows, validations, skipped = extract_from_pdf(pdf, on_page=on_page_progress)
            all_rows.extend(rows)
            all_validations.extend(validations)
            all_skipped_rows.extend(skipped)
            
            # Periodic GUI updates
            if gui_update_callback:
                confidences = [v.confidence_score for v in all_validations]
                cur_avg_conf = sum(confidences) / len(confidences) if confidences else 0.0
                cur_warnings = sum(len(v.warnings) for v in all_validations)
                cur_products = sum(v.rows_accepted for v in all_validations)
                gui_update_callback(
                    files_count=idx + 1,
                    invoices_count=len(all_validations),
                    products_count=cur_products,
                    unique_products_count=len(all_rows),
                    current_invoice=validations[-1].invoice_number if validations else "-",
                    avg_confidence=cur_avg_conf,
                    warnings_count=cur_warnings,
                    elapsed=(datetime.datetime.now() - start_time).total_seconds(),
                    validations=all_validations
                )
        except Exception as e:
            logger.info(f"⚠ Error in {pdf.name}: {e}")

    if not all_rows:
        logger.info("No CS rows found.")
        return None

    if on_progress:
        on_progress(100, "Saving PDF Report…" if fmt == "pdf" else "Saving Excel Spreadsheet…")

    df = pd.DataFrame(all_rows, columns=["Invoice Number", "Product Name", "UPC", "MRP", "CS"])
    df["MRP"] = pd.to_numeric(df["MRP"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    df["CS"]  = pd.to_numeric(df["CS"].astype(str).str.replace(",", "", regex=False),  errors="coerce").fillna(0)

    before = len(df)
    df_merged = (
        df.groupby(["Product Name", "UPC", "MRP"], as_index=False)
          .agg(**{
              "Invoice Numbers": (
                  "Invoice Number",
                  lambda s: ", ".join(sorted({
                      v.strip() for v in s.astype(str)
                      if v.strip() and v.lower() != "nan"
                  }))
              ),
              "CS": ("CS", "sum"),
          })
    )
    logger.info(f"→ merged {before} rows → {len(df_merged)} unique products")
    df_merged = df_merged[["Invoice Numbers", "Product Name", "UPC", "MRP", "CS"]]

    elapsed_total = (datetime.datetime.now() - start_time).total_seconds()

    # Compile global statistics
    total_pages = sum(v.pages for v in all_validations)
    total_tables = sum(v.tables_found for v in all_validations)
    total_rows_seen = sum(v.rows_seen for v in all_validations)
    total_rows_accepted = sum(v.rows_accepted for v in all_validations)
    total_rows_skipped = sum(v.rows_skipped for v in all_validations)
    total_skipped_totals = sum(v.total_rows_skipped for v in all_validations)
    total_parse_errors = sum(v.cs_parse_failures for v in all_validations)
    
    confidences = [v.confidence_score for v in all_validations]
    avg_confidence = sum(confidences) / len(confidences) if confidences else 0.0
    highest_confidence = max(confidences) if confidences else 0.0
    lowest_confidence = min(confidences) if confidences else 0.0
    total_warnings = sum(len(v.warnings) for v in all_validations)

    sum_stated_cs = sum(v.invoice_cs_total for v in all_validations)
    sum_extracted_cs = sum(v.extracted_cs_total for v in all_validations)
    cum_diff = sum_extracted_cs - sum_stated_cs
    status_str = "PASS" if abs(cum_diff) < 0.01 else "MISMATCH"
    logger.info(
        f"📊 Cumulative Verification | Stated Total CS: {sum_stated_cs:.2f} | "
        f"Extracted Total CS: {sum_extracted_cs:.2f} | "
        f"Difference: {cum_diff:.2f} | Status: {status_str}"
    )

    global_stats = GlobalStats(
        execution_time=elapsed_total,
        pages=total_pages,
        tables=total_tables,
        rows_seen=total_rows_seen,
        rows_accepted=total_rows_accepted,
        rows_skipped=total_rows_skipped,
        skipped_totals=total_skipped_totals,
        parse_errors=total_parse_errors,
        avg_confidence=avg_confidence,
        highest_confidence=highest_confidence,
        lowest_confidence=lowest_confidence,
        warnings=total_warnings,
        stated_cs_total=sum_stated_cs,
        extracted_cs_total=sum_extracted_cs,
        cumulative_difference=cum_diff
    )

    if gui_update_callback:
        gui_update_callback(
            files_count=len(pdfs),
            invoices_count=len(all_validations),
            products_count=total_rows_accepted,
            unique_products_count=len(df_merged),
            current_invoice="",
            avg_confidence=avg_confidence,
            warnings_count=total_warnings,
            elapsed=elapsed_total,
            validations=all_validations
        )

    if save_path is None:
        try:
            root = tk.Tk()
            root.withdraw()
            stem = pdfs[0].parent / (pdfs[0].parent / (pdfs[0].stem if len(pdfs) == 1 else "Combined"))
            ext  = ".pdf" if fmt == "pdf" else ".xlsx"
            default_name = f"{stem.name}_CS_Extract{ext}"
            ftypes = ([("PDF Document", "*.pdf")] if fmt == "pdf"
                      else [("Excel Workbook", "*.xlsx")])
            save_path = filedialog.asksaveasfilename(
                title=f"Save {'PDF Report' if fmt == 'pdf' else 'Excel File'}",
                defaultextension=ext,
                initialfile=default_name,
                filetypes=ftypes,
            )
            root.destroy()
        except Exception:
            stem = pdfs[0].parent / (pdfs[0].stem if len(pdfs) == 1 else "Combined")
            ext  = ".pdf" if fmt == "pdf" else ".xlsx"
            save_path = str(stem) + f"_CS_Extract{ext}"

        if not save_path:
            logger.info("⚠ Save cancelled.")
            return None

    outfile = Path(save_path)
    if fmt == "pdf":
        save_as_pdf_v02(
            df_merged,
            outfile,
            [p.name for p in pdfs],
            all_validations,
            all_skipped_rows,
            global_stats,
            detailed_pdf=detailed_pdf
        )
    else:
        df_merged.to_excel(outfile, index=False)

    logger.info(f"✓ Saved → {outfile.resolve()}")
    return outfile

# ── PDF Generation Section ───────────────────────────────────────────────────

class TOCRegister(Flowable):
    """Hidden flowable to register page numbers of sections in NumberedCanvas."""
    def __init__(self, section_name: str) -> None:
        super().__init__()
        self.section_name: str = section_name
        self.width = 0
        self.height = 0

    def draw(self) -> None:
        if hasattr(self, "canv") and self.canv is not None:
            # Call the register method on our NumberedCanvas
            register_fn = getattr(self.canv, "register_section_page", None)
            if register_fn:
                register_fn(self.section_name)

class NumberedCanvas(canvas.Canvas):
    """Two-pass ReportLab Canvas that draws footers and builds Table of Contents."""
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._saved_page_states: List[dict] = []
        self.section_pages: Dict[str, int] = {}
        
    def register_section_page(self, name: str) -> None:
        self.section_pages[name] = self._pageNumber

    def showPage(self) -> None:
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, num_pages: int) -> None:
        self.saveState()
        
        # Set report properties
        self.setCreator("CS Extractor v0.2")
        self.setAuthor("CS Extractor")
        self.setTitle("CS Extraction Report")

        # Footer
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#4b5563"))
        footer_text = f"Generated by CS Extractor v0.2   |   Page {self._pageNumber} of {num_pages}"
        self.drawCentredString(842 / 2.0, 25, footer_text)
        
        # Developed by + logo icon
        if hasattr(sys, "_MEIPASS"):
            logo_path = Path(sys._MEIPASS) / "logo.png"
        else:
            logo_path = Path("logo.png")
            if not logo_path.exists():
                logo_path = Path(__file__).parent / "logo.png"
            
        if logo_path.exists():
            self.drawString(50, 25, "Developed by ")
            dev_w = self.stringWidth("Developed by ", "Helvetica", 8)
            self.drawImage(str(logo_path), 50 + dev_w + 2, 24, width=20, height=20, mask='auto')
        
        # Render Table of Contents on its registered page
        toc_page = self.section_pages.get("Table of Contents")
        if toc_page and self._pageNumber == toc_page:
            self.draw_toc()
            
        self.restoreState()

    def draw_toc(self) -> None:
        self.setFont("Helvetica-Bold", 18)
        self.setFillColor(colors.HexColor("#1f2937"))
        self.drawString(50, 520, "Table of Contents")
        
        self.setStrokeColor(colors.HexColor("#d1d5db"))
        self.setLineWidth(0.5)
        self.line(50, 510, 792, 510)
        
        self.setFont("Helvetica", 10)
        self.setFillColor(colors.HexColor("#374151"))
        
        toc_items = [
            ("Products Table", self.section_pages.get("Products", 1)),
            ("Cover Page", self.section_pages.get("Cover", 2)),
            ("Executive Summary", self.section_pages.get("Summary", 3)),
            ("Table of Contents", self.section_pages.get("Table of Contents", 4)),
            ("Annexure A: Extraction Confidence Report", self.section_pages.get("Annexure A", 5)),
            ("Annexure B: Validation Summary", self.section_pages.get("Annexure B", 6)),
            ("Annexure C: Processing Statistics", self.section_pages.get("Annexure C", 7)),
        ]
        
        if getattr(self, "detailed_pdf", True):
            toc_items.extend([
                ("Annexure D: Detailed Processing Log", self.section_pages.get("Annexure D", 8)),
                ("Annexure E: Skipped Rows", self.section_pages.get("Annexure E", 9)),
                ("Annexure F: Warnings Summary", self.section_pages.get("Annexure F", 10)),
            ])
        else:
            toc_items.extend([
                ("Annexure F: Warnings Summary", self.section_pages.get("Annexure F", 8)),
            ])
        
        y = 470
        for name, page in toc_items:
            self.drawString(50, y, name)
            page_str = f"Page {page}"
            self.drawRightString(792, y, page_str)
            
            # Dynamic dot leader calculation
            dots_start = 320
            page_w = self.stringWidth(page_str, "Helvetica", 10)
            dots_end = 792 - page_w - 5
            avail_w = dots_end - dots_start
            dot_w = self.stringWidth(".", "Helvetica", 10)
            num_dots = int(avail_w / dot_w)
            dots = "." * num_dots if num_dots > 0 else ""
            self.drawString(dots_start, y, dots)
            y -= 25

def save_as_pdf_v02(
    df_merged: pd.DataFrame,
    outfile: Path,
    source_names: List[str],
    validations: List[ValidationResult],
    skipped_rows: List[Dict[str, Any]],
    stats: GlobalStats,
    detailed_pdf: bool = True
) -> None:
    """Generates the multi-page validation and extraction PDF report."""
    story: List[Any] = []
    styles = getSampleStyleSheet()

    # Style sheet setup
    title_style = ParagraphStyle(
        "CoverTitle",
        parent=styles["Normal"],
        fontSize=28,
        leading=34,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1f2937"),
        alignment=TA_CENTER,
        spaceAfter=15
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Normal"],
        fontSize=18,
        leading=22,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1f2937"),
        spaceAfter=15
    )
    meta_label_style = ParagraphStyle(
        "MetaLabel",
        parent=styles["Normal"],
        fontSize=10,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#4b5563")
    )
    meta_val_style = ParagraphStyle(
        "MetaVal",
        parent=styles["Normal"],
        fontSize=10,
        fontName="Helvetica",
        textColor=colors.HexColor("#1f2937")
    )
    header_style_small = ParagraphStyle(
        "HeaderStyleSmall",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#374151"),
        alignment=TA_CENTER
    )
    val_style_small = ParagraphStyle(
        "ValStyleSmall",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        fontName="Helvetica",
        textColor=colors.HexColor("#1f2937"),
        alignment=TA_CENTER
    )

    # ── Table Pre-constructions ──────────────────────────────────────────────

    # Cover Page Table setup
    overall_status = get_overall_status(validations)
    status_color = "#10b981" if overall_status == "PASS" else "#f59e0b" if overall_status == "WARNING" else "#ef4444"
    status_html = f"<b><font color='{status_color}'>{overall_status}</font></b>"

    source_names_str = ", ".join(source_names)
    source_files_style = ParagraphStyle(
        "CoverSourceFiles",
        parent=meta_val_style,
        fontSize=8,
        leading=10
    )

    cover_data = [
        [Paragraph("Generated By:", meta_label_style), Paragraph("CS Extractor", meta_val_style)],
        [Paragraph("Version:", meta_label_style), Paragraph("0.2", meta_val_style)],
        [Paragraph("Generated:", meta_label_style), Paragraph(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), meta_val_style)],
        [Paragraph("Source Files:", meta_label_style), Paragraph(source_names_str, source_files_style)],
        [Paragraph("Products:", meta_label_style), Paragraph(str(stats.rows_accepted), meta_val_style)],
        [Paragraph("Invoices:", meta_label_style), Paragraph(str(len(validations)), meta_val_style)],
        [Paragraph("Average Confidence:", meta_label_style), Paragraph(f"{stats.avg_confidence:.1f}%", meta_val_style)],
        [Paragraph("Overall Status:", meta_label_style), Paragraph(status_html, meta_val_style)],
    ]
    cover_table = Table(cover_data, colWidths=[200, 400])
    cover_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#e5e7eb")),
    ]))

    # Decorative line for Cover Page
    line_table = Table([[""]], colWidths=[600], rowHeights=[3])
    line_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#374151")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
    ]))

    # Executive Summary Table setup
    cum_status_color = "#10b981" if abs(stats.cumulative_difference) < 0.01 else "#ef4444"
    cum_status_text = "PASS" if abs(stats.cumulative_difference) < 0.01 else f"MISMATCH ({stats.cumulative_difference:+.2f})"
    cum_status_html = f"<b><font color='{cum_status_color}'>{cum_status_text}</font></b>"

    summary_data = [
        [Paragraph("<b>Files Processed</b>", meta_label_style), Paragraph(str(len(source_names)), meta_val_style),
         Paragraph("<b>Warnings</b>", meta_label_style), Paragraph(str(stats.warnings), meta_val_style)],
        [Paragraph("<b>Total Invoices</b>", meta_label_style), Paragraph(str(len(validations)), meta_val_style),
         Paragraph("<b>Execution Time</b>", meta_label_style), Paragraph(f"{stats.execution_time:.2f}s", meta_val_style)],
        [Paragraph("<b>Total Products</b>", meta_label_style), Paragraph(str(stats.rows_accepted), meta_val_style),
         Paragraph("<b>Overall Validation</b>", meta_label_style), Paragraph(status_html, meta_val_style)],
        [Paragraph("<b>Unique Products</b>", meta_label_style), Paragraph(str(len(df_merged)), meta_val_style),
         Paragraph("<b>Average Confidence</b>", meta_label_style), Paragraph(f"{stats.avg_confidence:.1f}%", meta_val_style)],
        [Paragraph("<b>Stated Cumulative CS</b>", meta_label_style), Paragraph(f"{stats.stated_cs_total:.2f}", meta_val_style),
         Paragraph("<b>Cumulative Verification</b>", meta_label_style), Paragraph(cum_status_html, meta_val_style)],
    ]
    summary_table = Table(summary_data, colWidths=[150, 150, 150, 150])
    summary_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f9fafb")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e5e7eb")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))

    # Products Table setup
    col_widths = [150, 350, 80, 80, 97]
    prod_header_row = [
        Paragraph("<b>Invoice Numbers</b>", meta_label_style),
        Paragraph("<b>Product Name</b>", meta_label_style),
        Paragraph("<b>UPC</b>", meta_label_style),
        Paragraph("<b>MRP</b>", meta_label_style),
        Paragraph("<b>CS</b>", meta_label_style),
    ]
    prod_data_rows = []
    for r in df_merged.itertuples(index=False):
        prod_data_rows.append([
            Paragraph(str(r[0]), meta_val_style),
            Paragraph(str(r[1]), meta_val_style),
            Paragraph(str(r[2]), meta_val_style),
            Paragraph(f"{r[3]:.2f}" if pd.notna(r[3]) else "", meta_val_style),
            Paragraph(str(int(r[4])) if r[4].is_integer() else f"{r[4]:.2f}", meta_val_style),
        ])
        
    # Build Products Total Row
    total_cs_sum = df_merged["CS"].sum()
    total_cs_str = str(int(total_cs_sum)) if total_cs_sum.is_integer() else f"{total_cs_sum:.2f}"
    prod_total_row = [
        Paragraph("<b>Total</b>", meta_label_style),
        Paragraph("", meta_val_style),
        Paragraph("", meta_val_style),
        Paragraph("", meta_val_style),
        Paragraph(f"<b>{total_cs_str}</b>", meta_label_style),
    ]
    
    prod_table_data = [prod_header_row] + prod_data_rows + [prod_total_row]
    prod_table = Table(prod_table_data, colWidths=col_widths, repeatRows=1)
    prod_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d1d5db")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#f9fafb")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#e5e7eb")),
        ("LINEABOVE", (0, -1), (-1, -1), 1.0, colors.HexColor("#374151")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    # Annexure A Table setup
    col_widths_a = [85, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 72]
    annex_a_header = [
        Paragraph("<b>Invoice</b>", header_style_small),
        Paragraph("<b>Confidence</b>", header_style_small),
        Paragraph("<b>Risk</b>", header_style_small),
        Paragraph("<b>Header</b>", header_style_small),
        Paragraph("<b>Product</b>", header_style_small),
        Paragraph("<b>UPC</b>", header_style_small),
        Paragraph("<b>MRP</b>", header_style_small),
        Paragraph("<b>CS</b>", header_style_small),
        Paragraph("<b>Rows Seen</b>", header_style_small),
        Paragraph("<b>Rows Acc.</b>", header_style_small),
        Paragraph("<b>Rows Skip</b>", header_style_small),
        Paragraph("<b>Validation</b>", header_style_small),
    ]
    annex_a_rows = []
    for val in validations:
        status = get_validation_status(val)
        s_color = "#10b981" if status == "PASS" else "#f59e0b" if status == "WARNING" else "#ef4444"
        status_html = f"<b><font color='{s_color}'>{status}</font></b>"
        
        r_color = "#10b981" if val.risk_level == "LOW" else "#f59e0b" if val.risk_level == "MEDIUM" else "#ef4444"
        risk_html = f"<b><font color='{r_color}'>{val.risk_level}</font></b>"
        
        def yes_no(col_name: str) -> str:
            return "Yes" if col_name and col_name != "not found" else "No"
            
        annex_a_rows.append([
            Paragraph(val.invoice_number, val_style_small),
            Paragraph(f"{val.confidence_score:.1f}%", val_style_small),
            Paragraph(risk_html, val_style_small),
            Paragraph("Yes" if val.header_row is not None else "No", val_style_small),
            Paragraph(yes_no(val.product_column), val_style_small),
            Paragraph(yes_no(val.upc_column), val_style_small),
            Paragraph(yes_no(val.mrp_column), val_style_small),
            Paragraph(yes_no(val.cs_column), val_style_small),
            Paragraph(str(val.rows_seen), val_style_small),
            Paragraph(str(val.rows_accepted), val_style_small),
            Paragraph(str(val.rows_skipped + val.total_rows_skipped + val.cs_parse_failures), val_style_small),
            Paragraph(status_html, val_style_small),
        ])
    annex_a_table_data = [annex_a_header] + annex_a_rows
    annex_a_table = Table(annex_a_table_data, colWidths=col_widths_a, repeatRows=1)
    annex_a_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d1d5db")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f9fafb")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    # Annexure B Table setup
    col_widths_b = [200, 120, 120, 120, 197]
    annex_b_header = [
        Paragraph("<b>Invoice</b>", header_style_small),
        Paragraph("<b>Invoice CS Total</b>", header_style_small),
        Paragraph("<b>Extracted CS Total</b>", header_style_small),
        Paragraph("<b>Difference</b>", header_style_small),
        Paragraph("<b>Status</b>", header_style_small),
    ]
    annex_b_rows = []
    for val in validations:
        status = get_validation_status(val)
        s_color = "#10b981" if status == "PASS" else "#f59e0b" if status == "WARNING" else "#ef4444"
        status_html = f"<b><font color='{s_color}'>{status}</font></b>"
        
        annex_b_rows.append([
            Paragraph(val.invoice_number, val_style_small),
            Paragraph(f"{val.invoice_cs_total:.2f}", val_style_small),
            Paragraph(f"{val.extracted_cs_total:.2f}", val_style_small),
            Paragraph(f"{val.difference:.2f}", val_style_small),
            Paragraph(status_html, val_style_small),
        ])
    annex_b_table_data = [annex_b_header] + annex_b_rows
    annex_b_table = Table(annex_b_table_data, colWidths=col_widths_b, repeatRows=1)
    annex_b_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d1d5db")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f9fafb")]),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    # Annexure C Table setup
    stats_data = [
        [Paragraph("<b>Files Processed</b>", meta_label_style), Paragraph(str(len(source_names)), meta_val_style)],
        [Paragraph("<b>Total Invoices Found</b>", meta_label_style), Paragraph(str(len(validations)), meta_val_style)],
        [Paragraph("<b>Total Pages</b>", meta_label_style), Paragraph(str(stats.pages), meta_val_style)],
        [Paragraph("<b>Total Tables Found</b>", meta_label_style), Paragraph(str(stats.tables), meta_val_style)],
        [Paragraph("<b>Total Rows Seen</b>", meta_label_style), Paragraph(str(stats.rows_seen), meta_val_style)],
        [Paragraph("<b>Total Rows Accepted</b>", meta_label_style), Paragraph(str(stats.rows_accepted), meta_val_style)],
        [Paragraph("<b>Total Rows Skipped (due to CS &lt;= 0)</b>", meta_label_style), Paragraph(str(stats.rows_skipped), meta_val_style)],
        [Paragraph("<b>Total Rows Skipped (due to Totals/Header/Preamble)</b>", meta_label_style), Paragraph(str(stats.skipped_totals), meta_val_style)],
        [Paragraph("<b>CS Parse Errors</b>", meta_label_style), Paragraph(str(stats.parse_errors), meta_val_style)],
        [Paragraph("<b>Average Confidence Score</b>", meta_label_style), Paragraph(f"{stats.avg_confidence:.2f}%", meta_val_style)],
        [Paragraph("<b>Highest Confidence Score</b>", meta_label_style), Paragraph(f"{stats.highest_confidence:.2f}%", meta_val_style)],
        [Paragraph("<b>Lowest Confidence Score</b>", meta_label_style), Paragraph(f"{stats.lowest_confidence:.2f}%", meta_val_style)],
        [Paragraph("<b>Stated Cumulative CS Total</b>", meta_label_style), Paragraph(f"{stats.stated_cs_total:.2f}", meta_val_style)],
        [Paragraph("<b>Extracted Cumulative CS Total</b>", meta_label_style), Paragraph(f"{stats.extracted_cs_total:.2f}", meta_val_style)],
        [Paragraph("<b>Cumulative CS Difference</b>", meta_label_style), Paragraph(f"{stats.cumulative_difference:.2f}", meta_val_style)],
        [Paragraph("<b>Execution Time</b>", meta_label_style), Paragraph(f"{stats.execution_time:.2f}s", meta_val_style)],
    ]
    stats_table = Table(stats_data, colWidths=[300, 457])
    stats_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e5e7eb")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#f9fafb"), colors.white]),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))

    # Annexure D Table setup (Detailed Logs)
    log_style = ParagraphStyle(
        "LogStyle",
        parent=styles["Normal"],
        fontName="Courier",
        fontSize=6.5,
        leading=8,
        textColor=colors.HexColor("#1f2937")
    )
    log_rows = []
    for line in logger.logs:
        safe_line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        log_rows.append([Paragraph(safe_line, log_style)])
        
    log_table = Table(log_rows, colWidths=[757])
    log_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f9fafb")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e5e7eb")),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))

    # Annexure E Table setup (Skipped Rows)
    col_widths_e = [110, 60, 60, 60, 170, 297]
    annex_e_header = [
        Paragraph("<b>Invoice</b>", header_style_small),
        Paragraph("<b>Page</b>", header_style_small),
        Paragraph("<b>Table</b>", header_style_small),
        Paragraph("<b>Row</b>", header_style_small),
        Paragraph("<b>Reason</b>", header_style_small),
        Paragraph("<b>Raw Data</b>", header_style_small),
    ]
    annex_e_rows = []
    if not skipped_rows:
        annex_e_rows.append([
            Paragraph("No skipped rows found.", val_style_small),
            Paragraph("-", val_style_small),
            Paragraph("-", val_style_small),
            Paragraph("-", val_style_small),
            Paragraph("-", val_style_small),
            Paragraph("-", val_style_small),
        ])
    else:
        for r in skipped_rows:
            raw_clean = str(r["Raw Data"]).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            annex_e_rows.append([
                Paragraph(r["Invoice"], val_style_small),
                Paragraph(str(r["Page"]), val_style_small),
                Paragraph(str(r["Table"]), val_style_small),
                Paragraph(str(r["Row"]), val_style_small),
                Paragraph(r["Reason"], val_style_small),
                Paragraph(f"<font face='Courier' size='6'>{raw_clean}</font>", val_style_small),
            ])
    annex_e_table_data = [annex_e_header] + annex_e_rows
    annex_e_table = Table(annex_e_table_data, colWidths=col_widths_e, repeatRows=1)
    annex_e_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d1d5db")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f9fafb")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    # Annexure F Table setup (Warnings)
    warning_invoices = []
    for val in validations:
        status = get_validation_status(val)
        if val.confidence_score < 95 or len(val.warnings) > 0 or status != "PASS":
            warning_invoices.append(val)
            
    col_widths_f = [150, 100, 507]
    annex_f_header = [
        Paragraph("<b>Invoice</b>", header_style_small),
        Paragraph("<b>Confidence</b>", header_style_small),
        Paragraph("<b>Reason for Alert / Action Required</b>", header_style_small),
    ]
    annex_f_rows = []
    if not warning_invoices:
        annex_f_rows.append([
            Paragraph("No warnings found. All invoices are clean.", val_style_small),
            Paragraph("-", val_style_small),
            Paragraph("All checks passed.", val_style_small),
        ])
    else:
        for val in warning_invoices:
            reasons = list(val.warnings)
            total_skipped = val.rows_skipped + val.total_rows_skipped + val.cs_parse_failures
            if total_skipped > 0 and not any("skipped" in r.lower() for r in reasons):
                reasons.append(f"{total_skipped} row(s) skipped.")
            if val.header_row is not None and val.header_row > 0 and not any("fallback" in r.lower() for r in reasons):
                reasons.append("Header detected using fallback.")
                
            reasons_html = "<br/>".join([f"• {r}" for r in reasons]) + "<br/>Please verify manually."
            r_color = "#10b981" if val.risk_level == "LOW" else "#f59e0b" if val.risk_level == "MEDIUM" else "#ef4444"
            confidence_html = f"{val.confidence_score:.1f}% (<font color='{r_color}'><b>{val.risk_level}</b></font>)"
            
            annex_f_rows.append([
                Paragraph(val.invoice_number, val_style_small),
                Paragraph(confidence_html, val_style_small),
                Paragraph(reasons_html, ParagraphStyle("WarnText", parent=styles["Normal"], fontSize=8, leading=10)),
            ])
    annex_f_table_data = [annex_f_header] + annex_f_rows
    annex_f_table = Table(annex_f_table_data, colWidths=col_widths_f, repeatRows=1)
    annex_f_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d1d5db")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f9fafb")]),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))

    # ── Story Construction (Products Table FIRST) ────────────────────────────

    # 1. Products Table (Page 1)
    story.append(TOCRegister("Products"))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("Products Table", heading_style))
    story.append(prod_table)

    if detailed_pdf:
        # 2. Cover Page (Page 2)
        story.append(PageBreak())
        story.append(TOCRegister("Cover"))
        story.append(Spacer(1, 15*mm))
        story.append(Paragraph("CS Extraction Report", title_style))
        story.append(line_table)
        story.append(Spacer(1, 10*mm))
        story.append(cover_table)

        # 3. Executive Summary Page (Page 3)
        story.append(PageBreak())
        story.append(TOCRegister("Summary"))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("Executive Summary", heading_style))
        story.append(summary_table)

        # 4. Table of Contents Page (Page 4)
        story.append(PageBreak())
        story.append(TOCRegister("Table of Contents"))
        story.append(PageBreak())

        # 5. Annexure A: Extraction Confidence Report (Page 5)
        story.append(TOCRegister("Annexure A"))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("Annexure A: Extraction Confidence Report", heading_style))
        story.append(annex_a_table)

        # 6. Annexure B: Validation Summary (Page 6)
        story.append(PageBreak())
        story.append(TOCRegister("Annexure B"))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("Annexure B: Validation Summary", heading_style))
        story.append(annex_b_table)

        # 7. Annexure C: Processing Statistics (Page 7)
        story.append(PageBreak())
        story.append(TOCRegister("Annexure C"))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("Annexure C: Processing Summary Statistics", heading_style))
        story.append(stats_table)

        # 8. Annexure D: Detailed Processing Log (Page 8)
        story.append(PageBreak())
        story.append(TOCRegister("Annexure D"))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("Annexure D: Detailed Processing Log", heading_style))
        story.append(log_table)

        # 9. Annexure E: Skipped Rows Details (Page 9+)
        story.append(PageBreak())
        story.append(TOCRegister("Annexure E"))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("Annexure E: Skipped Rows Details", heading_style))
        story.append(annex_e_table)

        # 10. Annexure F: Warnings Summary
        story.append(PageBreak())
        story.append(TOCRegister("Annexure F"))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("Annexure F: Alerts and Warnings", heading_style))
        story.append(annex_f_table)

    # Build PDF using custom canvasmaker closure
    doc = SimpleDocTemplate(
        str(outfile),
        pagesize=landscape(A4),
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=10*mm, bottomMargin=15*mm,
    )
    
    def make_canvas(*args, **kwargs):
        canvas_obj = NumberedCanvas(*args, **kwargs)
        canvas_obj.detailed_pdf = detailed_pdf
        return canvas_obj
        
    doc.build(story, canvasmaker=make_canvas)

# ── GUI Palette ───────────────────────────────────────────────────────────────

BG      = "#0f0f13"
SURFACE = "#18181f"
CARD    = "#1e1e28"
BORDER  = "#2a2a38"
ACCENT  = "#6c63ff"
ACCENT2 = "#5a52d5"
TEXT    = "#e8e8f0"
MUTED   = "#7a7a94"
SUCCESS = "#34d399"
WARN    = "#fbbf24"
ERR     = "#f87171"

# ── GUI Section ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CS Extractor")
        self.resizable(True, True)
        self.minsize(750, 750)
        self.configure(bg=BG)
        self._files: List[str] = []
        self._fmt_var = tk.StringVar(value="xlsx")
        self._is_running = False
        self._latest_validations: List[ValidationResult] = []
        
        self._build_ui()
        self.update_idletasks()
        self.geometry("820x840")

    def _build_ui(self) -> None:
        self._build_header()
        self._build_stats_panel()
        self._build_dropzone()
        self._build_filelist()
        self._build_progress()
        self._build_format_selector()
        self._build_actions()
        self._build_log()

    def _build_header(self) -> None:
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=20, pady=(20, 0))
        tk.Label(hdr, text="CS Extractor v0.2", font=("Segoe UI", 16, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        self._badge = tk.Label(hdr, text="0 files", font=("Segoe UI", 10),
                               bg=CARD, fg=MUTED, padx=10, pady=3, relief="flat")
        self._badge.pack(side="right", pady=4)

    def _build_stats_panel(self) -> None:
        """Constructs the 5x2 grid of statistics cards."""
        self._stats_frame = tk.Frame(self, bg=BG)
        self._stats_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        self._stat_files = tk.StringVar(value="0")
        self._stat_invoices = tk.StringVar(value="0")
        self._stat_failed_invoices = tk.StringVar(value="0")
        self._stat_products = tk.StringVar(value="0")
        self._stat_unique = tk.StringVar(value="0")
        self._stat_current = tk.StringVar(value="-")
        self._stat_avg_conf = tk.StringVar(value="0%")
        self._stat_warnings = tk.StringVar(value="0")
        self._stat_elapsed = tk.StringVar(value="0.0s")
        
        stats = [
            ("Files", self._stat_files, 0, 0),
            ("Invoices", self._stat_invoices, 0, 1),
            ("Failed Invoices", self._stat_failed_invoices, 0, 2),
            ("Products", self._stat_products, 0, 3),
            ("Unique Products", self._stat_unique, 0, 4),
            ("Current Invoice", self._stat_current, 1, 0),
            ("Average Confidence", self._stat_avg_conf, 1, 2),
            ("Warnings", self._stat_warnings, 1, 3),
            ("Elapsed Time", self._stat_elapsed, 1, 4),
        ]
        
        for text, var, row, col in stats:
            card = tk.Frame(self._stats_frame, bg=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0)
            cspan = 2 if text == "Current Invoice" else 1
            card.grid(row=row, column=col, columnspan=cspan, padx=4, pady=4, sticky="nsew")
            
            lbl = tk.Label(card, text=text, font=("Segoe UI", 9), bg=CARD, fg=MUTED)
            lbl.pack(pady=(6, 2))
            
            val = tk.Label(card, textvariable=var, font=("Segoe UI", 12, "bold"), bg=CARD, fg=TEXT)
            val.pack(pady=(0, 6))
            
            if text == "Warnings":
                self._stat_warnings_label = val
            elif text == "Average Confidence":
                self._stat_avg_conf_label = val
            elif text == "Failed Invoices":
                self._stat_failed_invoices_label = val
                
            # Connect card glow hover events
            def bind_glow(w, c=card):
                w.bind("<Enter>", lambda e: c.configure(highlightbackground=ACCENT))
                w.bind("<Leave>", lambda e: c.configure(highlightbackground=BORDER))
            bind_glow(card)
            bind_glow(lbl)
            bind_glow(val)
                
        for i in range(5):
            self._stats_frame.grid_columnconfigure(i, weight=1)

    def _build_dropzone(self) -> None:
        outer = tk.Frame(self, bg=BORDER)
        outer.pack(fill="x", padx=20, pady=(12, 0))
        self._drop = tk.Frame(outer, bg=CARD)
        self._drop.pack(fill="both", padx=1, pady=1)
        inner = tk.Frame(self._drop, bg=CARD, pady=12)
        inner.pack(fill="x")
        self._icon_lbl  = tk.Label(inner, text="⬆", font=("Segoe UI", 20), bg=CARD, fg=ACCENT)
        self._icon_lbl.pack()
        self._drop_title = tk.Label(inner, text="Choose PDF files",
                                    font=("Segoe UI", 12, "bold"), bg=CARD, fg=TEXT)
        self._drop_title.pack(pady=(6, 2))
        self._drop_sub   = tk.Label(inner, text="Click anywhere in this area to browse",
                                    font=("Segoe UI", 9), bg=CARD, fg=MUTED)
        self._drop_sub.pack()
        
        for w in (self._drop, inner, self._icon_lbl, self._drop_title, self._drop_sub):
            w.bind("<Button-1>", lambda e: self._pick_files())
            w.bind("<Enter>",    lambda e: self._hover(True))
            w.bind("<Leave>",    lambda e: self._hover(False))

    def _build_filelist(self) -> None:
        self._list_frame = tk.Frame(self, bg=BG)
        self._list_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        self._scroll = ttk.Scrollbar(self._list_frame, orient="vertical")
        self._canvas = tk.Canvas(self._list_frame, bg=BG, bd=0,
                                 highlightthickness=0, height=0,
                                 yscrollcommand=self._scroll.set)
        self._scroll.configure(command=self._canvas.yview)
        
        self._scroll.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        
        self._rows_frame = tk.Frame(self._canvas, bg=BG)
        self._canvas_win = self._canvas.create_window((0, 0), window=self._rows_frame, anchor="nw")
        self._rows_frame.bind("<Configure>", self._on_rows_resize)
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(self._canvas_win, width=e.width))

    def _build_progress(self) -> None:
        prog_frame = tk.Frame(self, bg=BG)
        prog_frame.pack(fill="x", padx=20, pady=(10, 0))
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Thin.Horizontal.TProgressbar",
                        troughcolor=CARD, background=ACCENT,
                        thickness=4, borderwidth=0, relief="flat")
        self._progress_var = tk.DoubleVar(value=0.0)
        self._progress = ttk.Progressbar(
            prog_frame, style="Thin.Horizontal.TProgressbar",
            mode="determinate", variable=self._progress_var, maximum=100)
        self._progress.pack(fill="x")
        self._status_var = tk.StringVar(value="")
        tk.Label(prog_frame, textvariable=self._status_var,
                 font=("Segoe UI", 9), bg=BG, fg=MUTED, anchor="w"
                 ).pack(fill="x", pady=(4, 0))

    def _build_format_selector(self) -> None:
        fmt_frame = tk.Frame(self, bg=BG)
        fmt_frame.pack(fill="x", padx=20, pady=(12, 0))
        tk.Label(fmt_frame, text="Output format", font=("Segoe UI", 9),
                 bg=BG, fg=MUTED).pack(side="left", padx=(0, 12))
        
        self._format_rbs = []
        for label, value in [("Excel (.xlsx)", "xlsx"), ("PDF Report", "pdf")]:
            rb = tk.Radiobutton(
                fmt_frame, text=label, variable=self._fmt_var, value=value,
                font=("Segoe UI", 9, "bold"), bg=CARD, fg=TEXT,
                activebackground=BORDER, activeforeground=TEXT,
                selectcolor=CARD, relief="flat", cursor="hand2",
                indicatoron=0, padx=16, pady=6, bd=0)
            rb.pack(side="left", padx=(0, 6))
            self._format_rbs.append((rb, value))
            
            # Hover bindings
            rb.bind("<Enter>", lambda e, r=rb: r.configure(bg=ACCENT if self._fmt_var.get() == r.cget("value") else BORDER))
            rb.bind("<Leave>", lambda e, r=rb: r.configure(bg=ACCENT if self._fmt_var.get() == r.cget("value") else CARD))

        # Checkbox variable - defaults to False (concise mode)
        self._detailed_pdf_var = tk.BooleanVar(value=False)
        
        # Checkbox widget
        self._detailed_pdf_cb = tk.Checkbutton(
            fmt_frame, text="Include Cover Page, Summary & Detailed Verification",
            variable=self._detailed_pdf_var, font=("Segoe UI", 9),
            bg=BG, fg=MUTED, selectcolor=BG, activebackground=BG, activeforeground=ACCENT,
            cursor="hand2", relief="flat", bd=0, highlightthickness=0, state="disabled"
        )
        self._detailed_pdf_cb.pack(side="left", padx=(20, 0))

        def on_format_change(*args):
            current = self._fmt_var.get()
            # Segmented tab style update
            for rb, val in self._format_rbs:
                if val == current:
                    rb.configure(bg=ACCENT, fg="white", activebackground=ACCENT2, activeforeground="white")
                else:
                    rb.configure(bg=CARD, fg=TEXT, activebackground=BORDER, activeforeground=TEXT)
                    
            if current == "pdf":
                self._detailed_pdf_cb.configure(state="normal", fg=TEXT)
            else:
                self._detailed_pdf_cb.configure(state="disabled", fg=MUTED)
                self._detailed_pdf_var.set(False)

        self._fmt_var.trace_add("write", on_format_change)
        # Set initial state
        on_format_change()

    def _build_actions(self) -> None:
        row = tk.Frame(self, bg=BG)
        row.pack(fill="x", padx=20, pady=(8, 0))
        self._run_btn = tk.Button(
            row, text="Extract CS data",
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white", activebackground=ACCENT2,
            activeforeground="white", relief="flat",
            cursor="hand2", height=2, padx=20,
            command=self._run)
        self._run_btn.pack(side="left", fill="x", expand=True)
        self._run_btn.bind("<Enter>", lambda e: self._run_btn.configure(bg=ACCENT2))
        self._run_btn.bind("<Leave>", lambda e: self._run_btn.configure(bg=ACCENT))
        
        tk.Frame(row, width=8, bg=BG).pack(side="left")
        
        clear_btn = tk.Button(row, text="Clear",
                  font=("Segoe UI", 11),
                  bg=CARD, fg=MUTED, activebackground=BORDER,
                  activeforeground=TEXT, relief="flat",
                  cursor="hand2", height=2, padx=16,
                  command=self._clear)
        clear_btn.pack(side="left")
        clear_btn.bind("<Enter>", lambda e: clear_btn.configure(bg=BORDER, fg=TEXT))
        clear_btn.bind("<Leave>", lambda e: clear_btn.configure(bg=CARD, fg=MUTED))

    def _build_log(self) -> None:
        log_frame = tk.Frame(self, bg=SURFACE)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(14, 20))
        
        # Header subframe to hold the label on the left, and copy buttons on the right
        header_sub = tk.Frame(log_frame, bg=SURFACE)
        header_sub.pack(fill="x", padx=12, pady=4)
        
        tk.Label(header_sub, text="Detailed Log", font=("Segoe UI", 9, "bold"),
                 bg=SURFACE, fg=MUTED, anchor="w").pack(side="left")
                 
        # Frame for copy buttons on the right
        btn_frame = tk.Frame(header_sub, bg=SURFACE)
        btn_frame.pack(side="right")
        
        # Helper to style the copy buttons
        def make_copy_btn(text, cmd):
            btn = tk.Button(
                btn_frame, text=text, font=("Segoe UI", 8, "bold"),
                bg=CARD, fg=TEXT, activebackground=BORDER, activeforeground=TEXT,
                relief="flat", cursor="hand2", bd=1, highlightthickness=0,
                padx=6, pady=2
            )
            btn.pack(side="left", padx=4)
            # Bind hover
            btn.bind("<Enter>", lambda e, b=btn: b.configure(fg=ACCENT))
            btn.bind("<Leave>", lambda e, b=btn: b.configure(fg=TEXT))
            btn.configure(command=lambda: cmd(btn, text))
            return btn
            
        make_copy_btn("Copy Logs", self._copy_logs_click)
        make_copy_btn("Copy Warnings", self._copy_warnings_click)
        make_copy_btn("Copy Errors", self._copy_errors_click)
        make_copy_btn("Copy Conf. Reasons", self._copy_conf_click)
        
        tk.Frame(log_frame, bg=BORDER, height=1).pack(fill="x")
        self._log = tk.Text(
            log_frame, height=14,
            font=("Cascadia Code", 9) if self._font_exists("Cascadia Code")
                 else ("Courier New", 9),
            bg=SURFACE, fg=TEXT, insertbackground=TEXT,
            borderwidth=0, highlightthickness=0,
            padx=12, pady=10, state="disabled", wrap="word")
        scroll = tk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)
        self._log.tag_config("ok",   foreground=SUCCESS)
        self._log.tag_config("warn", foreground=WARN)
        self._log.tag_config("err",  foreground=ERR)
        self._log.tag_config("mute", foreground=MUTED)

    def _copy_to_clipboard(self, text: str, button: tk.Button, original_text: str) -> None:
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        button.configure(text="Copied!", fg=SUCCESS)
        self.after(1000, lambda: button.configure(text=original_text, fg=TEXT))

    def _copy_logs_click(self, button: tk.Button, original_text: str) -> None:
        text = "\n".join(logger.logs)
        self._copy_to_clipboard(text or "No logs available.", button, original_text)

    def _copy_warnings_click(self, button: tk.Button, original_text: str) -> None:
        warnings = []
        for val in getattr(self, "_latest_validations", []):
            for w in val.warnings:
                warnings.append(f"Invoice {val.invoice_number}: {w}")
        text = "\n".join(warnings) if warnings else "No warnings found."
        self._copy_to_clipboard(text, button, original_text)

    def _copy_errors_click(self, button: tk.Button, original_text: str) -> None:
        errors = []
        for log in logger.logs:
            if "[error]" in log.lower() or "fail" in log.lower() or "err" in log.lower() or "error" in log.lower():
                if "warning" not in log.lower() and "warn" not in log.lower():
                    errors.append(log)
        for val in getattr(self, "_latest_validations", []):
            if val.cs_parse_failures > 0:
                errors.append(f"Invoice {val.invoice_number}: {val.cs_parse_failures} CS parse failures encountered.")
        text = "\n".join(errors) if errors else "No errors found."
        self._copy_to_clipboard(text, button, original_text)

    def _copy_conf_click(self, button: tk.Button, original_text: str) -> None:
        reasons = []
        for val in getattr(self, "_latest_validations", []):
            val_reasons = []
            if not val.invoice_number:
                val_reasons.append("Missing Invoice Number (-10 pts)")
            if val.tables_found == 0:
                val_reasons.append("No tables found (-10 pts)")
            if val.header_row is None:
                val_reasons.append("No header row found (-10 pts)")
            if not val.product_column or val.product_column == "not found":
                val_reasons.append("Product column not found (-10 pts)")
            if not val.cs_column or val.cs_column == "not found":
                val_reasons.append("CS column not found (-15 pts)")
            if not val.upc_column or val.upc_column == "not found":
                val_reasons.append("UPC column not found (-5 pts)")
            if not val.mrp_column or val.mrp_column == "not found":
                val_reasons.append("MRP column not found (-5 pts)")
            if val.rows_accepted == 0:
                val_reasons.append("No rows accepted (-15 pts)")
            if val.invoice_cs_total == 0 or abs(val.difference) >= 0.01:
                val_reasons.append(f"Totals mismatch: stated={val.invoice_cs_total}, extracted={val.extracted_cs_total} (-15 pts)")
            if val.cs_parse_failures > 0:
                val_reasons.append(f"{val.cs_parse_failures} CS parse failures (-5 pts)")
                
            if val_reasons:
                reasons.append(f"Invoice {val.invoice_number} (Confidence: {val.confidence_score}%, Risk: {val.risk_level}):\n" + "\n".join([f"  • {r}" for r in val_reasons]))
        
        text = "\n\n".join(reasons) if reasons else "No confidence reductions found (All invoices have 100% confidence)."
        self._copy_to_clipboard(text, button, original_text)

    @staticmethod
    def _font_exists(name: str) -> bool:
        try:
            tkfont.Font(family=name)
            return True
        except Exception:
            return False

    def _hover(self, on: bool) -> None:
        bg_col = "#20202a" if on else CARD
        border_col = ACCENT if on else BORDER
        
        self._drop.configure(bg=bg_col)
        self._icon_lbl.master.configure(bg=bg_col)
        self._icon_lbl.configure(bg=bg_col)
        self._drop_title.configure(bg=bg_col)
        self._drop_sub.configure(bg=bg_col)
        self._drop.master.configure(bg=border_col)

    def _pick_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select PDF files", filetypes=[("PDF files", "*.pdf")])
        for p in paths:
            if p not in self._files:
                self._files.append(p)
        self._refresh_list()

    def _clear(self) -> None:
        self._files.clear()
        self._refresh_list()
        self._progress_var.set(0.0)
        self._status_var.set("")
        
        self._stat_files.set("0")
        self._stat_invoices.set("0")
        self._stat_failed_invoices.set("0")
        self._stat_products.set("0")
        self._stat_unique.set("0")
        self._stat_current.set("-")
        self._stat_avg_conf.set("0%")
        self._stat_warnings.set("0")
        self._stat_elapsed.set("0.0s")
        self._stat_avg_conf_label.configure(fg=TEXT)
        self._stat_warnings_label.configure(fg=TEXT)
        self._stat_failed_invoices_label.configure(fg=TEXT)
        
        self._latest_validations = []
        
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _remove_file(self, path: str) -> None:
        self._files.remove(path)
        self._refresh_list()

    def _refresh_list(self) -> None:
        for w in self._rows_frame.winfo_children():
            w.destroy()
        for path in self._files:
            name = Path(path).name
            row  = tk.Frame(self._rows_frame, bg=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0)
            row.pack(fill="x", pady=(0, 4))
            
            # Subframe for padding
            inner_row = tk.Frame(row, bg=CARD, pady=5, padx=8)
            inner_row.pack(fill="x")
            
            tk.Label(inner_row, text="📄", font=("Segoe UI Emoji", 11),
                     bg=CARD, fg=ACCENT).pack(side="left")
            tk.Label(inner_row, text=name, font=("Segoe UI", 10),
                     bg=CARD, fg=TEXT, anchor="w").pack(side="left", padx=8,
                                                        fill="x", expand=True)
            
            del_btn = tk.Button(
                inner_row, text="✕", font=("Segoe UI", 9, "bold"),
                bg=CARD, fg=MUTED, activebackground="#2a1e28",
                activeforeground=ERR, relief="flat", cursor="hand2", bd=0, padx=6, pady=2,
                command=lambda p=path: self._remove_file(p)
            )
            del_btn.pack(side="right")
            
            def bind_del_hover(btn=del_btn):
                btn.bind("<Enter>", lambda e: btn.configure(fg=ERR, bg="#2a1e28"))
                btn.bind("<Leave>", lambda e: btn.configure(fg=MUTED, bg=CARD))
            bind_del_hover()
            
        n = len(self._files)
        self._rows_frame.update_idletasks()
        self._canvas.configure(height=min(self._rows_frame.winfo_reqheight(), 160))
        self._badge.configure(
            text=f"{n} file{'s' if n != 1 else ''}",
            fg=ACCENT if n else MUTED)
        if n:
            self._drop_title.configure(text=f"{n} file{'s' if n != 1 else ''} selected")
            self._drop_sub.configure(text="Click to add more")
        else:
            self._drop_title.configure(text="Choose PDF files")
            self._drop_sub.configure(text="Click anywhere in this area to browse")

    def _on_rows_resize(self, _event=None) -> None:
        self._canvas.configure(height=min(self._rows_frame.winfo_reqheight(), 160))

    def _log_write(self, msg: str) -> None:
        lower = msg.lower()
        if "pass" in lower or "✓" in lower:
            tag = "ok"
        elif "fail" in lower or "err" in lower or "error" in lower:
            tag = "err"
        elif "warning" in lower or "warn" in lower or "skipped" in lower or "confidence" in lower:
            tag = "warn"
        elif "opening" in lower or "loading" in lower:
            tag = "mute"
        else:
            tag = ""

        self._log.config(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _on_progress(self, pct: float, status: str) -> None:
        self.after(0, lambda: self._progress_var.set(pct))
        self.after(0, lambda: self._status_var.set(status))

    def _gui_update(
        self,
        files_count=0,
        invoices_count=0,
        products_count=0,
        unique_products_count=0,
        current_invoice="-",
        avg_confidence=0.0,
        warnings_count=0,
        elapsed=0.0,
        validations=None
    ) -> None:
        self._stat_files.set(str(files_count))
        self._stat_invoices.set(str(invoices_count))
        self._stat_products.set(str(products_count))
        self._stat_unique.set(str(unique_products_count))
        self._stat_current.set(str(current_invoice))
        self._stat_avg_conf.set(f"{avg_confidence:.1f}%")
        self._stat_warnings.set(str(warnings_count))
        self._stat_elapsed.set(f"{elapsed:.1f}s")
        
        if validations is not None:
            self._latest_validations = validations
            
        failed_count = 0
        if hasattr(self, "_latest_validations") and self._latest_validations:
            failed_count = sum(1 for v in self._latest_validations if get_validation_status(v) == "FAIL")
            
        self._stat_failed_invoices.set(str(failed_count))
        
        # Color coding avg confidence
        if avg_confidence >= 96:
            self._stat_avg_conf_label.configure(fg=SUCCESS)
        elif avg_confidence >= 90:
            self._stat_avg_conf_label.configure(fg=WARN)
        else:
            self._stat_avg_conf_label.configure(fg=ERR)
            
        # Color coding warnings
        if warnings_count > 0:
            self._stat_warnings_label.configure(fg=WARN)
        else:
            self._stat_warnings_label.configure(fg=TEXT)

        # Color coding failed invoices
        if failed_count > 0:
            self._stat_failed_invoices_label.configure(fg=ERR)
        else:
            self._stat_failed_invoices_label.configure(fg=TEXT)

    def _update_timer(self, start_time: float) -> None:
        if self._is_running:
            elapsed = time.time() - start_time
            self._stat_elapsed.set(f"{elapsed:.1f}s")
            self.after(100, lambda: self._update_timer(start_time))

    def _run(self) -> None:
        if not self._files:
            messagebox.showwarning("No files", "Add at least one PDF first.")
            return
        
        self._is_running = True
        self._run_btn.configure(state="disabled", text="Extracting…")
        self._progress_var.set(0.0)
        self._status_var.set("Starting…")
        
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        
        pdfs = [Path(f) for f in self._files]
        logger.gui_callback = self._log_write

        def worker():
            start_t = time.time()
            self.after(0, lambda: self._update_timer(start_t))
            try:
                def update_cb(**kwargs):
                    self.after(0, lambda: self._gui_update(**kwargs))

                outfile = process(
                    pdfs,
                    on_progress=self._on_progress,
                    fmt=self._fmt_var.get(),
                    gui_update_callback=update_cb,
                    detailed_pdf=self._detailed_pdf_var.get()
                )
                if outfile:
                    self.after(0, lambda: self._status_var.set(f"Done  ·  {outfile.name}"))
                    self.after(0, lambda: messagebox.showinfo(
                        "Extraction complete", f"Saved to:\n{outfile.resolve()}"))
                else:
                    self.after(0, lambda: self._status_var.set("No CS rows found"))
                    self.after(0, lambda: messagebox.showwarning(
                        "Nothing found", "No CS rows were found in the selected files."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
                self.after(0, lambda: self._status_var.set(f"Error: {e}"))
            finally:
                self.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self) -> None:
        self._is_running = False
        self._progress_var.set(100.0)
        self._run_btn.configure(state="normal", text="Extract CS data")

# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Connect logger to stdout printing for CLI runs
        process([Path(p) for p in sys.argv[1:]])
    else:
        App().mainloop()
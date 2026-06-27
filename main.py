import sys, threading, re
from pathlib import Path

VERSION = "0.1.0"
__version__ = VERSION

import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

CS_NAMES      = {"cs", "c/s", "case", "cases", "case qty", "case quantity"}
SL_NAMES      = {"sl", "sl.", "s.no", "sno", "sr", "sr.", "no", "no.", "#", "item"}
TOTAL_KEYWORDS = {"total", "sub total", "subtotal", "grand total"}

INVOICE_PATTERNS = [
    re.compile(r"Bill\s*(?:No|NO|Number)?\s*:?\s*([A-Z0-9/-]+)", re.I),
    re.compile(r"Invoice\s*(?:No|NO|Number)?\s*:?\s*([A-Z0-9/-]+)", re.I),
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def safe_str(val) -> str:
    """Convert any cell value (including None) to a clean string."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "none" else s

def extract_invoice_number(text: str) -> str:
    for p in INVOICE_PATTERNS:
        m = p.search(text or "")
        if m:
            return m.group(1).strip()
    return ""

def find_cs(headers: list[str]) -> int | None:
    for i, h in enumerate(headers):
        n = h.lower().strip().replace(".", "").replace("_", " ")
        if n in CS_NAMES:
            return i
    return None

def find_sl(headers: list[str]) -> int | None:
    for i, h in enumerate(headers):
        n = h.lower().strip().replace(".", "").replace("_", " ")
        if n in SL_NAMES:
            return i
    return None

def is_total_row(vals: list[str], sl_idx: int | None) -> bool:
    if sl_idx is not None and not vals[sl_idx].strip():
        return True
    for v in vals:
        if v.lower().strip() in TOTAL_KEYWORDS:
            return True
    return False

def parse_header_row(row: list) -> list[str]:
    """Build headers from a raw table row, safely handling None and duplicates."""
    headers = []
    seen: dict[str, int] = {}
    for i, cell in enumerate(row):
        h = safe_str(cell) or f"Column_{i+1}"
        # deduplicate: Column_1, Column_1_2, Column_1_3 …
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 1
        headers.append(h)
    return headers

def parse_data_row(row: list, n_cols: int) -> list[str]:
    """Convert a data row to strings, pad/trim to n_cols."""
    vals = [safe_str(c) for c in row]
    while len(vals) < n_cols:
        vals.append("")
    return vals[:n_cols]

def find_header_row(table: list) -> int | None:
    """
    Scan rows top-down and return the index of the first row that contains
    a CS column name. The invoice header / address block often occupies
    row 0, with the real column headers on row 1 or 2.
    """
    for i, row in enumerate(table):
        headers = parse_header_row(row)
        if find_cs(headers) is not None:
            return i
    return None

# ── Core logic ────────────────────────────────────────────────────────────────

def extract_from_pdf(pdf_path: Path, log, on_page=None) -> list[dict]:
    rows_out: list[dict] = []
    current_invoice = ""

    log(f"  ┌ Opening: {pdf_path.name}")
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        log(f"  │ Pages: {total_pages}")

        for page_num, page in enumerate(pdf.pages, 1):
            if on_page:
                on_page(page_num, total_pages)

            page_text = page.extract_text() or ""
            inv = extract_invoice_number(page_text)
            if inv and inv != current_invoice:
                current_invoice = inv
                log(f"  │ [p{page_num}] Invoice detected: {current_invoice}")

            tables = page.extract_tables() or []
            if not tables:
                log(f"  │ [p{page_num}] No tables found")
                continue

            log(f"  │ [p{page_num}] {len(tables)} table(s) found")

            for t_idx, table in enumerate(tables, 1):
                prefix = f"  │ [p{page_num}/t{t_idx}]"

                if not table or len(table) < 2:
                    log(f"{prefix} Skipped — too few rows ({len(table) if table else 0})")
                    continue

                # Scan all rows to find the real header (row 0 may be address block)
                header_row_idx = find_header_row(table)
                if header_row_idx is None:
                    log(f"{prefix} Skipped — no CS column found in any row")
                    continue

                if header_row_idx > 0:
                    log(f"{prefix} Real header at row {header_row_idx} "
                        f"(skipped {header_row_idx} preamble row(s))")

                headers  = parse_header_row(table[header_row_idx])
                n_cols   = len(headers)
                cs_idx   = find_cs(headers)
                sl_idx   = find_sl(headers)

                prod_idx = next((i for i, h in enumerate(headers)
                                 if "product" in h.lower() and "name" in h.lower()), None)
                upc_idx  = next((i for i, h in enumerate(headers)
                                 if h.lower().strip().replace(".", "")
                                 in {"upc", "upc code", "barcode"}), None)
                mrp_idx  = next((i for i, h in enumerate(headers)
                                 if h.lower().strip().replace(".", "").replace(" ", "")
                                 in {"mrp", "mrp rs", "mrprs"}), None)

                log(f"{prefix} CS col={headers[cs_idx]!r}  "
                    f"SL={'col '+str(sl_idx) if sl_idx is not None else 'not found'}  "
                    f"Product={'col '+str(prod_idx) if prod_idx is not None else 'not found'}  "
                    f"UPC={'col '+str(upc_idx) if upc_idx is not None else 'not found'}  "
                    f"MRP={'col '+str(mrp_idx) if mrp_idx is not None else 'not found'}")

                n_cols    = len(headers)
                data_rows = table[header_row_idx + 1:]  # rows after the real header
                accepted   = 0
                skipped_total = 0
                skipped_cs    = 0

                for row_num, raw_row in enumerate(data_rows, 2):
                    if not raw_row or not any(safe_str(c) for c in raw_row):
                        skipped_total += 1
                        continue

                    vals = parse_data_row(raw_row, n_cols)

                    if is_total_row(vals, sl_idx):
                        skipped_total += 1
                        continue

                    raw_cs = vals[cs_idx]
                    try:
                        cs = float(raw_cs.replace(",", ""))
                    except (ValueError, AttributeError):
                        log(f"{prefix}   row {row_num}: CS parse failed on {raw_cs!r} — skipping")
                        skipped_cs += 1
                        continue

                    if cs <= 0:
                        skipped_cs += 1
                        continue

                    product = vals[prod_idx] if prod_idx is not None else ""
                    upc     = vals[upc_idx]  if upc_idx  is not None else ""
                    mrp     = vals[mrp_idx]  if mrp_idx  is not None else ""

                    if not product:
                        log(f"{prefix}   row {row_num}: ⚠ Product Name empty (CS={cs}, UPC={upc})")

                    rows_out.append({
                        "Invoice Number": current_invoice or pdf_path.stem,
                        "Product Name":   product,
                        "UPC":            upc,
                        "MRP":            mrp,
                        "CS":             cs,
                    })
                    accepted += 1

                log(f"{prefix} ✓ {accepted} rows accepted  "
                    f"| {skipped_total} total/empty skipped  "
                    f"| {skipped_cs} zero/bad-CS skipped")

    log(f"  └ Done: {len(rows_out)} CS rows extracted from {pdf_path.name}")
    return rows_out


def process(pdfs, log=print, on_progress=None, save_path=None, fmt="xlsx"):
    total    = len(pdfs)
    all_rows: list[dict] = []

    for idx, pdf in enumerate(pdfs):
        log(f"ℹ Loading {pdf.name}")
        try:
            def on_page(page_num, total_pages, _idx=idx, _total=total):
                if on_progress:
                    pct = ((_idx + page_num / total_pages) / _total) * 100
                    on_progress(
                        pct,
                        f"Processing file {_idx+1}/{_total} | {pdf.name} | Page {page_num}/{total_pages}"
                    )
            all_rows.extend(extract_from_pdf(pdf, log, on_page=on_page))
        except Exception as e:
            log(f"⚠ Error in {pdf.name}: {e}")

    if not all_rows:
        log("No CS rows found.")
        return None

    if on_progress:
        on_progress(100, "Saving PDF…" if fmt == "pdf" else "Saving spreadsheet…")

    df = pd.DataFrame(all_rows, columns=["Invoice Number", "Product Name", "UPC", "MRP", "CS"])
    df["MRP"] = pd.to_numeric(df["MRP"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    df["CS"]  = pd.to_numeric(df["CS"].astype(str).str.replace(",", "", regex=False),  errors="coerce").fillna(0)

    before = len(df)
    df = (
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
    log(f"  → merged {before} rows → {len(df)} unique products")
    df = df[["Invoice Numbers", "Product Name", "UPC", "MRP", "CS"]]

    # ── Save ──────────────────────────────────────────────────────────────────
    if save_path is None:
        stem = pdfs[0].stem if len(pdfs) == 1 else "Combined"
        ext  = ".pdf" if fmt == "pdf" else ".xlsx"
        default_name = f"{stem}_CS_Extract{ext}"
        ftypes = ([("PDF Document", "*.pdf")] if fmt == "pdf"
                  else [("Excel Workbook", "*.xlsx")])
        save_path = filedialog.asksaveasfilename(
            title=f"Save {'PDF Report' if fmt == 'pdf' else 'Excel File'}",
            defaultextension=ext,
            initialfile=default_name,
            filetypes=ftypes,
        )
        if not save_path:
            log("⚠ Save cancelled.")
            return None

    outfile = Path(save_path)
    if fmt == "pdf":
        save_as_pdf(df, outfile, [p.name for p in pdfs])
    else:
        df.to_excel(outfile, index=False)

    log(f"✓ Saved → {outfile.resolve()}")
    return outfile


def save_as_pdf(df, outfile: Path, source_names: list[str]):
    # ── Grayscale palette ─────────────────────────────────────────────────────
    G_HEADER_BG   = colors.Color(0.93, 0.93, 0.93)   # light gray header
    G_HEADER_TXT  = colors.white
    G_ROW_A       = colors.white
    G_ROW_B       = colors.Color(0.93, 0.93, 0.93)   # light gray alternate
    G_GRID        = colors.Color(0.75, 0.75, 0.75)
    G_TITLE       = colors.Color(0.15, 0.15, 0.15)
    G_MUTED       = colors.Color(0.45, 0.45, 0.45)

    doc = SimpleDocTemplate(
        str(outfile),
        pagesize=landscape(A4),
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
    )
    styles    = getSampleStyleSheet()
    title_sty = ParagraphStyle("title", parent=styles["Normal"],
                               fontSize=14, fontName="Helvetica-Bold",
                               textColor=G_TITLE, spaceAfter=4)
    sub_sty   = ParagraphStyle("sub",   parent=styles["Normal"],
                               fontSize=8,  fontName="Helvetica",
                               textColor=G_MUTED, spaceAfter=10)
    cell_sty  = ParagraphStyle("cell",  parent=styles["Normal"],
                               fontSize=8,  fontName="Helvetica", leading=10)

    elements = [
        Paragraph("CS Extractor — Output Report", title_sty),
        Paragraph("Sources: " + ", ".join(source_names), sub_sty),
    ]

    header_row = [Paragraph(f"<b>{c}</b>", cell_sty) for c in df.columns]
    data_rows  = [
        [Paragraph("" if str(v) == "nan" else str(v), cell_sty) for v in row]
        for row in df.itertuples(index=False)
    ]
    table_data = [header_row] + data_rows

    page_w     = landscape(A4)[0] - 30*mm
    col_widths = [page_w * r for r in [0.20, 0.50, 0.10, 0.10, 0.10]]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1,  0), G_HEADER_BG),
        ("TEXTCOLOR",      (0, 0), (-1,  0), G_HEADER_TXT),
        ("FONTNAME",       (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",       (0, 0), (-1,  0), 8),
        ("BOTTOMPADDING",  (0, 0), (-1,  0), 6),
        ("TOPPADDING",     (0, 0), (-1,  0), 6),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [G_ROW_A, G_ROW_B]),
        ("FONTNAME",       (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",       (0, 1), (-1, -1), 8),
        ("TOPPADDING",     (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING",  (0, 1), (-1, -1), 4),
        ("GRID",           (0, 0), (-1, -1), 0.25, G_GRID),
        ("LINEBELOW",      (0, 0), (-1,  0), 1,    G_HEADER_BG),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
    ]))

    elements.append(tbl)
    elements.append(Spacer(1, 6*mm))
    total_cs = df["CS"].sum()
    elements.append(Paragraph(
        f"Total products: {len(df)}  ·  Total CS: {int(total_cs)}",
        ParagraphStyle("footer", parent=styles["Normal"],
                       fontSize=8, fontName="Helvetica",
                       textColor=G_MUTED, alignment=TA_LEFT)
    ))
    doc.build(elements)


# ── Palette ───────────────────────────────────────────────────────────────────

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

# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CS Extractor")
        self.resizable(True, True)
        self.minsize(700, 700)
        self.configure(bg=BG)
        self._files: list[str] = []
        self._fmt_var = tk.StringVar(value="xlsx")
        self._build_ui()
        self.update_idletasks()
        self.geometry("820x760")

    def _build_ui(self):
        self._build_header()
        self._build_dropzone()
        self._build_filelist()
        self._build_progress()
        self._build_format_selector()
        self._build_actions()
        self._build_log()

    def _build_header(self):
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=20, pady=(20, 0))
        tk.Label(hdr, text="CS Extractor", font=("Segoe UI", 16, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        self._badge = tk.Label(hdr, text="0 files", font=("Segoe UI", 10),
                               bg=CARD, fg=MUTED, padx=10, pady=3, relief="flat")
        self._badge.pack(side="right", pady=4)

    def _build_dropzone(self):
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

    def _build_filelist(self):
        self._list_frame = tk.Frame(self, bg=BG)
        self._list_frame.pack(fill="x", padx=20, pady=(10, 0))
        self._canvas = tk.Canvas(self._list_frame, bg=BG, bd=0,
                                 highlightthickness=0, height=0)
        self._canvas.pack(fill="both", expand=False)
        self._rows_frame = tk.Frame(self._canvas, bg=BG)
        self._canvas.create_window((0, 0), window=self._rows_frame, anchor="nw")
        self._rows_frame.bind("<Configure>", self._on_rows_resize)

    def _build_progress(self):
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

    def _build_format_selector(self):
        fmt_frame = tk.Frame(self, bg=BG)
        fmt_frame.pack(fill="x", padx=20, pady=(12, 0))
        tk.Label(fmt_frame, text="Output format", font=("Segoe UI", 9),
                 bg=BG, fg=MUTED).pack(side="left", padx=(0, 12))
        for label, value in [("Excel (.xlsx)", "xlsx"), ("PDF Report", "pdf")]:
            rb = tk.Radiobutton(
                fmt_frame, text=label, variable=self._fmt_var, value=value,
                font=("Segoe UI", 10), bg=BG, fg=TEXT,
                activebackground=BG, activeforeground=ACCENT,
                selectcolor=CARD, relief="flat", cursor="hand2",
                indicatoron=0, padx=14, pady=5, bd=0)
            rb.pack(side="left", padx=(0, 6))
            rb.bind("<Enter>", lambda e, b=rb: b.configure(fg=ACCENT))
            rb.bind("<Leave>", lambda e, b=rb: b.configure(
                fg=ACCENT if self._fmt_var.get() == b.cget("value") else TEXT))

    def _build_actions(self):
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
        tk.Frame(row, width=8, bg=BG).pack(side="left")
        tk.Button(row, text="Clear",
                  font=("Segoe UI", 11),
                  bg=CARD, fg=MUTED, activebackground=BORDER,
                  activeforeground=TEXT, relief="flat",
                  cursor="hand2", height=2, padx=16,
                  command=self._clear).pack(side="left")

    def _build_log(self):
        log_frame = tk.Frame(self, bg=SURFACE)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(14, 20))
        tk.Label(log_frame, text="Log", font=("Segoe UI", 9, "bold"),
                 bg=SURFACE, fg=MUTED, anchor="w", padx=12, pady=6).pack(fill="x")
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

    @staticmethod
    def _font_exists(name: str) -> bool:
        try:
            tk.font.Font(family=name)  # type: ignore
            return True
        except Exception:
            return False

    def _hover(self, on: bool):
        self._drop.configure(bg="#242432" if on else CARD)

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF files", filetypes=[("PDF files", "*.pdf")])
        for p in paths:
            if p not in self._files:
                self._files.append(p)
        self._refresh_list()

    def _clear(self):
        self._files.clear()
        self._refresh_list()
        self._progress_var.set(0.0)
        self._status_var.set("")

    def _remove_file(self, path: str):
        self._files.remove(path)
        self._refresh_list()

    def _refresh_list(self):
        for w in self._rows_frame.winfo_children():
            w.destroy()
        for path in self._files:
            name = Path(path).name
            row  = tk.Frame(self._rows_frame, bg=CARD, pady=6, padx=10)
            row.pack(fill="x", pady=(0, 2))
            tk.Label(row, text="📄", font=("Segoe UI Emoji", 12),
                     bg=CARD, fg=ACCENT).pack(side="left")
            tk.Label(row, text=name, font=("Segoe UI", 10),
                     bg=CARD, fg=TEXT, anchor="w").pack(side="left", padx=8,
                                                        fill="x", expand=True)
            tk.Button(row, text="✕", font=("Segoe UI", 9),
                      bg=CARD, fg=MUTED, activebackground=CARD,
                      activeforeground=ERR, relief="flat", cursor="hand2",
                      command=lambda p=path: self._remove_file(p)).pack(side="right")
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

    def _on_rows_resize(self, _event=None):
        self._canvas.configure(height=min(self._rows_frame.winfo_reqheight(), 160))

    def _log_write(self, msg: str):
        tag = ("ok"   if msg.startswith("✓") or msg.startswith("→") else
               "err"  if msg.startswith("⚠") else
               "mute" if msg.startswith("ℹ") else "")
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _on_progress(self, pct: float, status: str):
        self.after(0, lambda: self._progress_var.set(pct))
        self.after(0, lambda: self._status_var.set(status))

    def _run(self):
        if not self._files:
            messagebox.showwarning("No files", "Add at least one PDF first.")
            return
        self._run_btn.configure(state="disabled", text="Extracting…")
        self._progress_var.set(0.0)
        self._status_var.set("Starting…")
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        pdfs = [Path(f) for f in self._files]

        def worker():
            try:
                outfile = process(pdfs, log=self._log_write,
                                  on_progress=self._on_progress,
                                  fmt=self._fmt_var.get())
                if outfile:
                    self.after(0, lambda: self._status_var.set(f"Done  ·  {outfile}"))
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

    def _done(self):
        self._progress_var.set(100.0)
        self._run_btn.configure(state="normal", text="Extract CS data")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        process([Path(p) for p in sys.argv[1:]])
    else:
        App().mainloop()
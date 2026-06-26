import sys, threading
from pathlib import Path
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

CS_NAMES = {"cs", "c/s", "case", "cases", "case qty", "case quantity"}
SL_NAMES = {"sl", "sl.", "s.no", "sno", "sr", "sr.", "no", "no.", "#", "item"}
TOTAL_KEYWORDS = {"total", "sub total", "subtotal", "grand total"}

# ── Core logic ────────────────────────────────────────────────────────────────

def find_cs(headers):
    for i, h in enumerate(headers):
        n = h.lower().strip().replace(".", "").replace("_", " ")
        if n in CS_NAMES:
            return i
    return None

def find_sl(headers):
    for i, h in enumerate(headers):
        n = h.lower().strip().replace(".", "").replace("_", " ")
        if n in SL_NAMES:
            return i
    return None

def is_total_row(vals: list[str], sl_idx: int | None) -> bool:
    # Total rows have an empty SL column
    if sl_idx is not None and not vals[sl_idx].strip():
        return True
    # Or a cell literally says "total"
    for v in vals:
        if v.lower().strip() in TOTAL_KEYWORDS:
            return True
    return False



def find_invoice_numbers(page):
    """Return all invoice numbers found on a page."""
    txt = page.extract_text() or ""
    pattern = re.compile(
        r'(?i)(?:invoice\s*(?:no|number)?|inv\.?\s*no\.?|bill\s*no\.?|document\s*no\.?)\s*[:#-]?\s*([A-Z0-9\-/]+)'
    )
    nums = []
    for m in pattern.finditer(txt):
        n=m.group(1).strip()
        if n not in nums:
            nums.append(n)
    return nums

def extract_from_pdf(pdf_path: Path, log, on_page=None):
    rows_out = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages, 1):
            if on_page:
                on_page(page_num, total_pages)

            page_invoices = find_invoice_numbers(page)
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 2:
                    continue
                headers = [str(c).strip() if c else f"Column_{i+1}"
                           for i, c in enumerate(table[0])]
                cs_idx = find_cs(headers)
                if cs_idx is None:
                    continue
                sl_idx = find_sl(headers)

                prod_idx = next((i for i, h in enumerate(headers)
                                 if "product" in h.lower() and "name" in h.lower()), None)
                upc_idx  = next((i for i, h in enumerate(headers)
                                 if h.lower().strip().replace(".", "") in {"upc", "upc code", "barcode"}), None)
                mrp_idx  = next((i for i, h in enumerate(headers)
                                 if h.lower().strip().replace(".", "").replace(" ", "") in {"mrp", "mrp rs", "mrprs"}), None)

                for row in table[1:]:
                    if not row:
                        continue
                    vals = [str(c).strip() if c else "" for c in row]
                    vals.extend([""] * (len(headers) - len(vals)))
                    if is_total_row(vals, sl_idx):
                        continue
                    try:
                        cs = float(vals[cs_idx].replace(",", ""))
                    except:
                        cs = 0
                    if cs > 0:
                        rows_out.append({
                            "Source PDF":  pdf_path.name,
                            "Page Number": page_num,
                            "Product Name": vals[prod_idx] if prod_idx is not None else "",
                            "UPC":         vals[upc_idx]  if upc_idx  is not None else "",
                            "MRP":         vals[mrp_idx]  if mrp_idx  is not None else "",
                            "CS":          cs,
                            "Invoices":   ", ".join(page_invoices),
                        })

    log(f"  → {len(rows_out)} rows in {pdf_path.name}")
    return rows_out

def process(pdfs, log=print, on_progress=None, save_path=None):
    total = len(pdfs)
    all_rows = []
    for idx, pdf in enumerate(pdfs):
        log(f"[{idx+1}/{total}] {pdf.name}")
        try:
            def on_page(page_num, total_pages, _idx=idx, _total=total):
                if on_progress:
                    pct = ((_idx + page_num / total_pages) / _total) * 100
                    on_progress(pct, f"{pdf.name}  ·  page {page_num} of {total_pages}")
            all_rows.extend(extract_from_pdf(pdf, log, on_page=on_page))
        except Exception as e:
            log(f"  ⚠  {e}")
    if not all_rows:
        log("No CS rows found.")
        return None
    if on_progress:
        on_progress(100, "Saving spreadsheet…")
    df = pd.DataFrame(all_rows, columns=["Source PDF", "Page Number", "Product Name", "UPC", "MRP", "CS"])

    df["MRP"] = pd.to_numeric(df["MRP"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    df["CS"]  = pd.to_numeric(df["CS"].astype(str).str.replace(",", "", regex=False),  errors="coerce").fillna(0)

    before = len(df)
    df = (
        df.groupby(["Source PDF", "Product Name", "UPC", "MRP"], as_index=False)
          .agg(
              **{
                  "Pages": ("Page Number",
                            lambda s: ", ".join(str(p) for p in sorted(s.unique()))),
                  "Invoices": ("Invoices",
                            lambda s: ", ".join(sorted({i.strip() for v in s for i in v.split(",") if i.strip()}))),
                  "CS": ("CS", "sum"),
              }
          )
    )
    log(f"  → merged {before} rows → {len(df)} unique products")

    df = df[["Source PDF", "Pages", "Invoices", "Product Name", "UPC", "MRP", "CS"]]
    if save_path is None:
        default_name = (
            pdfs[0].stem + "_CS_Extract.xlsx"
            if len(pdfs) == 1
            else "Combined_CS_Extract.xlsx"
        )
        save_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not save_path:
            log("⚠ Save cancelled.")
            return None
    outfile = Path(save_path)
    df.to_excel(outfile, index=False)
    log(f"✓  Saved → {outfile.resolve()}")
    return outfile

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
        self.resizable(False, False)
        self.configure(bg=BG)
        self._files: list[str] = []
        self._build_ui()
        self.update_idletasks()
        self.geometry(f"520x{self.winfo_reqheight()}")

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_dropzone()
        self._build_filelist()
        self._build_progress()
        self._build_actions()
        self._build_log()

    def _build_header(self):
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=20, pady=(20, 0))
        tk.Label(hdr, text="CS Extractor", font=("Segoe UI", 16, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        self._badge = tk.Label(hdr, text="0 files", font=("Segoe UI", 10),
                               bg=CARD, fg=MUTED, padx=10, pady=3)
        self._badge.pack(side="right", pady=4)
        self._badge.configure(relief="flat")

    def _build_dropzone(self):
        outer = tk.Frame(self, bg=BORDER, bd=0)
        outer.pack(fill="x", padx=20, pady=(12, 0))

        self._drop = tk.Frame(outer, bg=CARD, bd=0)
        self._drop.pack(fill="both", padx=1, pady=1)

        inner = tk.Frame(self._drop, bg=CARD, pady=24)
        inner.pack(fill="x")

        self._icon_lbl = tk.Label(inner, text="⬆", font=("Segoe UI", 28),
                                  bg=CARD, fg=ACCENT)
        self._icon_lbl.pack()

        self._drop_title = tk.Label(inner, text="Choose PDF files",
                                    font=("Segoe UI", 12, "bold"),
                                    bg=CARD, fg=TEXT)
        self._drop_title.pack(pady=(6, 2))

        self._drop_sub = tk.Label(inner, text="Click anywhere in this area to browse",
                                  font=("Segoe UI", 9),
                                  bg=CARD, fg=MUTED)
        self._drop_sub.pack()

        for w in (self._drop, inner, self._icon_lbl,
                  self._drop_title, self._drop_sub):
            w.bind("<Button-1>", lambda e: self._pick_files())
            w.bind("<Enter>",    lambda e: self._hover(True))
            w.bind("<Leave>",    lambda e: self._hover(False))

    def _build_filelist(self):
        self._list_frame = tk.Frame(self, bg=BG)
        self._list_frame.pack(fill="x", padx=20, pady=(10, 0))

        # Scrollable canvas for file rows
        self._canvas = tk.Canvas(self._list_frame, bg=BG, bd=0,
                                 highlightthickness=0, height=0)
        self._canvas.pack(fill="x")
        self._rows_frame = tk.Frame(self._canvas, bg=BG)
        self._canvas.create_window((0, 0), window=self._rows_frame, anchor="nw")
        self._rows_frame.bind("<Configure>", self._on_rows_resize)

    def _build_progress(self):
        prog_frame = tk.Frame(self, bg=BG)
        prog_frame.pack(fill="x", padx=20, pady=(14, 0))

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Thin.Horizontal.TProgressbar",
                        troughcolor=CARD,
                        background=ACCENT,
                        thickness=4,
                        borderwidth=0,
                        relief="flat")

        self._progress_var = tk.DoubleVar(value=0.0)
        self._progress = ttk.Progressbar(
            prog_frame, style="Thin.Horizontal.TProgressbar",
            mode="determinate", variable=self._progress_var, maximum=100)
        self._progress.pack(fill="x")

        self._status_var = tk.StringVar(value="")
        self._status_lbl = tk.Label(prog_frame, textvariable=self._status_var,
                                    font=("Segoe UI", 9), bg=BG, fg=MUTED,
                                    anchor="w")
        self._status_lbl.pack(fill="x", pady=(4, 0))

    def _build_actions(self):
        row = tk.Frame(self, bg=BG)
        row.pack(fill="x", padx=20, pady=(12, 0))

        self._run_btn = tk.Button(
            row, text="Extract CS data",
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white", activebackground=ACCENT2,
            activeforeground="white", relief="flat",
            cursor="hand2", height=2, padx=20,
            command=self._run)
        self._run_btn.pack(side="left", fill="x", expand=True)

        tk.Frame(row, width=8, bg=BG).pack(side="left")

        self._clear_btn = tk.Button(
            row, text="Clear",
            font=("Segoe UI", 11),
            bg=CARD, fg=MUTED, activebackground=BORDER,
            activeforeground=TEXT, relief="flat",
            cursor="hand2", height=2, padx=16,
            command=self._clear)
        self._clear_btn.pack(side="left")

    def _build_log(self):
        log_frame = tk.Frame(self, bg=SURFACE, bd=0)
        log_frame.pack(fill="x", padx=20, pady=(14, 20))

        tk.Label(log_frame, text="Log", font=("Segoe UI", 9, "bold"),
                 bg=SURFACE, fg=MUTED, anchor="w",
                 padx=12, pady=6).pack(fill="x")

        sep = tk.Frame(log_frame, bg=BORDER, height=1)
        sep.pack(fill="x")

        self._log = tk.Text(
            log_frame, height=7,
            font=("Cascadia Code", 9) if self._font_exists("Cascadia Code")
                 else ("Courier New", 9),
            bg=SURFACE, fg=TEXT, insertbackground=TEXT,
            borderwidth=0, highlightthickness=0,
            padx=12, pady=10,
            state="disabled", wrap="word")
        self._log.pack(fill="x")

        # Tag colours
        self._log.tag_config("ok",   foreground=SUCCESS)
        self._log.tag_config("warn", foreground=WARN)
        self._log.tag_config("err",  foreground=ERR)
        self._log.tag_config("mute", foreground=MUTED)

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _font_exists(name: str) -> bool:
        try:
            tk.font.Font(family=name)  # type: ignore
            return True
        except Exception:
            return False

    def _hover(self, on: bool):
        col = "#242432" if on else CARD
        for w in (self._drop,):
            w.configure(bg=col)

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf")])
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
            row = tk.Frame(self._rows_frame, bg=CARD, pady=6, padx=10)
            row.pack(fill="x", pady=(0, 2))

            tk.Label(row, text="📄", font=("Segoe UI Emoji", 12),
                     bg=CARD, fg=ACCENT).pack(side="left")

            tk.Label(row, text=name, font=("Segoe UI", 10),
                     bg=CARD, fg=TEXT, anchor="w").pack(side="left", padx=8, fill="x", expand=True)

            rm = tk.Button(row, text="✕", font=("Segoe UI", 9),
                           bg=CARD, fg=MUTED, activebackground=CARD,
                           activeforeground=ERR, relief="flat",
                           cursor="hand2",
                           command=lambda p=path: self._remove_file(p))
            rm.pack(side="right")

        n = len(self._files)
        # resize canvas
        self._rows_frame.update_idletasks()
        h = min(self._rows_frame.winfo_reqheight(), 160)
        self._canvas.configure(height=h)

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
        h = min(self._rows_frame.winfo_reqheight(), 160)
        self._canvas.configure(height=h)

    def _log_write(self, msg: str):
        if msg.startswith("✓") or msg.startswith("→"):
            tag = "ok"
        elif msg.startswith("⚠"):
            tag = "err"
        elif msg.startswith("["):
            tag = "mute"
        else:
            tag = ""
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _on_progress(self, pct: float, status: str):
        self.after(0, lambda: self._progress_var.set(pct))
        self.after(0, lambda: self._status_var.set(status))

    # ── Run ───────────────────────────────────────────────────────────────────

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
                                  on_progress=self._on_progress)
                if outfile:
                    self.after(0, lambda: self._status_var.set(f"Done  ·  {outfile}"))
                    self.after(0, lambda: messagebox.showinfo(
                        "Extraction complete",
                        f"Saved to:\n{outfile.resolve()}"))
                else:
                    self.after(0, lambda: self._status_var.set("No CS rows found"))
                    self.after(0, lambda: messagebox.showwarning(
                        "Nothing found",
                        "No CS rows were found in the selected files."))
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

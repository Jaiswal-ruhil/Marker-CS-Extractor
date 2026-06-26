import sys, threading
from pathlib import Path
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

CS_NAMES = {"cs", "c/s", "case", "cases", "case qty", "case quantity"}

# ── Core logic ────────────────────────────────────────────────────────────────

def find_cs(headers):
    for i, h in enumerate(headers):
        n = h.lower().strip().replace(".", "").replace("_", " ")
        if n in CS_NAMES:
            return i
    return None

def extract_from_pdf(pdf_path: Path, log):
    rows_out = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 2:
                    continue
                # First row = headers
                headers = [str(c).strip() if c else f"Column_{i+1}"
                           for i, c in enumerate(table[0])]
                cs_idx = find_cs(headers)
                if cs_idx is None:
                    continue
                for row in table[1:]:
                    if not row:
                        continue
                    vals = [str(c).strip() if c else "" for c in row]
                    vals.extend([""] * (len(headers) - len(vals)))
                    rec = dict(zip(headers, vals))
                    try:
                        cs = float(str(rec[headers[cs_idx]]).replace(",", ""))
                    except:
                        cs = 0
                    if cs > 0:
                        rec["Source PDF"] = pdf_path.name
                        rows_out.append(rec)
    log(f"  → {len(rows_out)} rows found in {pdf_path.name}")
    return rows_out

def process(pdfs, log=print):
    all_rows = []
    for pdf in pdfs:
        log(f"Processing {pdf.name}…")
        try:
            all_rows.extend(extract_from_pdf(pdf, log))
        except Exception as e:
            log(f"  ⚠ Error: {e}")
    if not all_rows:
        log("No matching CS rows found.")
        return None
    df = pd.DataFrame(all_rows)
    cols = ["Source PDF"] + [c for c in df.columns if c != "Source PDF"]
    df = df[cols]
    out = Path("output")
    out.mkdir(exist_ok=True)
    outfile = out / "Extracted.xlsx"
    df.to_excel(outfile, index=False)
    log(f"✅ Saved: {outfile.resolve()}")
    return outfile

# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Marker CS Extractor")
        self.resizable(False, False)
        self.configure(bg="#1e1e2e")
        self._files = []
        self._build_ui()

    def _build_ui(self):
        PAD    = 16
        BG     = "#1e1e2e"
        CARD   = "#2a2a3d"
        ACCENT = "#7c6af7"
        TEXT   = "#cdd6f4"
        SUBTEXT = "#6c7086"

        # Drop zone
        self._drop_frame = tk.Frame(self, bg=CARD, bd=0, width=460, height=180)
        self._drop_frame.pack(padx=PAD, pady=(PAD, 8))
        self._drop_frame.pack_propagate(False)

        tk.Label(self._drop_frame, text="📄", font=("Segoe UI Emoji", 36),
                 bg=CARD, fg=TEXT).pack(pady=(28, 4))
        self._drop_label = tk.Label(
            self._drop_frame, text="Click to choose PDF files",
            font=("Segoe UI", 12), bg=CARD, fg=SUBTEXT)
        self._drop_label.pack()

        for w in (self._drop_frame, self._drop_label):
            w.bind("<Button-1>", lambda e: self._pick_files())
            w.bind("<Enter>",    lambda e: self._drop_frame.config(bg="#32324a"))
            w.bind("<Leave>",    lambda e: self._drop_frame.config(bg=CARD))

        # File list
        self._list_var = tk.StringVar(value="")
        self._listbox = tk.Listbox(
            self, listvariable=self._list_var,
            font=("Segoe UI", 10), bg=CARD, fg=TEXT,
            selectbackground=ACCENT, activestyle="none",
            borderwidth=0, highlightthickness=0, height=5)
        self._listbox.pack(fill="x", padx=PAD, pady=(0, 8))

        btn_row = tk.Frame(self, bg=BG)
        btn_row.pack(fill="x", padx=PAD, pady=(0, 8))
        tk.Button(btn_row, text="＋ Add more", font=("Segoe UI", 10),
                  bg=CARD, fg=TEXT, activebackground="#32324a",
                  activeforeground=TEXT, relief="flat", cursor="hand2",
                  command=self._pick_files).pack(side="left")
        tk.Button(btn_row, text="✕ Clear", font=("Segoe UI", 10),
                  bg=CARD, fg=TEXT, activebackground="#32324a",
                  activeforeground=TEXT, relief="flat", cursor="hand2",
                  command=self._clear).pack(side="left", padx=8)

        # Progress bar
        self._progress = ttk.Progressbar(self, mode="indeterminate", length=460)
        self._progress.pack(padx=PAD, pady=(0, 8))
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TProgressbar", troughcolor=CARD,
                        background=ACCENT, thickness=6)

        # Run button
        self._run_btn = tk.Button(
            self, text="Extract CS Data",
            font=("Segoe UI", 12, "bold"),
            bg=ACCENT, fg="white",
            activebackground="#6a5be0", activeforeground="white",
            relief="flat", cursor="hand2", pady=10,
            command=self._run)
        self._run_btn.pack(fill="x", padx=PAD, pady=(0, 8))

        # Log
        self._log = tk.Text(
            self, height=6, font=("Courier New", 9),
            bg=CARD, fg=TEXT, insertbackground=TEXT,
            borderwidth=0, highlightthickness=0, state="disabled")
        self._log.pack(fill="x", padx=PAD, pady=(0, PAD))

        self.geometry(f"492x{self.winfo_reqheight()}")

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

    def _refresh_list(self):
        names = [Path(f).name for f in self._files]
        self._list_var.set("\n".join(names))
        if names:
            self._drop_label.config(
                text=f"{len(names)} file{'s' if len(names) > 1 else ''} selected")
        else:
            self._drop_label.config(text="Click to choose PDF files")

    def _log_write(self, msg):
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _run(self):
        if not self._files:
            messagebox.showwarning("No files", "Please add at least one PDF.")
            return
        self._run_btn.config(state="disabled")
        self._progress.start(12)
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

        pdfs = [Path(f) for f in self._files]

        def worker():
            try:
                outfile = process(pdfs, log=self._log_write)
                if outfile:
                    self.after(0, lambda: messagebox.showinfo(
                        "Done", f"Saved to:\n{outfile.resolve()}"))
                else:
                    self.after(0, lambda: messagebox.showwarning(
                        "No data", "No CS rows found in the selected PDFs."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self._progress.stop()
        self._run_btn.config(state="normal")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        process([Path(p) for p in sys.argv[1:]])
    else:
        App().mainloop()
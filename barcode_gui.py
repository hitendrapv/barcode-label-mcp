"""
Barcode Generator & Lookup — GUI Application
Reads product data from Excel, lets you search/lookup products,
preview barcodes, and generate a printable PDF.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import io
import sys

import openpyxl
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Spacer, Paragraph, Image as RLImage)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER


# ── Colour palette ──────────────────────────────────────────────────────────
BG        = "#1e1e2e"
SURFACE   = "#2a2a3e"
SURFACE2  = "#313149"
ACCENT    = "#7c6af7"
ACCENT2   = "#5a4fcf"
TEXT      = "#e0e0f0"
TEXT_DIM  = "#9090b0"
SUCCESS   = "#4caf85"
WARNING   = "#f0a050"
DANGER    = "#e05050"
BORDER    = "#44446a"
WHITE     = "#ffffff"

BARCODE_TYPES = ["ean13", "code128", "code39", "upca", "isbn13"]

LABEL_FORMATS = {
    "Standard":   "Name + barcode + extra fields",
    "GS1-128":    "GS1 Application Identifiers (GTIN, Lot, Expiry)",
    "Retail":     "Price-tag style — name, price, EAN barcode",
    "Warehouse":  "Minimal — large barcode, SKU prominent",
}


# ══════════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════════

def read_excel(filepath, has_header=True):
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    headers, rows = [], []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 and has_header:
            headers = [str(c) if c is not None else f"Col{i}" for i, c in enumerate(row)]
            continue
        if all(c is None for c in row):
            continue
        rows.append([str(c).strip() if c is not None else "" for c in row])
    wb.close()
    return headers, rows


def make_barcode_pil(value, bc_type="ean13", width_px=320, height_px=120):
    """Return a PIL Image of the barcode, or None on error."""
    value = value.strip().replace(" ", "")
    try:
        BarcodeClass = {
            "ean13":   barcode.EAN13,
            "code128": barcode.Code128,
            "code39":  barcode.Code39,
            "upca":    barcode.UPCA,
            "isbn13":  barcode.ISBN13,
        }.get(bc_type.lower(), barcode.Code128)

        if bc_type == "ean13":
            digits = "".join(filter(str.isdigit, value))
            value  = digits[:12].zfill(12)

        bc  = BarcodeClass(value, writer=ImageWriter())
        buf = io.BytesIO()
        bc.write(buf, options={
            "module_width": 0.8, "module_height": 12.0,
            "font_size": 8,      "text_distance": 3.0,
            "quiet_zone": 2.0,   "write_text": True,
        })
        buf.seek(0)
        img = Image.open(buf).convert("RGB")
        img = img.resize((width_px, height_px), Image.LANCZOS)
        return img, None
    except Exception as e:
        return None, str(e)


def _make_label_cell(label_val, bc_val, extras, bc_type, fmt):
    """Build a single label Table cell based on the chosen format."""
    styles  = getSampleStyleSheet()
    center  = ParagraphStyle("c",   parent=styles["Normal"], alignment=TA_CENTER, fontSize=8)
    bold_c  = ParagraphStyle("bc",  parent=styles["Normal"], alignment=TA_CENTER,
                              fontSize=9,  fontName="Helvetica-Bold")
    small   = ParagraphStyle("sm",  parent=styles["Normal"], alignment=TA_CENTER, fontSize=7)
    price_s = ParagraphStyle("pr",  parent=styles["Normal"], alignment=TA_CENTER,
                              fontSize=14, fontName="Helvetica-Bold")

    if fmt == "GS1-128":
        # Force Code128, prefix value with GS1 AI (01) for GTIN display
        actual_bc_type = "code128"
        digits = "".join(filter(str.isdigit, bc_val))
        gtin   = digits.zfill(14)[:14]
        gs1_val = f"\x1d01{gtin}"   # FNC1 + AI 01
        buf, err = _barcode_bytes(gtin, actual_bc_type)
        if err:
            return None
        img = RLImage(buf, width=52*mm, height=16*mm)
        ai_text = f"(01) {gtin[:2]} {gtin[2:7]} {gtin[7:12]} {gtin[12:]}"
        rows_ = [
            [Paragraph(label_val, bold_c)],
            [img],
            [Paragraph(ai_text, small)],
        ]
        for ex in extras[:2]:
            rows_.append([Paragraph(ex, small)])
        col_w, bg = 58*mm, colors.HexColor("#fffdf0")

    elif fmt == "Retail":
        # Price-tag style: name top, big price, EAN barcode bottom
        actual_bc_type = "ean13"
        buf, err = _barcode_bytes(bc_val, actual_bc_type)
        if err:
            return None
        img = RLImage(buf, width=48*mm, height=16*mm)
        price = extras[0] if extras else ""
        rows_ = [
            [Paragraph(label_val, bold_c)],
            [Paragraph(price, price_s)],
            [img],
        ]
        col_w, bg = 52*mm, colors.HexColor("#fff8f8")

    elif fmt == "Warehouse":
        # Minimal: large barcode, SKU/code prominent, no frills
        buf, err = _barcode_bytes(bc_val, bc_type)
        if err:
            return None
        img = RLImage(buf, width=56*mm, height=22*mm)
        sku = ParagraphStyle("sk", parent=styles["Normal"], alignment=TA_CENTER,
                             fontSize=11, fontName="Courier-Bold")
        rows_ = [
            [Paragraph(bc_val, sku)],
            [img],
            [Paragraph(label_val, small)],
        ]
        col_w, bg = 60*mm, colors.HexColor("#f4f4f4")

    else:  # Standard
        buf, err = _barcode_bytes(bc_val, bc_type)
        if err:
            return None
        img = RLImage(buf, width=48*mm, height=18*mm)
        rows_ = [[Paragraph(label_val, bold_c)], [img]]
        for ex in extras:
            rows_.append([Paragraph(ex, center)])
        col_w, bg = 52*mm, colors.whitesmoke

    ct = Table(rows_, colWidths=[col_w])
    ct.setStyle(TableStyle([
        ("ALIGN",         (0,0),(-1,-1),"CENTER"),
        ("VALIGN",        (0,0),(-1,-1),"MIDDLE"),
        ("BOX",           (0,0),(-1,-1), 0.5, colors.grey),
        ("BACKGROUND",    (0,0),(-1,-1), bg),
        ("TOPPADDING",    (0,0),(-1,-1), 3),
        ("BOTTOMPADDING", (0,0),(-1,-1), 3),
    ]))
    return ct, col_w


def build_pdf(rows, label_col, bc_col, extra_cols, bc_type, per_row,
              title, output_path, label_format="Standard"):
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=10*mm, leftMargin=10*mm,
                            topMargin=15*mm,  bottomMargin=10*mm)
    styles  = getSampleStyleSheet()
    title_s = ParagraphStyle("t", parent=styles["Heading1"], alignment=TA_CENTER,
                              fontSize=16, spaceAfter=6)
    story = [Paragraph(title, title_s), Spacer(1, 4*mm)]

    cells, col_w_used, skipped = [], 52*mm, 0
    for row in rows:
        if len(row) <= bc_col:
            skipped += 1; continue
        bc_val    = row[bc_col]
        label_val = row[label_col] if label_col < len(row) else ""
        extras    = [row[c] for c in extra_cols if c < len(row)]
        if not bc_val:
            skipped += 1; continue

        result = _make_label_cell(label_val, bc_val, extras, bc_type, label_format)
        if result is None:
            skipped += 1; continue
        ct, col_w_used = result
        cells.append(ct)

    page_w = A4[0] - 20*mm
    col_w  = page_w / per_row
    grid_rows = []
    for i in range(0, len(cells), per_row):
        chunk = cells[i:i+per_row]
        while len(chunk) < per_row:
            chunk.append("")
        grid_rows.append(chunk)

    if grid_rows:
        grid = Table(grid_rows, colWidths=[col_w]*per_row, hAlign="CENTER")
        grid.setStyle(TableStyle([
            ("VALIGN", (0,0),(-1,-1),"TOP"),
            ("ALIGN",  (0,0),(-1,-1),"CENTER"),
            ("LEFTPADDING",  (0,0),(-1,-1), 3),
            ("RIGHTPADDING", (0,0),(-1,-1), 3),
            ("TOPPADDING",   (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
        ]))
        story.append(grid)

    doc.build(story)
    return len(cells), skipped


def _barcode_bytes(value, bc_type):
    """Return (BytesIO, error_str) for PDF embedding."""
    value = value.strip().replace(" ", "")
    try:
        BarcodeClass = {
            "ean13":   barcode.EAN13,
            "code128": barcode.Code128,
            "code39":  barcode.Code39,
            "upca":    barcode.UPCA,
            "isbn13":  barcode.ISBN13,
        }.get(bc_type.lower(), barcode.Code128)
        if bc_type == "ean13":
            digits = "".join(filter(str.isdigit, value))
            value  = digits[:12].zfill(12)
        bc  = BarcodeClass(value, writer=ImageWriter())
        buf = io.BytesIO()
        bc.write(buf, options={"module_width":0.8,"module_height":12.0,
                               "font_size":8,"text_distance":3.0,
                               "quiet_zone":2.0,"write_text":True})
        buf.seek(0)
        return buf, None
    except Exception as e:
        return None, str(e)


# ══════════════════════════════════════════════════════════════════════════════
# Styled widget helpers
# ══════════════════════════════════════════════════════════════════════════════

def styled_btn(parent, text, command, color=ACCENT, fg=WHITE, **kw):
    b = tk.Button(parent, text=text, command=command,
                  bg=color, fg=fg, activebackground=ACCENT2, activeforeground=WHITE,
                  relief="flat", bd=0, padx=12, pady=6,
                  font=("Segoe UI", 9, "bold"), cursor="hand2", **kw)
    b.bind("<Enter>", lambda e: b.config(bg=ACCENT2))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


def styled_entry(parent, textvariable=None, width=30, **kw):
    e = tk.Entry(parent, textvariable=textvariable, width=width,
                 bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=0, font=("Segoe UI", 10),
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT, **kw)
    return e


def label(parent, text, size=9, color=TEXT, bold=False, bg=BG, **kw):
    return tk.Label(parent, text=text, bg=bg, fg=color,
                    font=("Segoe UI", size, "bold" if bold else "normal"), **kw)


def frame(parent, bg=BG, **kw):
    return tk.Frame(parent, bg=bg, **kw)


# ══════════════════════════════════════════════════════════════════════════════
# Main Application
# ══════════════════════════════════════════════════════════════════════════════

class BarcodeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🏷️  Barcode Generator & Lookup")
        self.geometry("1060x720")
        self.minsize(860, 600)
        self.configure(bg=BG)

        # State
        self.excel_path   = tk.StringVar()
        self.bc_type      = tk.StringVar(value="ean13")
        self.search_var   = tk.StringVar()
        self.label_col    = tk.IntVar(value=0)
        self.bc_col       = tk.IntVar(value=1)
        self.per_row      = tk.IntVar(value=3)
        self.label_format = tk.StringVar(value="Standard")
        self.pdf_title    = tk.StringVar(value="Product Barcode Labels")
        self.status_var   = tk.StringVar(value="No file loaded")
        self.headers      = []
        self.all_rows     = []
        self.filtered     = []
        self._bc_image    = None   # keep reference for GC
        self._selected_row = None

        self._build_ui()
        self.search_var.trace_add("write", lambda *_: self._filter_table())

    # ── UI Construction ─────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top bar ──
        top = frame(self, bg=SURFACE)
        top.pack(fill="x", padx=0, pady=0)

        tk.Label(top, text="🏷️  Barcode Generator & Lookup",
                 bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=16, pady=10)

        # Status pill
        self._status_lbl = tk.Label(top, textvariable=self.status_var,
                                    bg=SURFACE, fg=TEXT_DIM,
                                    font=("Segoe UI", 8))
        self._status_lbl.pack(side="right", padx=16)

        # ── Welcome / file-picker overlay (shown until a file is loaded) ──
        self._welcome = frame(self, bg=BG)
        self._welcome.pack(fill="both", expand=True)
        self._welcome.columnconfigure(0, weight=1)
        self._welcome.rowconfigure(0, weight=1)

        inner = frame(self._welcome, bg=SURFACE)
        inner.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(inner, text="📂", bg=SURFACE, fg="#4ea6f5",
                 font=("Segoe UI", 72)).pack(pady=(50, 10))
        tk.Label(inner, text="Open an Excel file to get started",
                 bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 18, "bold")).pack()
        tk.Label(inner, text="Select a .xlsx file containing product names and barcodes",
                 bg=SURFACE, fg=TEXT_DIM,
                 font=("Segoe UI", 12)).pack(pady=(6, 30))

        tk.Button(inner, text="  📂   Browse for Excel File…",
                  command=self._browse_and_load,
                  bg="#4ea6f5", fg=BG,
                  activebackground="#2e86d4", activeforeground=BG,
                  relief="flat", bd=0,
                  font=("Segoe UI", 15, "bold"),
                  cursor="hand2", padx=30, pady=14).pack(pady=(0, 50))

        # ── Main body (hidden until file loaded) ──
        self._main_body = frame(self)
        self._main_body.columnconfigure(0, weight=3)
        self._main_body.columnconfigure(1, weight=2)
        self._main_body.rowconfigure(0, weight=1)

        self._build_left(self._main_body)
        self._build_right(self._main_body)

    # ── LEFT PANEL ──────────────────────────────────────────────────────────

    def _build_left(self, parent):
        lf = frame(parent, bg=BG)
        lf.grid(row=0, column=0, sticky="nsew", padx=(0,6))
        lf.rowconfigure(2, weight=1)
        lf.columnconfigure(0, weight=1)

        # File loader card
        card = frame(lf, bg=SURFACE)
        card.grid(row=0, column=0, sticky="ew", pady=(0,8))
        card.columnconfigure(1, weight=1)

        label(card, "Excel File", bold=True, bg=SURFACE).grid(
            row=0, column=0, padx=(12,6), pady=(10,2), sticky="w")

        fe = styled_entry(card, textvariable=self.excel_path, width=40)
        fe.grid(row=0, column=1, padx=4, pady=(10,2), sticky="ew")

        btn_row = frame(card, bg=SURFACE)
        btn_row.grid(row=1, column=0, columnspan=2, padx=12, pady=(4,4), sticky="w")

        styled_btn(btn_row, "Browse…", self._browse_excel).pack(side="left", padx=(0,8))
        styled_btn(btn_row, "Load ▶", self._load_excel, color=SUCCESS).pack(side="left")

        # Column config row
        cfg = frame(card, bg=SURFACE)
        cfg.grid(row=2, column=0, columnspan=2, padx=12, pady=(0,10), sticky="ew")

        def cfg_spin(lbl_txt, var, col):
            tk.Label(cfg, text=lbl_txt, bg=SURFACE, fg=TEXT_DIM,
                     font=("Segoe UI", 8)).pack(side="left", padx=(0,2))
            sb = tk.Spinbox(cfg, from_=0, to=20, textvariable=var, width=3,
                            bg=SURFACE2, fg=TEXT, relief="flat", bd=0,
                            buttonbackground=SURFACE2, insertbackground=TEXT,
                            highlightthickness=1, highlightbackground=BORDER,
                            font=("Segoe UI", 9))
            sb.pack(side="left", padx=(0,12))

        cfg_spin("Label col:", self.label_col, 0)
        cfg_spin("Barcode col:", self.bc_col, 1)

        tk.Label(cfg, text="Type:", bg=SURFACE, fg=TEXT_DIM,
                 font=("Segoe UI", 8)).pack(side="left", padx=(0,2))
        cb = ttk.Combobox(cfg, textvariable=self.bc_type,
                          values=BARCODE_TYPES, width=9, state="readonly")
        cb.pack(side="left", padx=(0,4))
        cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_preview())

        # Apply ttk style for combobox
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("TCombobox", fieldbackground=SURFACE2, background=SURFACE2,
                    foreground=TEXT, selectbackground=ACCENT,
                    selectforeground=WHITE, bordercolor=BORDER)

        # Search bar
        sb_frame = frame(lf, bg=BG)
        sb_frame.grid(row=1, column=0, sticky="ew", pady=(0,6))
        sb_frame.columnconfigure(1, weight=1)

        label(sb_frame, "🔍 Search", bold=True).grid(row=0, column=0, padx=(0,6))
        se = styled_entry(sb_frame, textvariable=self.search_var, width=30)
        se.grid(row=0, column=1, sticky="ew")
        se.bind("<Escape>", lambda e: self.search_var.set(""))

        styled_btn(sb_frame, "✕", lambda: self.search_var.set(""),
                   color=SURFACE2, fg=TEXT_DIM).grid(row=0, column=2, padx=(4,0))

        self._result_lbl = label(sb_frame, "", color=TEXT_DIM, size=8)
        self._result_lbl.grid(row=0, column=3, padx=(8,0))

        # Table
        tbl_frame = frame(lf, bg=SURFACE)
        tbl_frame.grid(row=2, column=0, sticky="nsew")
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        s.configure("Treeview",
                    background=SURFACE, foreground=TEXT,
                    fieldbackground=SURFACE, rowheight=26,
                    bordercolor=BORDER, borderwidth=0)
        s.configure("Treeview.Heading",
                    background=SURFACE2, foreground=TEXT,
                    font=("Segoe UI", 9, "bold"), relief="flat")
        s.map("Treeview",
              background=[("selected", ACCENT)],
              foreground=[("selected", WHITE)])

        self.tree = ttk.Treeview(tbl_frame, show="headings",
                                 selectmode="browse", style="Treeview")
        vsb = ttk.Scrollbar(tbl_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>",         self._on_select)

    # ── RIGHT PANEL ─────────────────────────────────────────────────────────

    def _build_right(self, parent):
        rf = frame(parent, bg=BG)
        rf.grid(row=0, column=1, sticky="nsew")
        rf.columnconfigure(0, weight=1)

        # ── Preview card ──
        prev = frame(rf, bg=SURFACE)
        prev.pack(fill="x", pady=(0, 8))

        hdr = frame(prev, bg="#2f3050")
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏷️  Barcode Preview", bg="#2f3050", fg="#4ea6f5",
                 font=("Segoe UI", 12, "bold")).pack(side="left", padx=16, pady=10)
        self._fmt_badge = tk.Label(hdr, text="Standard", bg="#4ea6f5", fg=BG,
                                   font=("Segoe UI", 8, "bold"), padx=8, pady=3)
        self._fmt_badge.pack(side="right", padx=16)

        self._product_lbl = tk.Label(prev, text="← Select a product from the list",
                                     bg=SURFACE, fg=TEXT_DIM,
                                     font=("Segoe UI", 13, "bold"), wraplength=400)
        self._product_lbl.pack(padx=16, pady=(12, 6))

        self._bc_canvas = tk.Label(prev, bg="#12121f", width=460, height=160,
                                   text="No barcode loaded",
                                   fg=TEXT_DIM, font=("Segoe UI", 10))
        self._bc_canvas.pack(padx=16, pady=(0, 6), fill="x")

        self._bc_value_lbl = tk.Label(prev, text="", bg=SURFACE,
                                      fg="#4ea6f5", font=("Courier New", 13, "bold"))
        self._bc_value_lbl.pack(pady=(0, 6))

        self._detail_frame = frame(prev, bg=SURFACE)
        self._detail_frame.pack(fill="x", padx=16)

        btn_row = frame(prev, bg=SURFACE)
        btn_row.pack(fill="x", padx=16, pady=(6, 12))
        styled_btn(btn_row, "📋 Copy Barcode", self._copy_barcode,
                   color=SURFACE2, fg=TEXT).pack(side="left", padx=(0, 8))
        styled_btn(btn_row, "🖨️ Print Selected", self._print_selected,
                   color=ACCENT).pack(side="left")

        # ── Export card ──
        exp = frame(rf, bg=SURFACE)
        exp.pack(fill="x", pady=(0, 8))

        tk.Label(exp, text="Export to PDF", bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 6))

        cfg = frame(exp, bg=SURFACE)
        cfg.pack(fill="x", padx=16, pady=(0, 4))
        cfg.columnconfigure(1, weight=1)

        tk.Label(cfg, text="PDF Title:", bg=SURFACE, fg=TEXT_DIM,
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=3, padx=(0,8))
        styled_entry(cfg, textvariable=self.pdf_title, width=28).grid(
                 row=0, column=1, sticky="ew", pady=3)

        tk.Label(cfg, text="Labels per row:", bg=SURFACE, fg=TEXT_DIM,
                 font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=3, padx=(0,8))
        tk.Spinbox(cfg, from_=1, to=6, textvariable=self.per_row, width=5,
                   bg=SURFACE2, fg=TEXT, relief="flat", bd=0,
                   buttonbackground=SURFACE2, insertbackground=TEXT,
                   highlightthickness=1, highlightbackground=BORDER,
                   font=("Segoe UI", 10)).grid(row=1, column=1, sticky="w", pady=3)

        # ── Label Format ──
        tk.Label(exp, text="Label Format", bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(10, 6))

        fmt_colors = {"Standard": ACCENT, "GS1-128": "#e07b39",
                      "Retail": SUCCESS,  "Warehouse": "#4ea6f5"}

        fmt_frame = frame(exp, bg=SURFACE)
        fmt_frame.pack(fill="x", padx=16, pady=(0, 4))
        fmt_frame.columnconfigure((0,1,2,3), weight=1)

        def _make_fmt_btn(f, col):
            def _select():
                self.label_format.set(f)
                self._fmt_badge.config(text=f, bg=fmt_colors.get(f, ACCENT))
                for name, b in _fmt_btns.items():
                    b.config(bg=fmt_colors[name] if name == f else SURFACE2,
                             fg=BG if name == f else TEXT_DIM)
                self._refresh_preview()
            b = tk.Button(fmt_frame, text=f, command=_select,
                          bg=ACCENT if f == "Standard" else SURFACE2,
                          fg=BG if f == "Standard" else TEXT_DIM,
                          relief="flat", bd=0, pady=10,
                          font=("Segoe UI", 10, "bold"), cursor="hand2")
            b.grid(row=0, column=col, sticky="ew", padx=(0, 4))
            return b

        _fmt_btns = {f: _make_fmt_btn(f, i) for i, f in enumerate(LABEL_FORMATS)}

        self._fmt_desc = tk.Label(exp, text=LABEL_FORMATS["Standard"],
                                  bg=SURFACE, fg=TEXT_DIM, font=("Segoe UI", 9),
                                  wraplength=260, justify="left")
        self._fmt_desc.pack(anchor="w", padx=16, pady=(2, 8))

        def _update_fmt_desc(*_):
            self._fmt_desc.config(text=LABEL_FORMATS.get(self.label_format.get(), ""))
        self.label_format.trace_add("write", _update_fmt_desc)

        # Progress bar
        self._progress = ttk.Progressbar(exp, mode="indeterminate", length=200)
        self._progress.pack(fill="x", padx=16, pady=(0, 4))
        s = ttk.Style()
        s.configure("TProgressbar", troughcolor=SURFACE2, background=ACCENT,
                    bordercolor=SURFACE, lightcolor=ACCENT, darkcolor=ACCENT2)

        self._pdf_status = label(exp, "", color=TEXT_DIM, size=9, bg=SURFACE)
        self._pdf_status.pack(anchor="w", padx=16, pady=(0, 6))

        # ── Export buttons ──
        ebr = frame(exp, bg=SURFACE)
        ebr.pack(fill="x", padx=16, pady=(4, 16))
        ebr.columnconfigure((0, 1), weight=1)

        tk.Button(ebr, text="📄  Export All", command=self._export_all,
                  bg=SUCCESS, fg=BG, activebackground="#3a9068", activeforeground=BG,
                  relief="flat", bd=0, font=("Segoe UI", 12, "bold"),
                  cursor="hand2", pady=12).grid(row=0, column=0, sticky="ew", padx=(0, 6))

        tk.Button(ebr, text="🔍  Export Filtered", command=self._export_filtered,
                  bg=ACCENT, fg=BG, activebackground=ACCENT2, activeforeground=BG,
                  relief="flat", bd=0, font=("Segoe UI", 12, "bold"),
                  cursor="hand2", pady=12).grid(row=0, column=1, sticky="ew")

    # ── File Loading ─────────────────────────────────────────────────────────

    def _browse_and_load(self):
        self.lift()
        self.focus_force()
        path = filedialog.askopenfilename(
            parent=self,
            title="Open Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if path:
            self.excel_path.set(path)
            self._load_excel()

    def _browse_excel(self):
        self.lift()
        self.focus_force()
        path = filedialog.askopenfilename(
            parent=self,
            title="Open Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if path:
            self.excel_path.set(path)

    def _load_excel(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("No File", "Please select an Excel file first.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Not Found", f"File not found:\n{path}")
            return
        try:
            self.headers, self.all_rows = read_excel(path)
            self._populate_table(self.all_rows)
            n = len(self.all_rows)
            self._set_status(f"✅  Loaded {n} products from {os.path.basename(path)}", SUCCESS)
            # Switch from welcome screen to main UI
            self._welcome.pack_forget()
            self._main_body.pack(fill="both", expand=True, padx=12, pady=10)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            self._set_status(f"❌  Error: {e}", DANGER)

    # ── Table ────────────────────────────────────────────────────────────────

    def _populate_table(self, rows):
        self.tree.delete(*self.tree.get_children())

        cols = self.headers if self.headers else \
               [f"Col {i}" for i in range(max((len(r) for r in rows), default=4))]
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c,
                              command=lambda _c=c: self._sort_by(_c))
            self.tree.column(c, width=120, minwidth=60, anchor="w")

        for row in rows:
            # Pad short rows
            padded = row + [""] * max(0, len(cols) - len(row))
            self.tree.insert("", "end", values=padded)

        self._result_lbl.config(text=f"{len(rows)} results")

    def _filter_table(self):
        query = self.search_var.get().lower().strip()
        if not query:
            self.filtered = []
            self._populate_table(self.all_rows)
            return
        self.filtered = [r for r in self.all_rows
                         if any(query in cell.lower() for cell in r)]
        self._populate_table(self.filtered)

    def _sort_by(self, col):
        idx = self.headers.index(col) if col in self.headers else 0
        self.all_rows.sort(key=lambda r: r[idx].lower() if idx < len(r) else "")
        self._filter_table()

    # ── Selection / Preview ───────────────────────────────────────────────────

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        self._selected_row = list(values)
        self._refresh_preview()

    def _refresh_preview(self):
        row = self._selected_row
        if not row:
            return

        lc = self.label_col.get()
        bc = self.bc_col.get()
        label_val = row[lc] if lc < len(row) else "—"
        bc_val    = row[bc] if bc < len(row) else ""

        self._product_lbl.config(text=label_val)
        self._bc_value_lbl.config(text=bc_val or "No barcode value")

        # Extra detail fields
        for w in self._detail_frame.winfo_children():
            w.destroy()

        extra_indices = [i for i in range(len(row)) if i not in (lc, bc)]
        for i in extra_indices[:4]:
            hdr = self.headers[i] if i < len(self.headers) else f"Col {i}"
            val = row[i]
            if not val:
                continue
            r = frame(self._detail_frame, bg=SURFACE)
            r.pack(fill="x", pady=1)
            tk.Label(r, text=f"{hdr}:", bg=SURFACE, fg=TEXT_DIM,
                     font=("Segoe UI", 8), width=12, anchor="w").pack(side="left")
            tk.Label(r, text=val, bg=SURFACE, fg=TEXT,
                     font=("Segoe UI", 8, "bold"), anchor="w").pack(side="left")

        # Generate barcode preview in background
        if bc_val:
            threading.Thread(target=self._async_preview,
                             args=(bc_val, self.bc_type.get()), daemon=True).start()
        else:
            self._bc_canvas.config(image="", text="No barcode value", fg=TEXT_DIM)
            self._bc_image = None

    def _async_preview(self, value, bc_type):
        img, err = make_barcode_pil(value, bc_type, width_px=460, height_px=180)
        if img:
            photo = ImageTk.PhotoImage(img)
            self.after(0, self._set_preview_image, photo)
        else:
            self.after(0, self._set_preview_error, err)

    def _set_preview_image(self, photo):
        self._bc_image = photo
        self._bc_canvas.config(image=photo, text="")

    def _set_preview_error(self, err):
        self._bc_canvas.config(image="", text=f"⚠️ {err}", fg=WARNING)
        self._bc_image = None

    # ── Copy & Single Print ───────────────────────────────────────────────────

    def _copy_barcode(self):
        row = self._selected_row
        if not row:
            return
        val = row[self.bc_col.get()] if self.bc_col.get() < len(row) else ""
        if val:
            self.clipboard_clear()
            self.clipboard_append(val)
            self._set_status(f"📋  Copied: {val}", SUCCESS)

    def _print_selected(self):
        row = self._selected_row
        if not row:
            messagebox.showinfo("No Selection", "Please select a product first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile="single_label.pdf",
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        threading.Thread(target=self._do_export,
                         args=([list(row)], path, "Selected Label"),
                         daemon=True).start()

    # ── PDF Export ────────────────────────────────────────────────────────────

    def _export_all(self):
        self._start_export(self.all_rows, "all")

    def _export_filtered(self):
        src = self.filtered if self.filtered else self.all_rows
        self._start_export(src, "filtered")

    def _start_export(self, rows, kind):
        if not rows:
            messagebox.showinfo("No Data", "No products to export.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=f"barcodes_{kind}.pdf",
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        threading.Thread(target=self._do_export,
                         args=(rows, path, self.pdf_title.get()),
                         daemon=True).start()

    def _do_export(self, rows, path, title):
        self.after(0, self._progress.start, 8)
        self.after(0, self._pdf_status.config, {"text": "Generating PDF…", "fg": TEXT_DIM})
        try:
            extra = [i for i in range(max(len(r) for r in rows))
                     if i not in (self.label_col.get(), self.bc_col.get())]
            ok, skipped = build_pdf(
                rows,
                label_col    = self.label_col.get(),
                bc_col       = self.bc_col.get(),
                extra_cols   = extra,
                bc_type      = self.bc_type.get(),
                per_row      = self.per_row.get(),
                title        = title,
                output_path  = path,
                label_format = self.label_format.get(),
            )
            self.after(0, self._export_done, path, ok, skipped)
        except Exception as e:
            self.after(0, self._export_error, str(e))

    def _export_done(self, path, ok, skipped):
        self._progress.stop()
        msg = f"✅  {ok} labels saved  |  ⚠️ {skipped} skipped"
        self._pdf_status.config(text=msg, fg=SUCCESS)
        self._set_status(f"PDF saved → {os.path.basename(path)}", SUCCESS)
        if messagebox.askyesno("Done!", f"{msg}\n\nOpen the PDF now?"):
            _open_file(path)

    def _export_error(self, err):
        self._progress.stop()
        self._pdf_status.config(text=f"❌  {err}", fg=DANGER)
        messagebox.showerror("Export Failed", err)

    # ── Utilities ─────────────────────────────────────────────────────────────

    def _set_status(self, msg, color=TEXT_DIM):
        self._status_lbl.config(text=msg, fg=color)

    def run(self):
        self.mainloop()


def _open_file(path):
    import subprocess, platform
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = BarcodeApp()
    # Auto-load: CLI arg first, then products.xlsx in same folder
    preload = None
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        preload = sys.argv[1]
    else:
        default = os.path.join(os.path.dirname(os.path.abspath(__file__)), "products.xlsx")
        if os.path.exists(default):
            preload = default
    if preload:
        app.excel_path.set(preload)
        app.after(300, app._load_excel)
    app.run()

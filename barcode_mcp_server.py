#!/usr/bin/env python3
"""
Barcode Label MCP Server
Exposes tools so Claude can list products, preview barcodes, and print PDF labels.
Runs over stdio using the MCP JSON-RPC protocol (no SDK required, works on Python 3.9+).
"""

import sys
import json
import os
import io
import subprocess
import platform
import threading

# ── Ensure dependencies from the project venv are importable ─────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
_VENV = os.path.join(os.path.dirname(_HERE), "venv", "lib")
for _d in os.listdir(_VENV) if os.path.isdir(_VENV) else []:
    _sp = os.path.join(_VENV, _d, "site-packages")
    if os.path.isdir(_sp) and _sp not in sys.path:
        sys.path.insert(0, _sp)

import openpyxl
import barcode
from barcode.writer import ImageWriter
from PIL import Image
from reportlab.lib.pagesizes import A4, LETTER
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Spacer, Paragraph, Image as RLImage)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# ── Default Excel file ────────────────────────────────────────────────────────
_CANDIDATES = [
    os.path.join(_HERE, "products.xlsx"),
    os.path.join(os.path.dirname(_HERE), "products.xlsx"),
    os.path.expanduser("~/Misc/products.xlsx"),
]
DEFAULT_EXCEL = next((p for p in _CANDIDATES if os.path.exists(p)),
                     _CANDIDATES[0])

BARCODE_TYPES = {"ean13", "code128", "code39", "upca", "isbn13"}

LABEL_FORMATS = {
    "standard":   "Name + barcode + extra fields",
    "gs1-128":    "GS1-128 with Application Identifiers (GTIN AI-01)",
    "retail":     "Price-tag style — name, price, EAN barcode",
    "warehouse":  "Minimal — large barcode, SKU prominent",
}

# ─────────────────────────────────────────────────────────────────────────────
# Core helpers (shared with GUI)
# ─────────────────────────────────────────────────────────────────────────────

def read_excel(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    headers, rows = [], []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(c) if c is not None else f"Col{i}"
                       for i, c in enumerate(row)]
            continue
        if all(c is None for c in row):
            continue
        rows.append([str(c).strip() if c is not None else "" for c in row])
    wb.close()
    return headers, rows


def _barcode_bytes(value, bc_type="code128"):
    value = value.strip().replace(" ", "")
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
    bc = BarcodeClass(value, writer=ImageWriter())
    buf = io.BytesIO()
    bc.write(buf, options={"module_width": 0.8, "module_height": 12.0,
                           "font_size": 8, "text_distance": 3.0,
                           "quiet_zone": 2.0, "write_text": True})
    buf.seek(0)
    return buf


def _make_label_cell(label_val, bc_val, extras, bc_type, fmt):
    styles  = getSampleStyleSheet()
    center  = ParagraphStyle("c",  parent=styles["Normal"], alignment=TA_CENTER, fontSize=8)
    bold_c  = ParagraphStyle("bc", parent=styles["Normal"], alignment=TA_CENTER,
                              fontSize=9, fontName="Helvetica-Bold")
    small   = ParagraphStyle("sm", parent=styles["Normal"], alignment=TA_CENTER, fontSize=7)
    price_s = ParagraphStyle("pr", parent=styles["Normal"], alignment=TA_CENTER,
                              fontSize=14, fontName="Helvetica-Bold")

    fmt = fmt.lower()

    if fmt == "gs1-128":
        digits = "".join(filter(str.isdigit, bc_val))
        gtin   = digits.zfill(14)[:14]
        buf    = _barcode_bytes(gtin, "code128")
        img    = RLImage(buf, width=52*mm, height=16*mm)
        ai_txt = f"(01) {gtin[:2]} {gtin[2:7]} {gtin[7:12]} {gtin[12:]}"
        rows_  = [[Paragraph(label_val, bold_c)], [img], [Paragraph(ai_txt, small)]]
        for ex in extras[:2]:
            rows_.append([Paragraph(ex, small)])
        col_w, bg = 58*mm, colors.HexColor("#fffdf0")

    elif fmt == "retail":
        buf   = _barcode_bytes(bc_val, "ean13")
        img   = RLImage(buf, width=48*mm, height=16*mm)
        price = extras[0] if extras else ""
        rows_ = [[Paragraph(label_val, bold_c)],
                 [Paragraph(price, price_s)],
                 [img]]
        col_w, bg = 52*mm, colors.HexColor("#fff8f8")

    elif fmt == "warehouse":
        buf  = _barcode_bytes(bc_val, bc_type)
        img  = RLImage(buf, width=56*mm, height=22*mm)
        sku  = ParagraphStyle("sk", parent=styles["Normal"], alignment=TA_CENTER,
                               fontSize=11, fontName="Courier-Bold")
        rows_ = [[Paragraph(bc_val, sku)], [img], [Paragraph(label_val, small)]]
        col_w, bg = 60*mm, colors.HexColor("#f4f4f4")

    else:  # standard
        buf   = _barcode_bytes(bc_val, bc_type)
        img   = RLImage(buf, width=48*mm, height=18*mm)
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


def build_pdf(rows, headers, label_col, bc_col, bc_type, per_row,
              title, output_path, label_format="standard"):
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=10*mm, leftMargin=10*mm,
                            topMargin=15*mm,  bottomMargin=10*mm)
    styles  = getSampleStyleSheet()
    title_s = ParagraphStyle("t", parent=styles["Heading1"], alignment=TA_CENTER,
                              fontSize=16, spaceAfter=6)
    story = [Paragraph(title, title_s), Spacer(1, 4*mm)]

    extra_cols = [i for i in range(len(headers)) if i not in (label_col, bc_col)]
    cells, skipped = [], 0

    for row in rows:
        if len(row) <= bc_col:
            skipped += 1; continue
        bc_val    = row[bc_col]
        label_val = row[label_col] if label_col < len(row) else ""
        extras    = [row[c] for c in extra_cols if c < len(row)]
        if not bc_val:
            skipped += 1; continue
        try:
            result = _make_label_cell(label_val, bc_val, extras, bc_type, label_format)
            cells.append(result[0])
        except Exception:
            skipped += 1

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


def open_file(path):
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# MCP tool implementations
# ─────────────────────────────────────────────────────────────────────────────

def tool_list_products(args):
    excel = args.get("excel_file", DEFAULT_EXCEL)
    if not os.path.exists(excel):
        return error_result(f"Excel file not found: {excel}")
    headers, rows = read_excel(excel)
    lines = [" | ".join(headers)]
    lines.append("-" * 60)
    for r in rows:
        lines.append(" | ".join(r))
    return text_result(
        f"Loaded {len(rows)} products from {os.path.basename(excel)}\n\n" +
        "\n".join(lines)
    )


def tool_search_products(args):
    excel  = args.get("excel_file", DEFAULT_EXCEL)
    query  = args.get("query", "").lower().strip()
    if not os.path.exists(excel):
        return error_result(f"Excel file not found: {excel}")
    headers, rows = read_excel(excel)
    matches = [r for r in rows if any(query in cell.lower() for cell in r)]
    if not matches:
        return text_result(f"No products found matching '{query}'")
    lines = [" | ".join(headers), "-" * 60]
    for r in matches:
        lines.append(" | ".join(r))
    return text_result(f"Found {len(matches)} product(s) matching '{query}':\n\n" + "\n".join(lines))


def tool_print_labels(args):
    excel        = args.get("excel_file", DEFAULT_EXCEL)
    output       = args.get("output_pdf", os.path.join(_HERE, "labels_output.pdf"))
    query        = args.get("filter", "").lower().strip()
    label_col    = int(args.get("label_col", 0))
    bc_col       = int(args.get("bc_col", 1))
    bc_type      = args.get("bc_type", "ean13").lower()
    per_row      = int(args.get("per_row", 3))
    title        = args.get("title", "Product Barcode Labels")
    label_format = args.get("label_format", "standard").lower()
    auto_open    = args.get("auto_open", True)

    if bc_type not in BARCODE_TYPES:
        return error_result(f"Unknown barcode type '{bc_type}'. Choose from: {', '.join(BARCODE_TYPES)}")
    if label_format not in LABEL_FORMATS:
        return error_result(f"Unknown label format '{label_format}'. Choose from: {', '.join(LABEL_FORMATS)}")
    if not os.path.exists(excel):
        return error_result(f"Excel file not found: {excel}")

    headers, rows = read_excel(excel)
    if query:
        rows = [r for r in rows if any(query in cell.lower() for cell in r)]
        if not rows:
            return text_result(f"No products match filter '{query}' — no PDF generated.")

    ok, skipped = build_pdf(rows, headers, label_col, bc_col, bc_type,
                            per_row, title, output, label_format)

    if auto_open:
        open_file(output)

    return text_result(
        f"✅ PDF generated: {output}\n"
        f"   Labels printed: {ok}\n"
        f"   Skipped: {skipped}\n"
        f"   Format: {label_format}\n"
        f"   Barcode type: {bc_type}"
    )


def tool_print_single(args):
    product_name = args.get("product_name", "").strip()
    bc_value     = args.get("barcode_value", "").strip()
    label_col    = int(args.get("label_col", 0))
    bc_col       = int(args.get("bc_col", 1))
    excel        = args.get("excel_file", DEFAULT_EXCEL)
    output       = args.get("output_pdf", os.path.join(_HERE, "single_label.pdf"))
    bc_type      = args.get("bc_type", "ean13").lower()
    label_format = args.get("label_format", "standard").lower()
    auto_open    = args.get("auto_open", True)

    # If barcode_value given directly, use it; otherwise look up from Excel
    if bc_value and product_name:
        rows = [[product_name, bc_value]]
        headers = ["Product Name", "Barcode"]
        label_col, bc_col = 0, 1
    elif product_name and os.path.exists(excel):
        headers, all_rows = read_excel(excel)
        matches = [r for r in all_rows
                   if product_name.lower() in (r[label_col].lower() if label_col < len(r) else "")]
        if not matches:
            return text_result(f"No product found matching '{product_name}'")
        rows = [matches[0]]
    else:
        return error_result("Provide 'product_name' (and optionally 'barcode_value'), or both.")

    ok, skipped = build_pdf(rows, headers, label_col, bc_col, bc_type,
                            1, product_name or "Label", output, label_format)
    if auto_open:
        open_file(output)

    return text_result(
        f"✅ Single label PDF: {output}\n"
        f"   Product: {rows[0][label_col] if label_col < len(rows[0]) else ''}\n"
        f"   Barcode: {rows[0][bc_col] if bc_col < len(rows[0]) else ''}\n"
        f"   Format: {label_format}"
    )


def tool_list_formats(_args):
    lines = [f"• {k}: {v}" for k, v in LABEL_FORMATS.items()]
    bc    = [f"• {t}" for t in sorted(BARCODE_TYPES)]
    return text_result(
        "Label Formats:\n" + "\n".join(lines) +
        "\n\nBarcode Types:\n" + "\n".join(bc)
    )


# ─────────────────────────────────────────────────────────────────────────────
# MCP JSON-RPC wire protocol (stdio, no SDK)
# ─────────────────────────────────────────────────────────────────────────────

TOOLS = [
    {
        "name": "list_products",
        "description": "List all products from the Excel file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "excel_file": {
                    "type": "string",
                    "description": f"Path to the Excel file. Defaults to {DEFAULT_EXCEL}"
                }
            }
        }
    },
    {
        "name": "search_products",
        "description": "Search products by name, barcode, or any field.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query":      {"type": "string", "description": "Search term"},
                "excel_file": {"type": "string", "description": "Path to Excel file"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "print_labels",
        "description": (
            "Generate and open a PDF of barcode labels. "
            "Can filter by product name/category. "
            "Supports formats: standard, gs1-128, retail, warehouse."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "excel_file":   {"type": "string",  "description": "Path to Excel file"},
                "output_pdf":   {"type": "string",  "description": "Output PDF path"},
                "filter":       {"type": "string",  "description": "Only print matching products (optional)"},
                "label_col":    {"type": "integer", "description": "Column index of product name (default 0)"},
                "bc_col":       {"type": "integer", "description": "Column index of barcode (default 1)"},
                "bc_type":      {"type": "string",  "description": "ean13 | code128 | code39 | upca | isbn13"},
                "per_row":      {"type": "integer", "description": "Labels per row (default 3)"},
                "title":        {"type": "string",  "description": "PDF title"},
                "label_format": {"type": "string",  "description": "standard | gs1-128 | retail | warehouse"},
                "auto_open":    {"type": "boolean", "description": "Open PDF after generating (default true)"}
            }
        }
    },
    {
        "name": "print_single_label",
        "description": "Print a single barcode label for one product by name or barcode value.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "product_name":   {"type": "string",  "description": "Product name to look up"},
                "barcode_value":  {"type": "string",  "description": "Barcode value (if known directly)"},
                "excel_file":     {"type": "string",  "description": "Path to Excel file"},
                "output_pdf":     {"type": "string",  "description": "Output PDF path"},
                "bc_type":        {"type": "string",  "description": "ean13 | code128 | code39 | upca | isbn13"},
                "label_format":   {"type": "string",  "description": "standard | gs1-128 | retail | warehouse"},
                "auto_open":      {"type": "boolean", "description": "Open PDF after generating (default true)"}
            }
        }
    },
    {
        "name": "list_formats",
        "description": "List all available label formats and barcode types.",
        "inputSchema": {"type": "object", "properties": {}}
    },
]

TOOL_FNS = {
    "list_products":     tool_list_products,
    "search_products":   tool_search_products,
    "print_labels":      tool_print_labels,
    "print_single_label": tool_print_single,
    "list_formats":      tool_list_formats,
}


def text_result(text):
    return {"content": [{"type": "text", "text": text}]}

def error_result(text):
    return {"content": [{"type": "text", "text": f"❌ Error: {text}"}], "isError": True}


def send(obj):
    sys.stdout.write(json.dumps(obj) + "\n")
    sys.stdout.flush()


def handle(req):
    rid    = req.get("id")
    method = req.get("method", "")
    params = req.get("params", {})

    def reply(result):
        if rid is not None:
            send({"jsonrpc": "2.0", "id": rid, "result": result})

    def reply_error(code, msg):
        if rid is not None:
            send({"jsonrpc": "2.0", "id": rid,
                  "error": {"code": code, "message": msg}})

    if method == "initialize":
        reply({
            "protocolVersion": "2024-11-05",
            "serverInfo": {"name": "barcode-label-server", "version": "1.0.0"},
            "capabilities": {"tools": {}}
        })

    elif method == "initialized":
        pass  # notification, no reply

    elif method == "tools/list":
        reply({"tools": TOOLS})

    elif method == "tools/call":
        name = params.get("name")
        args = params.get("arguments", {})
        fn   = TOOL_FNS.get(name)
        if fn is None:
            reply_error(-32601, f"Unknown tool: {name}")
        else:
            try:
                reply(fn(args))
            except Exception as e:
                reply(error_result(str(e)))

    elif method == "ping":
        reply({})

    else:
        reply_error(-32601, f"Method not found: {method}")


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError:
            continue
        try:
            handle(req)
        except Exception as e:
            rid = req.get("id")
            if rid is not None:
                send({"jsonrpc": "2.0", "id": rid,
                      "error": {"code": -32603, "message": str(e)}})


if __name__ == "__main__":
    main()

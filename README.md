# Barcode Label MCP Server

An MCP server that generates and prints barcode label PDFs from an Excel product list.
Works with Claude Desktop — just ask Claude to print labels.

## Features

- 📋 List and search products from any `.xlsx` file
- 🏷️ Print labels for all products or a filtered subset
- 🖨️ Four label formats: **Standard**, **GS1-128**, **Retail**, **Warehouse**
- 📊 Five barcode types: EAN-13, Code128, Code39, UPC-A, ISBN-13
- 📄 Auto-opens the PDF after generation

## Example prompts

```
Print GS1 labels for all Electronics products
Print a retail label for the Power Bank
Search for USB cables and print their labels in warehouse format
List all products in my Excel file
```

## Installation

### 1. Clone the repo

```bash
git clone https://github.com/hitendrapv/barcode-label-mcp.git
cd barcode-label-mcp
```

### 2. Install dependencies

```bash
pip install openpyxl python-barcode pillow reportlab
```

### 3. Add to Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "barcode-labels": {
      "command": "python3",
      "args": ["/path/to/barcode_mcp_server.py"]
    }
  }
}
```

Restart Claude Desktop. Done.

## Excel file format

| Product Name | Barcode | Price | Category |
|---|---|---|---|
| Apple iPhone Case | 8901234567890 | $19.99 | Electronics |
| Samsung Charger | 7890123456789 | $24.99 | Electronics |

- **Product Name** — column 0 (default)
- **Barcode** — column 1 (default, EAN-13 or Code128 value)
- Additional columns appear as extra fields on the label

## Tools

| Tool | Description |
|---|---|
| `list_products` | Show all products from the Excel file |
| `search_products` | Search by name, barcode, or any field |
| `print_labels` | Generate PDF for all or filtered products |
| `print_single_label` | Print one product's label by name or barcode |
| `list_formats` | List available formats and barcode types |

## Label Formats

| Format | Description |
|---|---|
| `standard` | Product name + barcode + extra fields |
| `gs1-128` | GS1 Application Identifier layout — GTIN `(01)` prefix |
| `retail` | Price-tag style — name, price prominent, EAN barcode |
| `warehouse` | Minimal — large barcode, SKU code prominent |

## Requirements

- Python 3.9+
- openpyxl, python-barcode, pillow, reportlab

## License

MIT

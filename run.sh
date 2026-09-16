#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
"$DIR/venv/bin/python3" "$DIR/barcode_gui.py"

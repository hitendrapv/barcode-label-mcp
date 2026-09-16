#!/usr/bin/env python3
"""
Barcode Label MCP Server — Streamable HTTP transport for Smithery/hosted deployments.
Implements MCP streamable HTTP: POST /mcp returns JSON or SSE stream.
"""

import sys, os, json
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

_HERE = os.path.dirname(os.path.abspath(__file__))
_VENV = os.path.join(os.path.dirname(_HERE), "venv", "lib")
for _d in (os.listdir(_VENV) if os.path.isdir(_VENV) else []):
    _sp = os.path.join(_VENV, _d, "site-packages")
    if os.path.isdir(_sp) and _sp not in sys.path:
        sys.path.insert(0, _sp)

from barcode_mcp_server import TOOLS, TOOL_FNS, text_result, error_result

PORT = int(os.environ.get("PORT", 8000))
BASE_URL = os.environ.get("RENDER_EXTERNAL_URL", "https://barcode-label-mcp.onrender.com").rstrip("/")

SERVER_CARD = {
    "name": "barcode-label-mcp",
    "version": "1.0.0",
    "description": "Generate and print barcode label PDFs from Excel. Supports GS1-128, Retail, Warehouse and Standard formats.",
    "author": {"name": "Hitendra Venkatappa"},
    "repository": "https://github.com/hitendrapv/barcode-label-mcp",
    "transport": [{"type": "http", "url": f"{BASE_URL}/mcp"}],
    "tools": [t["name"] for t in TOOLS]
}


def handle_mcp(req):
    rid    = req.get("id")
    method = req.get("method", "")
    params = req.get("params", {})

    def resp(result):
        return {"jsonrpc": "2.0", "id": rid, "result": result}
    def err(code, msg):
        return {"jsonrpc": "2.0", "id": rid, "error": {"code": code, "message": msg}}

    if method == "initialize":
        return resp({
            "protocolVersion": "2024-11-05",
            "serverInfo": {"name": "barcode-label-server", "version": "1.0.0"},
            "capabilities": {"tools": {}}
        })
    elif method == "initialized":
        return None
    elif method == "tools/list":
        return resp({"tools": TOOLS})
    elif method == "tools/call":
        name = params.get("name")
        args = params.get("arguments", {})
        fn   = TOOL_FNS.get(name)
        if fn is None:
            return err(-32601, f"Unknown tool: {name}")
        try:
            return resp(fn(args))
        except Exception as e:
            return resp(error_result(str(e)))
    elif method == "ping":
        return resp({})
    elif rid is not None:
        return err(-32601, f"Method not found: {method}")
    return None


class MCPHandler(BaseHTTPRequestHandler):
    def log_message(self, *_):
        pass

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS, DELETE")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Accept, Mcp-Session-Id")

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def _send_json(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/health":
            self._send_json(200, {"status": "ok"})
            return

        if path == "/.well-known/mcp/server-card.json":
            self._send_json(200, SERVER_CARD)
            return

        # SSE upgrade for GET /mcp — send immediate response then close
        if path == "/mcp":
            accept = self.headers.get("Accept", "")
            if "text/event-stream" in accept:
                self.send_response(200)
                self.send_header("Content-Type", "text/event-stream")
                self.send_header("Cache-Control", "no-cache")
                self._cors()
                self.end_headers()
                # Send endpoint event pointing back to POST /mcp
                self.wfile.write(b"event: endpoint\ndata: /mcp\n\n")
                self.wfile.flush()
                return
            self._send_json(405, {"error": "Use POST for MCP requests"})
            return

        self._send_json(404, {"error": "Not found"})

    def do_POST(self):
        path   = urlparse(self.path).path
        length = int(self.headers.get("Content-Length", 0))
        body   = self.rfile.read(length)

        if path != "/mcp":
            self._send_json(404, {"error": "Not found"})
            return

        try:
            req = json.loads(body)
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Invalid JSON"})
            return

        # Handle batch (list of requests)
        if isinstance(req, list):
            results = [r for r in (handle_mcp(r) for r in req) if r is not None]
            accept = self.headers.get("Accept", "")
            if "text/event-stream" in accept:
                self.send_response(200)
                self.send_header("Content-Type", "text/event-stream")
                self._cors()
                self.end_headers()
                for r in results:
                    line = f"data: {json.dumps(r)}\n\n".encode()
                    self.wfile.write(line)
                self.wfile.flush()
            else:
                self._send_json(200, results)
            return

        result = handle_mcp(req)
        if result is None:
            self.send_response(202)
            self._cors()
            self.end_headers()
            return

        accept = self.headers.get("Accept", "")
        if "text/event-stream" in accept:
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self._cors()
            self.end_headers()
            self.wfile.write(f"data: {json.dumps(result)}\n\n".encode())
            self.wfile.flush()
        else:
            self._send_json(200, result)

    def do_DELETE(self):
        # Session termination — just acknowledge
        self.send_response(200)
        self._cors()
        self.end_headers()


if __name__ == "__main__":
    server = HTTPServer(("0.0.0.0", PORT), MCPHandler)
    print(f"Barcode Label MCP Server running on port {PORT}", flush=True)
    server.serve_forever()

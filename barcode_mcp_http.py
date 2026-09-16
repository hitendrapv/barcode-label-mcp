#!/usr/bin/env python3
"""
Barcode Label MCP Server — HTTP/SSE transport for hosted deployments.
Compatible with Smithery and other MCP marketplaces.
"""

import sys, os, json, asyncio, uuid
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

_HERE = os.path.dirname(os.path.abspath(__file__))
_VENV = os.path.join(os.path.dirname(_HERE), "venv", "lib")
for _d in (os.listdir(_VENV) if os.path.isdir(_VENV) else []):
    _sp = os.path.join(_VENV, _d, "site-packages")
    if os.path.isdir(_sp) and _sp not in sys.path:
        sys.path.insert(0, _sp)

# Import tool implementations from the stdio server
from barcode_mcp_server import (
    TOOLS, TOOL_FNS, text_result, error_result,
    DEFAULT_EXCEL, LABEL_FORMATS, BARCODE_TYPES
)

PORT = int(os.environ.get("PORT", 8000))

# Active SSE sessions: session_id -> asyncio.Queue
_sessions: dict[str, asyncio.Queue] = {}


def handle_mcp_request(req: dict) -> dict | None:
    """Process one JSON-RPC request, return response dict or None for notifications."""
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
        pass  # suppress access logs

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Accept")

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/health":
            body = b'{"status":"ok"}'
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self._cors()
            self.end_headers()
            self.wfile.write(body)
            return

        if path == "/sse":
            # SSE endpoint — client connects here and receives server events
            session_id = str(uuid.uuid4())
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Connection", "keep-alive")
            self._cors()
            self.end_headers()

            # Send the endpoint URL the client should POST messages to
            endpoint_event = f"event: endpoint\ndata: /message?session={session_id}\n\n"
            try:
                self.wfile.write(endpoint_event.encode())
                self.wfile.flush()
            except BrokenPipeError:
                return

            # Keep connection alive, sending queued responses
            import queue
            q: asyncio.Queue = queue.Queue()
            _sessions[session_id] = q
            try:
                while True:
                    try:
                        msg = q.get(timeout=30)
                        data = f"data: {json.dumps(msg)}\n\n"
                        self.wfile.write(data.encode())
                        self.wfile.flush()
                    except Exception:
                        # Send keepalive ping
                        try:
                            self.wfile.write(b": ping\n\n")
                            self.wfile.flush()
                        except BrokenPipeError:
                            break
            finally:
                _sessions.pop(session_id, None)
            return

        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        path    = urlparse(self.path).path
        query   = urlparse(self.path).query
        params  = dict(p.split("=") for p in query.split("&") if "=" in p)
        length  = int(self.headers.get("Content-Length", 0))
        body    = self.rfile.read(length)

        if path == "/message":
            session_id = params.get("session")
            try:
                req = json.loads(body)
            except json.JSONDecodeError:
                self.send_response(400)
                self.end_headers()
                return

            result = handle_mcp_request(req)
            if result and session_id in _sessions:
                _sessions[session_id].put(result)

            self.send_response(202)
            self._cors()
            self.end_headers()
            return

        # Also support plain HTTP JSON-RPC (non-SSE)
        if path == "/mcp":
            try:
                req    = json.loads(body)
                result = handle_mcp_request(req)
                resp   = json.dumps(result or {}).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self._cors()
                self.end_headers()
                self.wfile.write(resp)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
            return

        self.send_response(404)
        self.end_headers()


if __name__ == "__main__":
    import queue as _queue_mod
    # Patch asyncio.Queue -> queue.Queue for sync HTTP server
    import barcode_mcp_server as _bms
    server = HTTPServer(("0.0.0.0", PORT), MCPHandler)
    print(f"Barcode Label MCP Server running on port {PORT}", flush=True)
    server.serve_forever()

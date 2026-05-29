"""
JupyXL kernel server.

Exposes a single WebSocket endpoint (`/ws`) that the Excel task pane connects
to. Each connection owns its own IPython kernel, so notebook state persists
across cells for the life of that connection.

Run it:

    cd server
    pip install -r requirements.txt
    python main.py            # ws://localhost:8765/ws
    python main.py --tls      # wss://localhost:8765/ws  (uses office dev certs)

------------------------------------------------------------------------------
Wire protocol
------------------------------------------------------------------------------
Client -> server (JSON):
    {"type": "execute",   "code": "...", "cell_id": "c1"}
    {"type": "interrupt"}
    {"type": "restart"}

Server -> client (JSON), all tagged with the originating cell_id:
    {"type": "status",     "state": "busy"|"idle",  "cell_id": "c1"}
    {"type": "stream",     "name": "stdout"|"stderr", "text": "...", "cell_id": "c1"}
    {"type": "result",     "data": "<repr>",          "cell_id": "c1"}
    {"type": "html",       "data": "<table>...",      "cell_id": "c1"}
    {"type": "image",      "mime": "image/png", "data": "<base64>", "cell_id": "c1"}
    {"type": "dataframe",  "payload": {...},          "cell_id": "c1"}
    {"type": "error",      "ename": "...", "evalue": "...", "traceback": [...], "cell_id": "c1"}
    {"type": "kernel_restarted"}
"""

from __future__ import annotations

import argparse
import os

from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware

from .kernel_manager import KernelSession

DF_MIME = "application/vnd.jupyxl.dataframe+json"

app = FastAPI(title="JupyXL Kernel Server")

# In dev the task pane is served from https://localhost:3000. WebSocket
# upgrades aren't gated by CORS, but the health endpoint and any future
# REST calls are, so allow everything locally.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
async def health():
    return {"ok": True, "service": "jupyxl-kernel"}


def _translate(msg: dict, cell_id: str) -> dict | None:
    """Map a raw iopub message to the task pane wire protocol."""
    mtype = msg["header"]["msg_type"]
    content = msg["content"]

    if mtype == "status":
        return {"type": "status", "state": content.get("execution_state"), "cell_id": cell_id}

    if mtype == "stream":
        return {"type": "stream", "name": content.get("name", "stdout"),
                "text": content.get("text", ""), "cell_id": cell_id}

    if mtype in ("execute_result", "display_data", "update_display_data"):
        data = content.get("data", {})
        if DF_MIME in data:
            return {"type": "dataframe", "payload": data[DF_MIME], "cell_id": cell_id}
        if "image/png" in data:
            return {"type": "image", "mime": "image/png", "data": data["image/png"], "cell_id": cell_id}
        if "image/jpeg" in data:
            return {"type": "image", "mime": "image/jpeg", "data": data["image/jpeg"], "cell_id": cell_id}
        if "text/html" in data:
            return {"type": "html", "data": data["text/html"], "cell_id": cell_id}
        if "text/plain" in data:
            return {"type": "result", "data": data["text/plain"], "cell_id": cell_id}
        return None

    if mtype == "error":
        return {"type": "error", "ename": content.get("ename"),
                "evalue": content.get("evalue"),
                "traceback": content.get("traceback", []), "cell_id": cell_id}

    return None


@app.websocket("/ws")
async def ws_endpoint(ws: WebSocket):
    await ws.accept()
    session = KernelSession()
    try:
        await session.start()
        await ws.send_json({"type": "kernel_ready"})
    except Exception as exc:  # kernel failed to boot
        await ws.send_json({"type": "fatal", "message": f"Kernel failed to start: {exc}"})
        await ws.close()
        return

    try:
        while True:
            req = await ws.receive_json()
            kind = req.get("type")

            if kind == "execute":
                cell_id = req.get("cell_id", "")
                code = req.get("code", "")
                async for msg in session.execute(code):
                    wire = _translate(msg, cell_id)
                    if wire is not None:
                        await ws.send_json(wire)

            elif kind == "interrupt":
                await session.interrupt()

            elif kind == "restart":
                await session.restart()
                await ws.send_json({"type": "kernel_restarted"})

    except WebSocketDisconnect:
        pass
    finally:
        await session.shutdown()


def _resolve_certs():
    """Look for office-addin-dev-certs in the usual ~/.office-addin-dev-certs dir."""
    base = os.path.expanduser("~/.office-addin-dev-certs")
    cert = os.path.join(base, "localhost.crt")
    key = os.path.join(base, "localhost.key")
    if os.path.exists(cert) and os.path.exists(key):
        return cert, key
    return None, None


if __name__ == "__main__":
    import uvicorn

    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--tls", action="store_true",
                        help="Serve over wss:// using office-addin-dev-certs "
                             "(run `npx office-addin-dev-certs install` first)")
    args = parser.parse_args()

    ssl_kwargs = {}
    if args.tls:
        cert, key = _resolve_certs()
        if not cert:
            raise SystemExit(
                "TLS requested but no certs found. Run: npx office-addin-dev-certs install"
            )
        ssl_kwargs = {"ssl_certfile": cert, "ssl_keyfile": key}
        print(f"[jupyxl] serving wss://{args.host}:{args.port}/ws")
    else:
        print(f"[jupyxl] serving ws://{args.host}:{args.port}/ws")

    uvicorn.run(app, host=args.host, port=args.port, **ssl_kwargs)

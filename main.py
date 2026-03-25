"""
Entry point for the Excel Knowledge Graph server.

Usage:
    uv run python main.py [<data_folder>] [--port <port>]

If <data_folder> is omitted, defaults to ./sample_data/
If --port is omitted, defaults to 8000.
"""
from __future__ import annotations

import asyncio
import sys
import webbrowser
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from api import routes
from core.graph_builder import GraphBuilder
from core.watcher import FolderWatcher


# ------------------------------------------------------------------
# Data directory
# ------------------------------------------------------------------

def _parse_args() -> tuple[Path, int]:
    args = sys.argv[1:]
    port = 8000
    data_dir_str: str | None = None

    i = 0
    while i < len(args):
        if args[i] == "--port" and i + 1 < len(args):
            try:
                port = int(args[i + 1])
            except ValueError:
                print(f"[KG] Error: invalid port '{args[i + 1]}'")
                sys.exit(1)
            i += 2
        elif not args[i].startswith("--"):
            data_dir_str = args[i]
            i += 1
        else:
            print(f"[KG] Unknown argument: {args[i]}")
            sys.exit(1)

    if data_dir_str:
        p = Path(data_dir_str).resolve()
        if not p.is_dir():
            print(f"[KG] Error: '{p}' is not a directory.")
            sys.exit(1)
    else:
        p = (Path(__file__).parent / "sample_data").resolve()
        p.mkdir(exist_ok=True)

    return p, port


DATA_DIR, PORT = _parse_args()

# ------------------------------------------------------------------
# Core services
# ------------------------------------------------------------------

builder = GraphBuilder(DATA_DIR)


async def _on_file_change(event_type: str, file_path: str) -> None:
    """Watchdog callback: incrementally update the graph and notify clients."""
    p = Path(file_path)
    try:
        rel = p.relative_to(DATA_DIR)
    except ValueError:
        return  # file outside data dir

    node_id = str(rel).replace("\\", "/")
    print(f"[KG] File {event_type}: {node_id}")

    if event_type == "deleted":
        builder.remove_node(node_id)
    else:
        builder.update_node(node_id)

    await routes.broadcast({"type": "update", "event": event_type, "node_id": node_id})


watcher = FolderWatcher(DATA_DIR, _on_file_change)

# ------------------------------------------------------------------
# FastAPI lifespan
# ------------------------------------------------------------------

@asynccontextmanager
async def lifespan(app: FastAPI):
    print(f"[KG] Scanning data directory: {DATA_DIR}")
    builder.build(include_implicit=True)
    print(f"[KG] Loaded {builder.node_count} node(s)")

    loop = asyncio.get_event_loop()
    watcher.start(loop)

    yield  # server is running

    watcher.stop()
    print("[KG] Shutdown complete.")


# ------------------------------------------------------------------
# App
# ------------------------------------------------------------------

app = FastAPI(title="Excel Knowledge Graph", lifespan=lifespan)

# Wire module-level globals in routes
routes.graph_builder = builder
routes.base_dir = DATA_DIR
routes.ws_connections = []

app.include_router(routes.router)

FRONTEND_DIR = Path(__file__).parent / "frontend"
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


@app.get("/")
async def index() -> FileResponse:
    return FileResponse(str(FRONTEND_DIR / "index.html"))


# ------------------------------------------------------------------
# Run
# ------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn

    print(f"[KG] Data directory : {DATA_DIR}")
    print(f"[KG] Opening        : http://localhost:{PORT}")
    webbrowser.open(f"http://localhost:{PORT}")
    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="warning")

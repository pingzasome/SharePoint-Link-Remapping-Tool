from __future__ import annotations

import os
import uuid
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from app.exporter import export_results_csv
from app.graph_client import GraphClient, GraphConfig
from app.matcher import run_exact_matching, summarize_results
from app.parser import parse_replace_magic_file


BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
LATEST_OUTPUT = OUTPUT_DIR / "latest-results.csv"

load_dotenv(BASE_DIR / ".env")

app = FastAPI(title="ReplaceMagic SharePoint URL Matcher")
app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "app" / "templates")

STATE: dict[str, object] = {"rows": [], "results": []}


class ConfigPayload(BaseModel):
    tenantId: str | None = None
    clientId: str | None = None
    clientSecret: str | None = None
    sharepointHost: str | None = None
    sharepointSitePath: str | None = None


@app.get("/", response_class=HTMLResponse)
async def home(request: Request) -> HTMLResponse:
    defaults = {
        "tenantId": os.getenv("TENANT_ID", ""),
        "clientId": os.getenv("CLIENT_ID", ""),
        "sharepointHost": os.getenv("SHAREPOINT_HOST", "yourtenant.sharepoint.com"),
        "sharepointSitePath": os.getenv("SHAREPOINT_SITE_PATH", "/sites/TestReplaceMagic"),
    }
    return templates.TemplateResponse("index.html", {"request": request, "defaults": defaults})


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)) -> dict:
    content = await file.read()
    if not file.filename:
        raise HTTPException(status_code=400, detail="Missing upload filename.")

    try:
        rows = parse_replace_magic_file(file.filename, content)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = Path(file.filename).name
    upload_path = UPLOAD_DIR / f"{uuid.uuid4().hex}-{safe_name}"
    upload_path.write_bytes(content)

    STATE["rows"] = rows
    STATE["results"] = []

    return {
        "filename": file.filename,
        "totalRows": len(rows),
        "preview": rows[:25],
    }


@app.post("/api/run")
async def run_matching(config: ConfigPayload) -> dict:
    rows = STATE.get("rows") or []
    if not rows:
        raise HTTPException(status_code=400, detail="Upload and preview a ReplaceMagic file before running matching.")

    graph_config = _build_graph_config(config)
    try:
        graph_client = GraphClient(graph_config)
        results = run_exact_matching(rows, graph_client)
        export_results_csv(results, LATEST_OUTPUT)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    STATE["results"] = results
    return {"results": results, "summary": summarize_results(results)}


@app.get("/api/export")
async def export_csv() -> FileResponse:
    results = STATE.get("results") or []
    if not results or not LATEST_OUTPUT.exists():
        raise HTTPException(status_code=404, detail="No results are available to export yet.")

    return FileResponse(
        LATEST_OUTPUT,
        media_type="text/csv",
        filename="replace-magic-sharepoint-matches.csv",
    )


@app.post("/api/clear")
async def clear_state() -> dict[str, bool]:
    STATE["rows"] = []
    STATE["results"] = []
    if LATEST_OUTPUT.exists():
        LATEST_OUTPUT.unlink()
    return {"ok": True}


def _build_graph_config(payload: ConfigPayload) -> GraphConfig:
    return GraphConfig(
        tenant_id=payload.tenantId or os.getenv("TENANT_ID", ""),
        client_id=payload.clientId or os.getenv("CLIENT_ID", ""),
        client_secret=payload.clientSecret or os.getenv("CLIENT_SECRET", ""),
        sharepoint_host=payload.sharepointHost or os.getenv("SHAREPOINT_HOST", ""),
        sharepoint_site_path=payload.sharepointSitePath or os.getenv("SHAREPOINT_SITE_PATH", ""),
    )

from __future__ import annotations

import asyncio
import json
import threading
import traceback
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

try:
    from .converter import ConversionOptions, PdfToPptConverter, write_artifacts
except ImportError:
    from converter import ConversionOptions, PdfToPptConverter, write_artifacts


APP_ROOT = Path(__file__).resolve().parent
JOB_ROOT = APP_ROOT / ".jobs"
JOB_ROOT.mkdir(parents=True, exist_ok=True)


@dataclass
class JobState:
    job_id: str
    status: str = "queued"
    progress: int = 0
    stage: str = "排队中"
    metrics: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    error: str | None = None
    created_at: str = field(default_factory=lambda: _now_iso())
    updated_at: str = field(default_factory=lambda: _now_iso())
    workdir: Path = field(default_factory=Path)
    input_path: Path | None = None
    output_path: Path | None = None
    report_path: Path | None = None
    graph_path: Path | None = None
    traceback_text: str | None = None

    def to_public(self) -> dict[str, Any]:
        return {
            "jobId": self.job_id,
            "status": self.status,
            "progress": self.progress,
            "stage": self.stage,
            "metrics": self.metrics,
            "warnings": self.warnings,
            "error": self.error,
            "createdAt": self.created_at,
            "updatedAt": self.updated_at,
        }


app = FastAPI(title="Local High Precision PDF->PPTX API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

_jobs: dict[str, JobState] = {}
_jobs_lock = threading.Lock()


@app.get("/api/v1/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/v1/jobs")
async def create_job(
    file: UploadFile = File(...),
    options: str = Form(default="{}"),
) -> dict[str, str]:
    if not file.filename:
        raise HTTPException(status_code=400, detail="Missing file name")

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF is supported")

    try:
        payload = json.loads(options or "{}")
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid options JSON: {exc}") from exc

    pdf_bytes = await file.read()
    if not pdf_bytes:
        raise HTTPException(status_code=400, detail="Empty file")

    job_id = uuid4().hex
    workdir = JOB_ROOT / job_id
    workdir.mkdir(parents=True, exist_ok=True)

    input_path = workdir / "input.pdf"
    input_path.write_bytes(pdf_bytes)

    state = JobState(
        job_id=job_id,
        status="queued",
        progress=0,
        stage="排队中",
        workdir=workdir,
        input_path=input_path,
    )

    with _jobs_lock:
        _jobs[job_id] = state

    asyncio.create_task(_run_job(job_id, pdf_bytes, payload))
    return {"jobId": job_id}


@app.get("/api/v1/jobs/{job_id}")
def get_job(job_id: str) -> dict[str, Any]:
    state = _get_job_or_404(job_id)
    return state.to_public()


@app.get("/api/v1/jobs/{job_id}/download")
def download_job(job_id: str):
    state = _get_job_or_404(job_id)
    if state.status != "done" or not state.output_path or not state.output_path.exists():
        raise HTTPException(status_code=409, detail="Job is not completed yet")

    return FileResponse(
        path=state.output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"{job_id}.pptx",
    )


@app.get("/api/v1/jobs/{job_id}/report")
def get_report(job_id: str):
    state = _get_job_or_404(job_id)
    if state.status != "done" or not state.report_path or not state.report_path.exists():
        raise HTTPException(status_code=409, detail="Job report is not ready")
    return JSONResponse(content=json.loads(state.report_path.read_text(encoding="utf-8")))


@app.get("/api/v1/jobs/{job_id}/page-graph")
def get_page_graph(job_id: str):
    state = _get_job_or_404(job_id)
    if state.status != "done" or not state.graph_path or not state.graph_path.exists():
        raise HTTPException(status_code=409, detail="Page graph is not ready")
    return JSONResponse(content=json.loads(state.graph_path.read_text(encoding="utf-8")))


async def _run_job(job_id: str, pdf_bytes: bytes, payload: dict[str, Any]) -> None:
    option_keys = {
        "mode",
        "vector_tolerance_pt",
        "cluster_gap_pt",
        "background_filter_ratio",
        "min_icon_size_pt",
        "max_icon_size_pt",
        "debug",
    }
    option_values = {k: payload[k] for k in option_keys if k in payload}
    options = ConversionOptions(**option_values)
    converter = PdfToPptConverter(options)

    _update_job(job_id, status="running", stage="启动任务", progress=1)

    def progress_callback(value: int, stage: str, metrics: dict[str, Any] | None) -> None:
        _update_job(
            job_id,
            progress=max(0, min(100, int(value))),
            stage=stage,
            metrics=metrics or {},
        )

    try:
        artifacts = await asyncio.to_thread(converter.convert, pdf_bytes, progress_callback)
        state = _get_job_or_404(job_id)
        output_path, report_path, graph_path = await asyncio.to_thread(write_artifacts, state.workdir, artifacts)

        _update_job(
            job_id,
            status="done",
            progress=100,
            stage="完成",
            metrics={
                "vector_icons_ok": artifacts.report.get("vector_icons_ok", 0),
                "vector_icons_fallback": artifacts.report.get("vector_icons_fallback", 0),
                "text_count": artifacts.report.get("text_count", 0),
                "image_count": artifacts.report.get("image_count", 0),
            },
            warnings=list(artifacts.report.get("warnings", [])),
            output_path=output_path,
            report_path=report_path,
            graph_path=graph_path,
        )
    except Exception as exc:
        _update_job(
            job_id,
            status="failed",
            stage="失败",
            error=str(exc),
            traceback_text=traceback.format_exc(),
            progress=100,
        )


def _get_job_or_404(job_id: str) -> JobState:
    with _jobs_lock:
        state = _jobs.get(job_id)
    if not state:
        raise HTTPException(status_code=404, detail="Job not found")
    return state


def _update_job(job_id: str, **kwargs: Any) -> None:
    with _jobs_lock:
        state = _jobs.get(job_id)
        if not state:
            return
        for key, value in kwargs.items():
            if hasattr(state, key):
                setattr(state, key, value)
        state.updated_at = _now_iso()


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()

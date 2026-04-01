"""FastAPI wrapper to run DDR generation as a live HTTP service."""

from __future__ import annotations

import os
import tempfile
import threading
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

import ddr


app = FastAPI(title="DDR Report API", version="1.0.0")

# ddr.generate_ddr writes shared debug artifacts; lock the pipeline for safe server usage.
_pipeline_lock = threading.Lock()


@app.get("/health")
def health() -> JSONResponse:
    api_ready = bool(ddr.GEMINI_API_KEY)
    return JSONResponse(
        {
            "status": "ok",
            "api_key_configured": api_ready,
            "fallback_enabled": ddr.ENV_ALLOW_FALLBACK,
        }
    )


@app.post("/generate")
def generate_report(
    inspection_pdf: UploadFile = File(...),
    thermal_pdf: UploadFile = File(...),
) -> FileResponse:
    if not inspection_pdf.filename or not inspection_pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="inspection_pdf must be a .pdf file")
    if not thermal_pdf.filename or not thermal_pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="thermal_pdf must be a .pdf file")

    if not ddr.GEMINI_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="Missing GEMINI_API_KEY. Set it in environment or .env file.",
        )

    with tempfile.TemporaryDirectory(prefix="ddr_api_") as tmp_dir:
        tmp_path = Path(tmp_dir)
        inspection_path = tmp_path / "inspection.pdf"
        thermal_path = tmp_path / "thermal.pdf"
        output_path = tmp_path / "DDR_Report.docx"

        inspection_bytes = inspection_pdf.file.read()
        thermal_bytes = thermal_pdf.file.read()

        inspection_path.write_bytes(inspection_bytes)
        thermal_path.write_bytes(thermal_bytes)

        try:
            with _pipeline_lock:
                original_img_dir = ddr.IMG_DIR
                try:
                    ddr.IMG_DIR = str(tmp_path / "extracted_images")
                    ddr.generate_ddr(
                        inspection_pdf=str(inspection_path),
                        thermal_pdf=str(thermal_path),
                        api_key=ddr.GEMINI_API_KEY,
                        out_path=str(output_path),
                        allow_fallback=ddr.ENV_ALLOW_FALLBACK,
                    )
                finally:
                    ddr.IMG_DIR = original_img_dir
        except (FileNotFoundError, RuntimeError) as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        except Exception as exc:  # pragma: no cover
            raise HTTPException(status_code=500, detail=f"Unexpected server error: {exc}") from exc

        # FileResponse needs file to exist after response starts streaming, so copy to named temp.
        persistent_fd, persistent_path = tempfile.mkstemp(prefix="ddr_report_", suffix=".docx")
        os.close(persistent_fd)
        Path(persistent_path).write_bytes(output_path.read_bytes())

    return FileResponse(
        path=persistent_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="DDR_Report.docx",
        background=BackgroundTask(lambda: Path(persistent_path).unlink(missing_ok=True)),
    )

"""FastAPI adapter for JSON-to-Excel conversions."""

from __future__ import annotations

import json
from typing import Any

from fastapi import Body, FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from converter import (
    MAX_UPLOAD_BYTES,
    build_tables,
    parse_json,
    workbook_bytes,
)


app = FastAPI(
    title="JSON to Excel Converter API",
    version="1.0.0",
    description=(
        "Convert nested JSON reports into a structured Excel workbook. "
        "Upload JSON as multipart form data or send a native JSON body."
    ),
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


def create_excel_response(
    data: Any,
    filename: str,
) -> StreamingResponse:
    tables = build_tables(data, filename)
    workbook = workbook_bytes(tables)

    safe_name = (
        filename
        .rsplit("/", maxsplit=1)[-1]
        .rsplit(".", maxsplit=1)[0]
    )

    output_name = f"{safe_name or 'converted'}.xlsx"

    return StreamingResponse(
        iter([workbook]),
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": (
                f'attachment; filename="{output_name}"'
            )
        },
    )


@app.get("/")
@app.get("/api")
def api_index() -> dict[str, Any]:
    return {
        "name": "JSON to Excel Converter API",
        "version": app.version,
        "docs": "/docs",
        "endpoints": {
            "upload": "POST /convert",
            "native_json": "POST /convert/json",
            "health": "GET /health",
        },
        "max_upload_bytes": MAX_UPLOAD_BYTES,
    }


@app.get("/health")
@app.get("/api/health")
def health() -> dict[str, str]:
    return {
        "status": "ok",
        "service": "json-to-excel",
    }


@app.post("/convert")
@app.post("/api/convert")
async def convert_file(
    file: UploadFile = File(...),
    filename: str | None = Query(default=None),
) -> StreamingResponse:
    raw = await file.read(MAX_UPLOAD_BYTES + 1)

    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail="The JSON file is larger than 25 MB.",
        )

    try:
        data = parse_json(raw)
    except (
        UnicodeDecodeError,
        json.JSONDecodeError,
    ) as error:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid JSON: {error}",
        ) from error

    return create_excel_response(
        data,
        filename or file.filename or "converted.json",
    )


@app.post("/convert/json")
@app.post("/api/convert/json")
async def convert_json(
    payload: Any = Body(...),
    filename: str = Query(default="converted.json"),
) -> StreamingResponse:
    return create_excel_response(payload, filename)
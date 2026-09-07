"""Streamlit user interface for the JSON-to-Excel converter."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import streamlit as st

from converter import MAX_UPLOAD_BYTES, build_tables, parse_json, workbook_bytes


SAMPLE_PATH = Path(__file__).parent / "sample_data" / "coldstore_report.json"

st.set_page_config(
    page_title="JSON to Excel",
    page_icon="▦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --ink: #183b56; --muted: #557080; --accent: #e8753d; }
    .stApp { background: #f7fafb; }
    [data-testid="stSidebar"] { background: #183b56; }
    [data-testid="stSidebar"] * { color: #f6fbfd !important; }
    .hero { padding: 1.5rem 0 1rem; }
    .eyebrow {
        color: #e8753d;
        font-weight: 700;
        letter-spacing: .12em;
        text-transform: uppercase;
        font-size: .76rem;
    }
    .hero h1 {
        color: #183b56;
        font-size: clamp(2rem, 5vw, 3.8rem);
        line-height: .98;
        letter-spacing: -.05em;
        margin: .35rem 0 .8rem;
    }
    .hero p {
        color: #557080;
        font-size: 1.05rem;
        max-width: 680px;
    }
    .info-panel {
        background: #eaf3f6;
        border-left: 4px solid #e8753d;
        border-radius: 8px;
        padding: .85rem 1rem;
        color: #183b56;
    }
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #d9e6eb;
        border-radius: 10px;
        padding: .8rem 1rem;
    }
    div[data-testid="stDownloadButton"] button {
        background: #e8753d;
        color: white;
        border: 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_bytes(raw: bytes) -> object:
    return parse_json(raw)


def main() -> None:
    with st.sidebar:
        st.markdown("### JSON to Excel")
        st.caption("A practical converter for nested reports and system exports.")

        use_sample = st.checkbox(
            "Use the included COLDSTORE report",
            value=True,
        )

        uploaded = st.file_uploader(
            "Or upload a JSON file",
            type=["json"],
        )

        st.divider()
        st.markdown("**Limits**")
        st.caption(
            f"Maximum upload size: {MAX_UPLOAD_BYTES // (1024 * 1024)} MB"
        )
        st.caption(
            "The export keeps nested data in separate, filterable sheets."
        )

    st.markdown(
        """
        <div class="hero">
          <div class="eyebrow">Data utility / 01</div>
          <h1>Turn JSON into a workbook<br>you can actually use.</h1>
          <p>
            Upload a nested JSON report, inspect the extracted tables,
            and download a polished Excel file. The same engine is available
            through a REST API for automations.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    source_name = ""
    raw: bytes | None = None

    if uploaded is not None:
        source_name = uploaded.name
        raw = uploaded.getvalue()
    elif use_sample and SAMPLE_PATH.exists():
        source_name = "coldstore_report.json"
        raw = SAMPLE_PATH.read_bytes()

    if raw is None:
        st.info(
            "Choose the included sample report or upload a JSON file to begin."
        )
        show_api_tab()
        return

    if len(raw) > MAX_UPLOAD_BYTES:
        st.error("That file is larger than the 25 MB limit.")
        return

    try:
        data = load_bytes(raw)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        st.error(f"Could not parse this file as JSON: {error}")
        return

    tables = build_tables(data, source_name)

    record_count = len(
        tables.get("Records", pd.DataFrame())
    )
    alarm_count = len(
        tables.get("Alarms", pd.DataFrame())
    )

    stat_1, stat_2, stat_3, stat_4 = st.columns(4)

    stat_1.metric("Workbook sheets", len(tables))
    stat_2.metric("Extracted records", f"{record_count:,}")
    stat_3.metric("Alarm rows", f"{alarm_count:,}")
    stat_4.metric("JSON size", f"{len(raw) / 1024:.0f} KB")

    st.write("")

    st.markdown(
        f"""
        <div class="info-panel">
            <strong>{source_name}</strong> is ready.
            Choose a sheet below to inspect it, then download the complete workbook.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")

    preview_tab, api_tab = st.tabs(
        ["Preview & export", "Connect via API"]
    )

    with preview_tab:
        sheet_name = st.selectbox(
            "Preview sheet",
            list(tables.keys()),
        )

        preview = tables[sheet_name]

        st.dataframe(
            preview.head(100),
            use_container_width=True,
            height=440,
        )

        if len(preview) > 100:
            st.caption(
                f"Showing the first 100 of {len(preview):,} rows. "
                "The download contains all rows."
            )

        export_name = Path(source_name).stem or "converted"

        st.download_button(
            "Download Excel workbook",
            data=workbook_bytes(tables),
            file_name=f"{export_name}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=False,
        )

        st.caption(
            "The workbook includes Summary, Records, JSON Paths, "
            "and Alarms when those structures are present."
        )

    with api_tab:
        show_api_tab()


def show_api_tab() -> None:
    st.subheader("Connect this converter to another system")

    st.write(
        "The companion FastAPI service accepts either a JSON file upload "
        "or a native JSON request body and returns an Excel workbook."
    )

    st.code(
        """curl -X POST http://localhost:8008/convert \\
  -F "file=@report.json" \\
  -o report.xlsx""",
        language="bash",
    )

    st.write("For systems that already have parsed JSON:")

    st.code(
        """curl -X POST "http://localhost:8008/convert/json?filename=report.json" \\
  -H "Content-Type: application/json" \\
  --data @report.json \\
  -o report.xlsx""",
        language="bash",
    )

    st.markdown(
        """
        **API endpoints**

        - `GET /health` — service health check
        - `POST /convert` — multipart file upload
        - `POST /convert/json` — native JSON body
        - `GET /docs` — interactive API documentation
        """
    )


if __name__ == "__main__":
    main()
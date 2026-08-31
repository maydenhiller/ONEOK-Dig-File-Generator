"""ONEOK Dig File Generator.

Upload a staking report template, a cheat sheet template, one or more dig
sheets, optionally the alignment sheets and a KMZ, and get back a filled cheat
sheet plus one staking report per dig.

Everything the app can establish from the uploads is filled in. Everything it
cannot is left as Unknown or blank and flagged in the review table - a wrong
value in a signed report is worse than a missing one.
"""

from __future__ import annotations

import io
import zipfile
from datetime import date

import streamlit as st

from digfiles import aerial as aerial_module
from digfiles.alignment import apply_alignment, parse_alignment_sheet
from digfiles.cheatsheet import build_cheat_sheet
from digfiles.digsheet import parse_digsheet
from digfiles.directions import DirectionsError, generate_directions
from digfiles.kmz import parse_kmz
from digfiles.models import UNKNOWN
from digfiles.staking import (
    DEFAULT_STAKING_NOTES,
    DEFAULT_TOOL_TYPE,
    ReportSettings,
    build_staking_report,
)

st.set_page_config(page_title="ONEOK Dig File Generator", page_icon="⛽", layout="wide")

STATE_KEYS = ["digs", "sheets", "pipeline", "outputs"]
for key in STATE_KEYS:
    st.session_state.setdefault(key, None)


def mapbox_token() -> str:
    try:
        return st.secrets["mapbox"]["token"]
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

token = mapbox_token()

st.sidebar.header("Report details")
surveyor_name = st.sidebar.text_input("Surveyed by", value="")
surveyor_phone = st.sidebar.text_input("Phone number", value="")
tool_type = st.sidebar.text_input("Tool type", value=DEFAULT_TOOL_TYPE)
staking_notes = st.sidebar.text_area("Staking notes", value=DEFAULT_STAKING_NOTES, height=110)

st.sidebar.header("Aerial image")
make_aerial = st.sidebar.checkbox("Generate aerial images", value=True)
basemap = st.sidebar.selectbox("Imagery", aerial_module.basemap_options(bool(token)), index=0)
span_feet = st.sidebar.slider("Image width across the ground (ft)", 800, 6000, 2600, 200)

st.sidebar.header("Directions")
make_directions = st.sidebar.checkbox(
    "Generate driving directions",
    value=bool(token),
    disabled=not token,
    help=(
        "Uses the same Mapbox logic as the Dig Site Directions Generator. "
        "Add a token to .streamlit/secrets.toml to enable."
    ),
)
if not token:
    st.sidebar.caption("No Mapbox token found — directions will be left blank.")

st.sidebar.header("Cheat sheet")
coordinate_mode = st.sidebar.radio(
    "Latitude / longitude columns",
    ["ILI coordinates from the dig sheet", "N/A", "Leave blank"],
    index=0,
)
auto_notes = st.sidebar.checkbox("Auto-note digs close to the previous one", value=True)

COORD_MODES = {
    "ILI coordinates from the dig sheet": "ili",
    "N/A": "na",
    "Leave blank": "blank",
}


# ---------------------------------------------------------------------------
# Uploads
# ---------------------------------------------------------------------------

st.title("ONEOK Dig File Generator")
st.caption(
    "Fills the staking report and cheat sheet from your dig sheets. "
    "Latitude, longitude, elevation, DOC, survey date and photos stay manual."
)

col_left, col_right = st.columns(2)
with col_left:
    template_file = st.file_uploader(
        "Staking report template (.xlsm)", type=["xlsm", "xlsx"], key="template"
    )
    cheat_file = st.file_uploader(
        "Cheat sheet template (.xlsx)", type=["xlsx"], key="cheat"
    )
    kmz_file = st.file_uploader(
        "Pipeline KMZ (for the aerial image)", type=["kmz", "kml"], key="kmz"
    )
with col_right:
    dig_files = st.file_uploader(
        "Dig sheets — PDF or Excel, several at once",
        type=["pdf", "xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
        key="digs_upload",
    )
    alignment_files = st.file_uploader(
        "Alignment sheets (optional)",
        type=["pdf"],
        accept_multiple_files=True,
        key="alignment_upload",
    )

ready = bool(template_file and cheat_file and dig_files)
if not ready:
    st.info(
        "Upload the staking report template, the cheat sheet template and at "
        "least one dig sheet to begin."
    )


# ---------------------------------------------------------------------------
# Read the uploads
# ---------------------------------------------------------------------------

if st.button("Read uploads", type="primary", disabled=not ready):
    digs = []
    problems = []

    with st.status("Reading dig sheets…", expanded=False):
        for upload in dig_files:
            try:
                found = parse_digsheet(upload.getvalue(), upload.name)
            except Exception as error:  # noqa: BLE001 - surfaced to the user
                problems.append(f"{upload.name}: {error}")
                continue
            if not found:
                problems.append(
                    f"{upload.name}: no row carrying a Dig Number was found."
                )
            digs.extend(found)

    sheets = []
    if alignment_files:
        with st.status("Reading alignment sheets…", expanded=False):
            for upload in alignment_files:
                try:
                    sheets.append(parse_alignment_sheet(upload.getvalue(), upload.name))
                except Exception as error:  # noqa: BLE001
                    problems.append(f"{upload.name}: {error}")

    for dig in digs:
        apply_alignment(dig, sheets)
        dig.staking_notes = staking_notes

    pipeline = None
    if kmz_file is not None:
        try:
            pipeline = parse_kmz(kmz_file.getvalue(), kmz_file.name)
            if pipeline.is_empty:
                problems.append(f"{kmz_file.name}: no lines or placemarks found.")
                pipeline = None
        except Exception as error:  # noqa: BLE001
            problems.append(f"{kmz_file.name}: {error}")

    digs.sort(key=lambda d: (d.source_file, d.odometer or 0))
    st.session_state.digs = digs
    st.session_state.sheets = sheets
    st.session_state.pipeline = pipeline
    st.session_state.outputs = None

    for problem in problems:
        st.warning(problem)


digs = st.session_state.digs or []

if digs:
    st.subheader(f"{len(digs)} dig{'s' if len(digs) != 1 else ''} found")

    sheets = st.session_state.sheets or []
    if sheets:
        with st.expander("Alignment sheets read", expanded=False):
            for sheet in sheets:
                station_range = (
                    f"{sheet.station_start}–{sheet.station_end}"
                    if sheet.station_start is not None
                    else "station range not found"
                )
                st.write(
                    f"**{sheet.sheet_id}** · {sheet.route_name} · "
                    f"{sheet.county}, {sheet.state} · {station_range} · "
                    f"HCA band {'read' if sheet.hca_reliable else 'not readable'}"
                )

    st.caption(
        "Check every value before generating. Fields the uploads could not "
        "establish are marked Unknown."
    )

    edited = st.data_editor(
        [
            {
                "Dig": dig.name,
                "ODO": dig.odometer,
                "Station": dig.station_ft,
                "U/S ref": dig.upstream_reference,
                "U/S ft": dig.upstream_feet_to_agm,
                "D/S ref": dig.downstream_reference,
                "D/S ft": dig.downstream_feet_to_agm,
                "Line name": dig.line_name,
                "Alignment sheet": dig.alignment_sheet,
                "Tract": dig.tract_number,
                "Legal description": dig.legal_description,
                "County": dig.county,
                "State": dig.state,
                "HCA": dig.hca,
                "Notes": dig.notes,
            }
            for dig in digs
        ],
        width="stretch",
        num_rows="fixed",
        key="review",
    )

    warnings = [(dig.name, warning) for dig in digs for warning in dig.warnings]
    if warnings:
        with st.expander(f"{len(warnings)} thing(s) to check", expanded=True):
            for name, warning in warnings:
                st.write(f"**{name}** — {warning}")

    # ---------------------------------------------------------------
    # Generate
    # ---------------------------------------------------------------
    if st.button("Generate files", type="primary"):
        for row, dig in zip(edited, digs):
            dig.line_name = row["Line name"]
            dig.alignment_sheet = row["Alignment sheet"]
            dig.tract_number = row["Tract"]
            dig.legal_description = row["Legal description"]
            dig.county = row["County"]
            dig.state = row["State"]
            dig.hca = row["HCA"]
            dig.notes = row["Notes"] or ""

        settings = ReportSettings(
            surveyor_name=surveyor_name,
            surveyor_phone=surveyor_phone,
            tool_type=tool_type,
            staking_notes=staking_notes,
        )

        template_bytes = template_file.getvalue()
        cheat_bytes = cheat_file.getvalue()
        pipeline = st.session_state.pipeline
        outputs = {}
        issues = []

        progress = st.progress(0.0, text="Starting…")
        total = len(digs)

        for index, dig in enumerate(digs, start=1):
            progress.progress((index - 0.5) / total, text=f"{dig.name} — aerial image")

            if make_aerial:
                try:
                    dig.aerial_png = aerial_module.render_aerial(
                        dig,
                        pipeline,
                        basemap=basemap,
                        span_feet=float(span_feet),
                        token=token,
                    )
                    if dig.aerial_png is None and dig.ili_latitude is not None:
                        issues.append(f"{dig.name}: aerial image could not be rendered.")
                except Exception as error:  # noqa: BLE001
                    issues.append(f"{dig.name}: aerial image failed — {error}")

            if make_directions and token and dig.ili_latitude is not None:
                progress.progress((index - 0.25) / total, text=f"{dig.name} — directions")
                try:
                    result = generate_directions(
                        float(dig.ili_latitude), float(dig.ili_longitude), token
                    )
                    dig.directions = result.paragraph
                except Exception as error:  # noqa: BLE001
                    issues.append(f"{dig.name}: directions failed — {error}")

            progress.progress(index / total, text=f"{dig.name} — staking report")
            try:
                report, report_warnings = build_staking_report(
                    template_bytes, dig, settings
                )
                outputs[f"{dig.output_basename}.xlsm"] = report
                issues.extend(f"{dig.name}: {w}" for w in report_warnings)
            except Exception as error:  # noqa: BLE001
                issues.append(f"{dig.name}: staking report failed — {error}")

        try:
            outputs["Dig Stake Cheat Sheet.xlsx"] = build_cheat_sheet(
                cheat_bytes,
                digs,
                coordinates=COORD_MODES[coordinate_mode],
                auto_notes=auto_notes,
            )
        except Exception as error:  # noqa: BLE001
            issues.append(f"Cheat sheet failed — {error}")

        progress.empty()
        st.session_state.outputs = outputs
        for issue in issues:
            st.warning(issue)


# ---------------------------------------------------------------------------
# Downloads
# ---------------------------------------------------------------------------

outputs = st.session_state.outputs or {}
if outputs:
    st.subheader("Files")

    bundle = io.BytesIO()
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in outputs.items():
            archive.writestr(name, data)

    st.download_button(
        "Download everything (.zip)",
        data=bundle.getvalue(),
        file_name=f"Dig Files {date.today():%Y-%m-%d}.zip",
        mime="application/zip",
        type="primary",
    )

    for name, data in outputs.items():
        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if name.endswith(".xlsm")
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(name, data=data, file_name=name, mime=mime, key=f"dl_{name}")

    previews = [dig for dig in (st.session_state.digs or []) if dig.aerial_png]
    if previews:
        with st.expander("Aerial images", expanded=False):
            for dig in previews:
                st.image(dig.aerial_png, caption=dig.name, width="stretch")

"""Streamlit frontend for the validated Stage 1 Manual-to-VJR backend."""
from __future__ import annotations

import csv
import hashlib
import io
import tempfile
from pathlib import Path
from typing import Any, Callable

from stage1_app_backend import (DEFAULT_OUTPUT_ROOT, FIXED_TEMPLATE, Stage1InputError,
                               generate_vjr, preview_input)


def _read_csv(path: str | Path) -> str:
    candidate = Path(path)
    if not candidate.is_file():
        return ""
    try:
        return candidate.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeError):
        return ""


def _review_rows(path: str | Path, limit: int = 50) -> list[str]:
    raw = _read_csv(path)
    if not raw:
        return []
    rows = list(csv.DictReader(io.StringIO(raw)))
    display = []
    for row in rows[:limit]:
        category = row.get("category", "record")
        field = row.get("field") or row.get("field_status") or row.get("status", "")
        reason = row.get("reason", "Needs review")
        source_row = row.get("source_physical_row") or row.get("source_row") or row.get("csv_row")
        display.append(f"**{category.title()} · {field}** (source row {source_row}): {reason}")
    if len(rows) > limit:
        display.append(f"Showing {limit} of {len(rows)} entries. Download the record audit for the full list.")
    return display


def _invalidate_previous_result(session_state: Any | None = None) -> None:
    """Drop completed outputs as soon as Streamlit reports an upload change."""
    if session_state is None:
        import streamlit as st
        session_state = st.session_state
    session_state.pop("stage1_last_result", None)


def _upload_fingerprint(filename: str, content: bytes) -> str:
    digest = hashlib.sha256()
    digest.update(filename.encode("utf-8", errors="replace"))
    digest.update(b"\0")
    digest.update(content)
    return digest.hexdigest()


def _render_result(st: Any, result: dict[str, Any]) -> None:
    m = result
    output_path = Path(m["output"])
    run_id = str(m.get("run_id") or m.get("job_id") or "")
    if not run_id:
        st.error("The processing result has no run ID. Downloads are unavailable.")
        return
    artifacts: dict[str, bytes] = {}
    for key in ("output", "report", "audit", "unresolved"):
        path = Path(m[key])
        if not path.is_file() or path.resolve().parent != output_path.resolve().parent:
            st.error("A result artifact is missing or does not belong to this processing run. Downloads are unavailable.")
            return
        artifacts[key] = path.read_bytes()
    expected_output_hash = m.get("output_sha256")
    actual_output_hash = hashlib.sha256(artifacts["output"]).hexdigest()
    if expected_output_hash and actual_output_hash != expected_output_hash:
        st.error("The generated workbook changed after validation. Downloads are unavailable.")
        return
    status = "PASS" if m.get("formula_changes", 0) == 0 and m.get("structure_preserved", False) else "REVIEW"
    if m.get("unresolved_records", 0) or m.get("unsupported_records", 0):
        status = "REVIEW"
    st.success("Processing completed")
    st.subheader("Results")
    st.markdown(
        f"""| Result | Count |
|---|---:|
| Records processed | {m.get('records_processed', 0)} |
| Records with writes | {m.get('records_with_writes', 0)} |
| Records requiring review | {m.get('unresolved_records', 0)} |
| Unsupported records | {m.get('unsupported_records', 0)} |
| Cells written | {m.get('cells_written', 0)} |
| Formula changes | {m.get('formula_changes', 0)} |
| Workbook structure preserved | {'Yes' if m.get('structure_preserved') else 'No'} |
| Validation status | **{status}** |""")
    st.markdown("**Category summary**")
    labels = (("rectangular", "Rectangular"), ("spiral", "Spiral"), ("flex", "Flex"),
              ("riser", "Riser"), ("dampers", "Dampers"), ("elbows", "Elbows"),
              ("accessories", "Accessories"), ("other/unknown", "Unsupported/Other"))
    counts = m.get("category_counts", {})
    st.markdown("  |  ".join(f"{label}: {counts.get(key, 0)}" for key, label in labels))
    st.caption(f"Processing run ID: {run_id}")

    with st.expander("Warnings", expanded=False):
        warnings = m.get("warnings", [])
        if warnings:
            for warning in warnings: st.write(f"• {warning}")
        else:
            st.write("No input warnings.")
    with st.expander("Unresolved records", expanded=False):
        detail = _review_rows(m.get("unresolved", ""))
        if detail:
            for item in detail: st.markdown(f"- {item}")
        else:
            st.write("No unresolved records.")
    with st.expander("Unsupported records", expanded=False):
        st.write(f"Unsupported records: {m.get('unsupported_records', 0)}")
        if m.get("unsupported_records"):
            details = [line for line in _review_rows(m.get("audit", "")) if "unsupported" in line.casefold()]
            for item in (details or ["See the downloaded record audit for unsupported item details."]): st.markdown(f"- {item}")
        else:
            st.write("No unsupported records.")
    with st.expander("Validation results", expanded=False):
        st.write(f"Formula changes: {m.get('formula_changes', 0)}")
        st.write(f"Workbook structure preserved: {'Yes' if m.get('structure_preserved') else 'No'}")
        st.write("Detailed workbook, formula, and field checks are included in the downloadable Validation Report.")

    st.subheader("Downloads")
    items = [
        ("Download Generated VJR", "output", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ("Download Validation Report", "report", "text/markdown"),
        ("Download Record Audit", "audit", "text/csv"),
        ("Download Unresolved Records", "unresolved", "text/csv"),
    ]
    for label, key, mime in items:
        path = Path(m[key])
        st.download_button(label, data=artifacts[key], file_name=path.name,
                           mime=mime, key=f"{run_id}_{key}")


def main(st_module: Any | None = None, backend: Callable[..., dict[str, Any]] = generate_vjr,
         detector: Callable[[str | Path], dict[str, Any]] = preview_input) -> None:
    if st_module is None:
        import streamlit as st_module
    st = st_module
    st.set_page_config(page_title="VJR Takeoff Automation", page_icon="📄", layout="centered")
    st.markdown("""<style>.block-container{max-width:900px;padding-top:2rem}</style>""", unsafe_allow_html=True)
    st.title("VJR Takeoff Automation")
    st.write("Upload a Manual Team CSV or Excel workbook to create a VJR with the fixed project template.")
    st.subheader("Step 1 — Upload Manual Takeoff")
    if not FIXED_TEMPLATE.is_file():
        st.error("The fixed VJR template is missing. Contact the project administrator.")
        return
    upload = st.file_uploader("Manual file", type=["csv", "xlsx"], accept_multiple_files=False,
                              key="stage1_manual_upload",
                              on_change=_invalidate_previous_result)
    if upload is None:
        st.session_state.pop("stage1_last_result", None)
        st.session_state.pop("stage1_active_upload_fingerprint", None)
        return
    filename = Path(upload.name).name
    content = upload.getvalue()
    upload_fingerprint = _upload_fingerprint(filename, content)
    if st.session_state.get("stage1_active_upload_fingerprint") != upload_fingerprint:
        st.session_state.pop("stage1_last_result", None)
        st.session_state["stage1_active_upload_fingerprint"] = upload_fingerprint
    st.caption(f"Uploaded: {filename} · {len(content):,} bytes")
    if not content:
        st.error("The uploaded file is empty. Choose a non-empty .xlsx or .csv Manual.")
        return

    with tempfile.TemporaryDirectory(prefix="stage1-preview-") as temp_dir:
        input_path = Path(temp_dir) / filename
        input_path.write_bytes(content)
        try:
            detection = detector(input_path)
        except Stage1InputError as exc:
            st.error(str(exc))
            return
        except Exception:
            st.error("We could not read this workbook. Check that it is a valid .xlsx or .csv with a recognizable Manual header row.")
            return

        st.subheader("Step 2 — Input Validation")
        st.write(f"**Detected source sheet(s):** {', '.join(detection.get('sheets', [])) or 'CSV'}")
        st.write(f"**Records detected:** {detection.get('detected_input_records', detection.get('total_records', 0))}")
        pretty = {"rectangular": "Rectangular", "spiral": "Spiral", "flex": "Flex", "riser": "Riser",
                  "dampers": "Dampers", "elbows": "Elbows", "accessories": "Accessories", "other/unknown": "Unknown/Unsupported"}
        categories = detection.get("detected_input_category_counts", detection.get("category_counts", {}))
        st.write("**Detected categories:** " + ", ".join(f"{pretty.get(k, k)}: {v}" for k, v in categories.items()))
        st.write("**Important headers:** " + "; ".join(
            f"{sheet}: {', '.join(headers)}" for sheet, headers in detection.get("columns", {}).items()))
        st.write("**Required information:** " + ("Available for detection; individual gaps are listed below." if detection.get("required_information_available") else "Incomplete"))
        warnings = detection.get("warnings", [])
        if warnings:
            for warning in warnings: st.warning(warning)
        selection = detection.get("source_selection")
        if selection:
            st.info(f"Rectangular authority: {selection['authoritative_sheet']} · {selection['grouped_records']} grouped rows · {selection['detail_records']} cross-reference detail records · {selection['grouped_length_total']} m")
        st.subheader("Step 3 — Generate VJR")
        if st.button("Generate VJR", type="primary"):
            # Clear first so a failed run cannot leave a previous workbook downloadable.
            st.session_state.pop("stage1_last_result", None)
            progress = st.progress(0, text="Preparing Manual file…")
            try:
                progress.progress(15, text="Running Stage 1 automation…")
                result = backend(input_path, FIXED_TEMPLATE, DEFAULT_OUTPUT_ROOT)
                run_id = str(result.get("run_id") or result.get("job_id") or "")
                if not run_id:
                    raise Stage1InputError("Generation returned no processing run ID.")
                result["run_id"] = run_id
                result["job_id"] = run_id
                result["upload_fingerprint"] = upload_fingerprint
                result["warnings"] = warnings
                progress.progress(100, text="Validation complete")
                st.session_state["stage1_last_result"] = result
            except Stage1InputError as exc:
                st.error(str(exc))
                return
            except Exception:
                st.error("Generation failed. No VJR is available to download. Check the file and template, then try again.")
                return
    current_result = st.session_state.get("stage1_last_result")
    if current_result and current_result.get("upload_fingerprint") == upload_fingerprint:
        st.subheader("Step 4 — Results")
        _render_result(st, current_result)


if __name__ == "__main__":
    main()

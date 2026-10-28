"""Run with: python -m streamlit run streamlit_app.py"""

from __future__ import annotations

import importlib

import streamlit as st

import checker as _checker

# A Cloud update can rerun this file while Python still holds older imports.
# Refresh the engine before importing new symbols, then refresh its dependent
# adapter so detection and comparison use the same release. Ordinary reruns
# keep their module objects and session results.
_engine_refreshed = not callable(getattr(_checker, "is_datagrid_pl_viber_blast", None))
if _engine_refreshed:
    importlib.invalidate_caches()
    importlib.reload(_checker)

import web_checker as _web_checker

if _engine_refreshed or getattr(_web_checker, "AUTO_DETECTION_VERSION", 0) < 8:
    importlib.invalidate_caches()
    importlib.reload(_web_checker)

from checker import Options, find_header, is_customer_name_column, is_datagrid_pl_viber_blast, is_datagrid_viber_blast, phone_column
from web_checker import AUTO_DETECTION_VERSION, Upload, configuration_signature, prepare_uploaded_check, run_prepared_check

PREVIEW_LIMIT = 500
PASSED = {"EXACT MATCH", "MAPPED ROWS MATCH", "SAME ROWS"}


def clear_session():
    epoch = st.session_state.get("upload_epoch", 0) + 1
    st.session_state.clear()
    st.session_state.upload_epoch = epoch


def text(value):
    return "" if value is None else str(value)


def reset_results():
    st.session_state.pop("check_run", None)
    st.session_state.pop("check_signature", None)


def apply_style():
    st.markdown("""
<style>
    :root {color-scheme: dark;}
    .stMainBlockContainer {max-width: 1120px; padding-top: 2.5rem; padding-bottom: 3rem;}
    h1 {font-size: 2.2rem !important; letter-spacing: -0.06rem;}
    h3 {font-size: 1.15rem !important;}
    .st-key-upload_database_card, .st-key-upload_files_card {
        background: #121c2a; border: 1px solid #29374b; border-radius: 16px;
    }
    [data-testid="stFileUploaderDropzone"] {
        background: #182435; border: 1px dashed #35465e; border-radius: 12px;
    }
    [data-testid="stMetricValue"] {font-size: 1.8rem;}
    [data-testid="stBaseButton-primary"] {
        background: #2dd4bf; color: #062b25; border-color: #2dd4bf;
        border-radius: 10px; font-weight: 600;
    }
    [data-testid="stBaseButton-primary"] p {color: inherit;}
    [data-testid="stBaseButton-primary"]:not(:disabled):hover {
        background: #5eead4; border-color: #5eead4; color: #062b25;
    }
    [data-testid="stBaseButton-primary"]:disabled {
        background: #182432; border-color: #29374b; color: #8392a7;
    }
    [data-testid="stBaseButton-secondary"] {
        background: #1b283a; border-color: #35445b; color: #e4ecf7; border-radius: 10px;
    }
    [data-testid="stBaseButton-secondary"]:not(:disabled):hover {border-color: #2dd4bf; color: #5eead4;}
</style>
""", unsafe_allow_html=True)


def render_file_details(result, run, index):
    st.caption(result.source.path.name)
    if result.error:
        st.error(result.error)
        return
    issues_tab, rows_tab, fields_tab = st.tabs(["Issues", "Rows", "Fields checked"])
    with issues_tab:
        has_issues = False
        if result.missing_columns or result.extra_columns:
            has_issues = True
            st.warning(f"Missing columns: {', '.join(result.missing_columns) or 'None'}. Extra columns: {', '.join(result.extra_columns) or 'None'}.")
        if result.differences:
            has_issues = True
            st.dataframe([{"Row": number, "Field": name, "Your file": text(actual), "Database": text(expected)}
                          for number, database_row, name, actual, expected in result.differences[:PREVIEW_LIMIT]],
                         hide_index=True, width="stretch")
            if len(result.differences) > PREVIEW_LIMIT:
                st.caption(f"Showing {PREVIEW_LIMIT} of {len(result.differences):,} differences. Download the report for all details.")
        key_position = result.headers.index(result.source.column)
        unmatched = [{"Row": number, result.source.column: text(values[key_position]), "Issue": status}
                     for number, status, matches, values in result.records
                     if status in {"NOT FOUND", "BLANK KEY", "EXTRA ROW"}]
        if unmatched:
            has_issues = True
            st.dataframe(unmatched[:PREVIEW_LIMIT], hide_index=True, width="stretch")
            if len(unmatched) > PREVIEW_LIMIT:
                st.caption(f"Showing {PREVIEW_LIMIT} of {len(unmatched):,} row issues. Download the report for all details.")
        if result.missing_rows:
            has_issues = True
            st.caption(f"{len(result.missing_rows):,} database rows are missing from this file.")
            st.dataframe([{"Database row": number, **{header: text(value) for header, value in zip(result.database_headers, values)}}
                          for number, values in result.missing_rows[:PREVIEW_LIMIT]], hide_index=True, width="stretch")
        if not has_issues:
            st.caption("No issues found in the compared fields.")
    with rows_tab:
        choices = ["All rows", *dict.fromkeys(record[1] for record in result.records)]
        if any(record[1] != "MATCHED" for record in result.records):
            choices.insert(0, "Needs review")
        status = st.selectbox("Show rows", choices, key=f"result_filter_{index}")
        records = [record for record in result.records
                   if status == "All rows" or (status == "Needs review" and record[1] != "MATCHED") or record[1] == status]
        key_position = result.headers.index(result.source.column)
        rows = [{"Row": number, result.source.column: text(values[key_position]), "Result": row_status, "Database row": matches}
                for number, row_status, matches, values in records[:PREVIEW_LIMIT]]
        if rows:
            st.dataframe(rows, hide_index=True, width="stretch")
        else:
            st.caption("No rows for this filter.")
        st.caption(f"Showing {len(rows):,} of {len(records):,} rows. Original columns are included in the Excel report.")
    with fields_tab:
        mapped = {item.file_column: item for item in result.mappings}
        coverage = []
        for name in result.headers:
            item = mapped.get(name)
            if name == result.flow_stage_column:
                coverage.append({"File column": name, "Database column": "", "Coverage": "FILE IDENTIFIER"})
            elif run.options.mapped_fields:
                coverage.append({"File column": name, "Database column": item.database_column if item else "",
                                 "Coverage": "CHECKED" if item else "NOT CHECKED"})
            else:
                equivalent = find_header(result.database_headers, name)
                coverage.append({"File column": name, "Database column": equivalent,
                                 "Coverage": "CHECKED" if equivalent else "EXTRA COLUMN"})
        if not result.source.has_headers:
            st.caption("Only CHCODE in column B is checked for headerless files. Other columns are not checked.")
        else:
            st.caption("Results cover the compared fields. Unmapped campaign fields are not checked.")
            if any(is_customer_name_column(item.file_column) or is_customer_name_column(item.database_column)
                   for item in result.mappings):
                st.caption("Question marks are ignored when comparing customer names. Original names are kept in the report.")
            if is_datagrid_pl_viber_blast(result.headers) and result.source.column.casefold() == "debtorid":
                st.caption("Debtorid is matched to DataGrid Debtor ID; FullName and Phone must match Name and Last Called No. on the same database row. Staff and campaign fields are not checked.")
            elif is_datagrid_viber_blast(result.headers) and result.source.column.casefold() == "phone":
                st.caption("Phone is matched to DataGrid Last Called No.; FullName must match Name on the same database row. CHCode and campaign amounts are not checked.")
            elif not result.flow_stage_column and phone_column(result.headers):
                st.caption("Phone numbers are compared with the database record for the same CHCODE.")
        st.caption(f"File identifier: {result.file_identifier}")
        st.dataframe(coverage, hide_index=True, width="stretch")


def render_results(run):
    results = run.results
    review_count = sum(result.verdict not in PASSED for result in results)
    matched = sum(result.count("MATCHED") for result in results)
    heading, download = st.columns([2, 1], vertical_alignment="center")
    with heading:
        st.subheader("Check results")
    with download:
        st.download_button("Download Excel report", data=run.report, file_name="CHCODE_check_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="download_report", type="secondary", on_click="ignore", width="stretch")
    if review_count:
        st.warning(f"{review_count} of {len(results)} file(s) need review. Open a result below to see the issues.")
    else:
        st.success("All files match the database across the compared fields.")
    columns = st.columns(3)
    for column, label, value in zip(columns, ["Files checked", "Rows matched", "Files to review"],
                                    [len(results), matched, review_count]):
        column.metric(label, f"{value:,}")
    for index, result in enumerate(results):
        status = ("Mapped fields match" if result.verdict == "MAPPED ROWS MATCH" else
                  "Exact match" if result.verdict in PASSED else "File error" if result.error else "Review needed")
        stage = result.file_identifier
        label = f"{stage} · {status}"
        if not result.error:
            label += f" · {result.count('MATCHED'):,}/{len(result.records):,} rows matched"
        with st.expander(label, expanded=False, icon=":material/check_circle:" if result.verdict in PASSED else ":material/error_outline:"):
            render_file_details(result, run, index)


def main():
    st.set_page_config(page_title="File Checker", page_icon=":material/fact_check:", layout="wide", initial_sidebar_state="collapsed")
    apply_style()
    title, restart = st.columns([4, 1], vertical_alignment="center")
    with title:
        st.title("File checker")
        st.caption("Check your campaign files against your database.")
    with restart:
        st.button("Start over", on_click=clear_session, key="clear_session", type="tertiary", width="stretch")
    epoch = st.session_state.get("upload_epoch", 0)
    left, right = st.columns(2)
    with left:
        with st.container(border=True, key="upload_database_card"):
            st.subheader("1. Add your database")
            database_file = st.file_uploader("Database file", type=["xlsx", "xlsm", "csv"],
                                             key=f"database_upload_{epoch}", label_visibility="collapsed")
    with right:
        with st.container(border=True, key="upload_files_card"):
            st.subheader("2. Add files to check")
            uploaded_files = st.file_uploader("Files to check", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
                                              key=f"files_upload_{epoch}", label_visibility="collapsed")

    database_upload = Upload(database_file.name, database_file.getvalue()) if database_file is not None else None
    file_uploads = [Upload(file.name, file.getvalue()) for file in uploaded_files]
    current_inputs = (AUTO_DETECTION_VERSION, database_upload.fingerprint if database_upload else None,
                      tuple(upload.fingerprint for upload in file_uploads))
    if st.session_state.get("input_signature") != current_inputs:
        reset_results()
        st.session_state.pop("preparation", None)
        st.session_state.input_signature = current_inputs

    prepared = None
    if database_upload:
        if "preparation" not in st.session_state:
            with st.spinner("Detecting worksheets, headers, and CHCODE columns..."):
                st.session_state.preparation = prepare_uploaded_check(database_upload, file_uploads)
        prepared = st.session_state.preparation

    if prepared and prepared.database_error:
        reset_results()
        st.error(f"Database: {prepared.database_error}")
    elif prepared and prepared.database_info:
        info = prepared.database_info
        worksheet = f"{info.sheet}, " if info.sheet else ""
        st.caption(f"Database: {database_upload.name} - {worksheet}{info.column} header found on row {info.header_row}.")

    if prepared and prepared.files:
        summary = []
        for item in prepared.files:
            info = item.inspection
            summary.append({
                "File": item.source.upload.name,
                "Worksheet": item.source.sheet,
                "Header row": (str(info.header_row) if info and item.source.has_headers
                               else "No headers" if info else ""),
                "Lookup column": item.source.column if info else "",
                "Check": ("Viber debtor ID, phone and name" if info and is_datagrid_pl_viber_blast(info.headers)
                          else "Viber phone and name" if info and is_datagrid_viber_blast(info.headers)
                          else "Phone by CHCODE" if info and info.has_headers
                          and not find_header(info.headers, "FLOW_STAGE") and phone_column(info.headers)
                          else "Mapped fields"),
                "Status": item.error or "Ready",
            })
        st.dataframe(summary, hide_index=True, width="stretch")
        for item in prepared.files:
            if item.note and not item.error:
                st.caption(f"{item.source.upload.name}: {item.note}")
        file_errors = sum(bool(item.error) for item in prepared.files)
        if file_errors:
            st.warning(f"{file_errors} file(s) need attention. Details are shown above and included in the report.")
        if not prepared.database_error:
            ready_count = len(prepared.files) - file_errors
            st.caption(f"{ready_count} file(s) ready to check.")

    ready = bool(prepared and prepared.database and not prepared.database_error and prepared.files)
    if database_upload is None:
        st.caption("Add a database and your files to get started. Excel and CSV are supported.")
    elif prepared and not prepared.database_error and not prepared.files:
        st.caption("Add at least one file to check.")
    st.caption("CHCODE can be in CHCODE, primary_key, or indexone_uid. The lookup column is detected automatically. Headerless files use column B and include row 1; only CHCODE is checked.")
    st.caption("Files without FLOW_STAGE use CHCODE to compare phone numbers with the matching database record.")
    st.caption("SALAD Viber blasts use Phone and FullName to check the same DataGrid row.")
    st.caption("PL Viber blasts use Debtorid to check FullName and Phone against the same DataGrid record.")
    st.caption("Uploads are processed in this session. Start over clears the current uploads and results.")

    options = Options(ignore_case=True, trim_spaces=True, mapped_fields=True,
                      require_flow_stage=True, allow_missing_flow_stage=True)
    signature = configuration_signature(prepared.database, [item.source for item in prepared.files], options) if ready else None
    if st.session_state.get("check_run") and st.session_state.get("check_signature") != signature:
        reset_results()
    if st.button("3. Check files", type="primary", key="run_check", disabled=not ready, width="stretch"):
        reset_results()
        progress = st.progress(0, text="Reading your database...")

        def update(index, total, result):
            progress.progress(index / total, text=f"Checking file {index} of {total}...")

        try:
            with st.spinner("Checking your files..."):
                st.session_state.check_run = run_prepared_check(prepared, update)
                st.session_state.check_signature = signature
        except Exception as error:
            st.error(f"Unable to complete the check: {error}")
        finally:
            progress.empty()
    if st.session_state.get("check_run"):
        st.divider()
        render_results(st.session_state.check_run)


if __name__ == "__main__":
    main()

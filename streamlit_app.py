"""Run with: python -m streamlit run streamlit_app.py"""

from __future__ import annotations

import importlib
from html import escape

import streamlit as st

import checker as _checker

# A Cloud update can rerun this file while Python still holds older imports.
# Refresh the engine before importing new symbols, then refresh its dependent
# adapter so detection and comparison use the same release. Ordinary reruns
# keep their module objects and session results.
_engine_refreshed = (not callable(getattr(_checker, "is_datagrid_pl_viber_blast", None))
                     or getattr(_checker, "COMPARISON_VERSION", 0) < 14)
if _engine_refreshed:
    importlib.invalidate_caches()
    importlib.reload(_checker)
    import auto_match as _auto_match
    importlib.reload(_auto_match)

import web_checker as _web_checker

if _engine_refreshed or getattr(_web_checker, "AUTO_DETECTION_VERSION", 0) < 22:
    importlib.invalidate_caches()
    importlib.reload(_web_checker)

from checker import Options, balance_viber_database_fields, find_header, is_cimb_database, is_customer_name_column, is_datagrid_balance_viber_blast, is_datagrid_pl_viber_blast, is_datagrid_salad_email_blast, is_datagrid_viber_blast, is_hsbc_mobile_pair, is_salad_endo_pair, is_salad_ml_database, phone_column
from web_checker import AUTO_DETECTION_VERSION, Upload, configuration_signature, prepare_uploaded_check, run_prepared_check, run_prepared_checks

PREVIEW_LIMIT = 500
PASSED = {"EXACT MATCH", "MAPPED ROWS MATCH", "SAME ROWS"}


def clear_session():
    epoch = st.session_state.get("upload_epoch", 0) + 1
    st.session_state.clear()
    st.session_state.upload_epoch = epoch


def text(value):
    return "" if value is None else str(value)


def comparison_text(value):
    value = text(value)
    return value if value.strip() else "(blank)"


def reset_results():
    st.session_state.pop("check_run", None)
    st.session_state.pop("check_signature", None)


def apply_style():
    st.markdown("""
<style>
    :root {color-scheme: dark;}
    .stMainBlockContainer {width: 100%; max-width: none; padding: 2.5rem 1.5rem 3rem;}
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
    .result-overview {display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); align-items: start; gap: 16px; margin: 8px 0 12px;}
    .result-panel {min-width: 0; border: 1px solid #29374b; border-radius: 12px; overflow-x: auto; background: #111a27;}
    .result-panel h4 {font-size: 16px; margin: 0; padding: 12px 14px; background: #172334;}
    .result-table {border-collapse: collapse; width: 100%; font-size: 15px; line-height: 1.4;}
    .result-table th {padding: 9px 10px; color: #a6b7cc; font-size: 13px; text-align: left; font-weight: 600; background: #131e2c;}
    .result-table td {padding: 9px 10px; border-top: 1px solid #243247; vertical-align: top; overflow-wrap: anywhere;}
    .result-table .result-number {white-space: nowrap; overflow-wrap: normal; font-variant-numeric: tabular-nums;}
    .result-table .result-key {font-variant-numeric: tabular-nums;}
    .result-table .result-issue {overflow-wrap: normal;}
    .result-table .result-file {min-width: 120px;}
    .result-table .result-detail {min-width: 180px;}
    .result-difference + .result-difference {border-top: 1px solid #35445b; margin-top: 8px; padding-top: 8px;}
    .result-value {display: grid; grid-template-columns: 5rem minmax(0, 1fr); gap: 8px;}
    .result-value + .result-value {margin-top: 4px;}
    .result-value-label {color: #a6b7cc; font-size: 12px; padding-top: 2px;}
    .result-value-text {min-width: 0; overflow-wrap: anywhere; white-space: pre-wrap;}
    .result-table small {display: block; color: #a6b7cc; font-size: 12px;}
    .result-status {font-size: 12px; font-weight: 700; white-space: nowrap;}
    .result-status.pass {color: #5eead4;}
    .result-status.review {color: #fbbf24;}
    .result-status.error {color: #fda4af;}
    .result-note {padding: 10px 14px; margin: 0; color: #a6b7cc; font-size: 12px; line-height: 1.4;}
    @media (max-width: 1000px) {.result-overview {grid-template-columns: minmax(0, 1fr);}}
    @media (max-width: 600px) {
        .result-issues, .result-issues tbody {display: block;}
        .result-issues thead {display: none;}
        .result-issues tbody tr {display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); border-top: 1px solid #243247;}
        .result-issues td {display: block; min-width: 0; border: 0;}
        .result-issues td::before {content: attr(data-label); display: block; color: #a6b7cc; font-size: 12px; margin-bottom: 4px;}
        .result-issues .result-detail {grid-column: 1 / -1; min-width: 0;}
    }
</style>
""", unsafe_allow_html=True)


def result_overview_html(run, issue_limit=12):
    """A compact, escaped overview of every comparison and the first row issues."""
    multiple_databases = len(run.databases) > 1
    files, issues = [], []
    for index, result in enumerate(run.results, 1):
        database = result.database_source or run.database
        file_name = escape(result.source.path.name)
        if multiple_databases:
            file_name = f"<small>{escape(database.path.name)}</small>{file_name}"
        review_rows = sum(status != "MATCHED" for _, status, _, _ in result.records) + len(result.missing_rows)
        if result.error:
            state, label, matched, review = "error", "ERROR", "—", "—"
        else:
            state = "pass" if result.verdict in PASSED else "review"
            label = "MAPPED MATCH" if result.verdict == "MAPPED ROWS MATCH" else "MATCH" if state == "pass" else "REVIEW"
            matched, review = f"{result.count('MATCHED'):,}/{len(result.records):,}", f"{review_rows:,}"
        files.append(f'<tr><td class="result-number">{index}</td><td class="result-file">{file_name}</td>'
                     f'<td class="result-number">{matched}</td><td class="result-number">{review}</td>'
                     f'<td><span class="result-status {state}">{label}</span></td></tr>')
        if result.error:
            issues.append((index, "—", "—", "ERROR", result.error))
            continue
        differences = {}
        for number, _, _field, actual, expected in result.differences:
            differences.setdefault(number, []).append((actual, expected))
        key_position = result.headers.index(result.source.column)
        for number, status, _, values in result.records:
            if status != "MATCHED":
                issues.append((index, str(number), text(values[key_position]), status,
                               differences.get(number, [])))
        database_key = result.database_match_column or database.column
        key_position = result.database_headers.index(database_key) if result.missing_rows else None
        for number, values in result.missing_rows:
            issues.append((index, "—", text(values[key_position]), "MISSING FROM FILE", f"Database row {number}"))
        for label, columns in (("MISSING COLUMNS", result.missing_columns), ("EXTRA COLUMNS", result.extra_columns)):
            if columns:
                issues.append((index, "—", "—", label, ", ".join(columns)))
    visible = issues[:issue_limit]
    show_details = any(row[4] for row in visible)
    issue_rows = []
    for index, number, key, status, detail in visible:
        if isinstance(detail, list):
            detail_html = "".join(
                '<div class="result-difference">'
                '<div class="result-value"><span class="result-value-label">Your file</span>'
                f'<span class="result-value-text">{escape(comparison_text(actual))}</span></div>'
                '<div class="result-value"><span class="result-value-label">Database</span>'
                f'<span class="result-value-text">{escape(comparison_text(expected))}</span></div></div>'
                for actual, expected in detail)
        else:
            detail_html = escape(detail)
        issue_rows.append(f'<tr><td class="result-number" data-label="File #">{index}</td>'
                          f'<td class="result-number" data-label="Row">{number}</td>'
                          f'<td class="result-key" data-label="Lookup value">{escape(key) or "—"}</td>'
                          f'<td class="result-issue" data-label="Issue">{escape(status)}</td>'
                          + (f'<td class="result-detail" data-label="Details">{detail_html}</td>' if show_details else "") + "</tr>")
    issue_header = '<th>File #</th><th>Row</th><th>Lookup value</th><th>Issue</th>' + ('<th>Details</th>' if show_details else '')
    issue_body = (f'<table class="result-table result-issues"><thead><tr>{issue_header}</tr></thead><tbody>{"".join(issue_rows)}</tbody></table>'
                  if issues else '<p class="result-note">No issues in the compared fields.</p>')
    if len(visible) < len(issues):
        issue_body += (f'<p class="result-note">Showing {len(visible):,} of {len(issues):,} issue entries. '
                       'Full details are in the Excel report.</p>')
    return ('<div class="result-overview"><section class="result-panel"><h4>File overview</h4>'
            '<table class="result-table"><thead><tr><th>#</th><th>File</th><th>Matched</th><th>Review</th><th>Result</th></tr></thead>'
            f'<tbody>{"".join(files)}</tbody></table></section><section class="result-panel"><h4>Issues across files</h4>{issue_body}</section></div>')


def render_file_details(result, run, index):
    if result.skipped_database_rows:
        st.caption(f"Skipped {len(result.skipped_database_rows):,} database row(s) with blank {result.database_match_column} values. Source row numbers are listed in the Excel report.")
    st.caption(result.source.path.name)
    if result.database_source:
        st.caption(f"Database: {result.database_source.path.name} - {result.database_source.sheet}")
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
            st.dataframe([{"Row": number, "Field": name, "Your file": comparison_text(actual), "Database": comparison_text(expected)}
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
            for header, reason in result.unchecked_reasons.items():
                st.caption(f"{header}: {reason}.")
            if any(is_customer_name_column(item.file_column) or is_customer_name_column(item.database_column)
                   for item in result.mappings):
                st.caption("Question marks are ignored when comparing customer names. Original names are kept in the report.")
                st.caption("The replacement character (�) matches Ñ/ñ at the same position in a customer name; other letters must match.")
            if is_hsbc_mobile_pair(result.headers, result.database_headers, result.source.column, result.database_match_column):
                st.caption("Leading Mr./Ms., Mr. and Ms. titles are ignored in last names. The surname and suffixes are still checked; original values stay in the report.")
            if result.source.database_column:
                st.caption(f"All mapped fields are compared on the same database record using {result.source.column} → {result.database_match_column}.")
                if any(item.comparison == "Phone" for item in result.mappings):
                    st.caption("Phone values are checked against the detected database column, which may have an unexpected header. Review the field mappings below.")
            elif is_salad_ml_database(result.database_headers):
                st.caption("Identifiers, account number, financing number, customer name and email are checked on the same SALAD master list record. Staff assignments and last4 have no confirmed mapping and remain unchecked.")
            elif is_salad_endo_pair(result.headers, result.database_headers, result.source.column, result.database_match_column):
                st.caption("Account number, customer name and email are checked on the same SBF endorsement record (SALAD / PL). Agentcode and Collector staff assignments are not compared.")
            elif is_datagrid_pl_viber_blast(result.headers) and result.source.column.casefold() == "debtorid":
                st.caption("Debtorid is matched to DataGrid Debtor ID; FullName and Phone must match Name and Last Called No. on the same database row. Staff and campaign fields are not checked.")
            elif is_datagrid_balance_viber_blast(result.headers) and result.source.column.casefold() == "phone":
                if "level" in balance_viber_database_fields(result.database_headers):
                    st.caption("Phone is matched to DataGrid Phone; FullName, OB and Level must match Name, OB and Level on the same database row. Monthly, Terms and Promo have no confirmed calculation rule and remain unchecked.")
                else:
                    st.caption("Phone is matched to DataGrid Last Called No.; FullName and OB must match Name and Balance on the same database row. Level, Monthly, Terms and Promo have no confirmed database mapping or calculation rule and remain unchecked.")
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
    multiple_databases = len(run.databases) > 1
    st.markdown('<style>.st-key-check_preparation {display: none;}</style>', unsafe_allow_html=True)
    heading, screenshot, download = st.columns([2, 1, 1], vertical_alignment="center")
    with heading:
        st.subheader("Check results")
    with screenshot:
        screenshot_view = st.toggle("Screenshot view", key="screenshot_view",
                                    help="Hide uploads and file details for a compact screenshot. Turn off to return to the full view.")
    if screenshot_view:
        st.markdown("""<style>
        .st-key-check_setup, .st-key-result_details, .st-key-result_divider {display: none !important;}
        [data-testid="stHeader"] {display: none !important;}
        .stMainBlockContainer {max-width: none !important; padding: .5rem 1rem 1rem !important;}
        .stMainBlockContainer > [data-testid="stVerticalBlock"] {gap: .6rem !important;}
        </style>""", unsafe_allow_html=True)
    with download:
        st.download_button("Download Excel report", data=run.report, file_name="CHCODE_check_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="download_report", type="secondary", on_click="ignore", width="stretch")
    columns = st.columns(3)
    labels = (["Database comparisons", "Rows matching all", "Comparisons to review"] if multiple_databases
              else ["Files checked", "Rows matched", "Files to review"])
    for column, label, value in zip(columns, labels,
                                    [len(results), run.matched_in_all if multiple_databases else matched, review_count]):
        column.metric(label, f"{value:,}")
    st.markdown(result_overview_html(run), unsafe_allow_html=True)
    for note in dict.fromkeys(note for result in results for note in result.validation_notes):
        if "scientific notation" in note:
            st.warning(note)
        else:
            st.caption(note)
    with st.container(key="result_details"):
        render_result_details(run)


def render_result_details(run):
    multiple_databases = len(run.databases) > 1
    results = run.results
    for index, result in enumerate(results):
        status = ("Mapped fields match" if result.verdict == "MAPPED ROWS MATCH" else
                  "Exact match" if result.verdict in PASSED else "File error" if result.error else "Review needed")
        stage = result.file_identifier
        label = f"{stage} · {status}"
        if multiple_databases:
            label = f"{result.database_source.path.name} · {label}"
        if not result.error:
            label += f" · {result.count('MATCHED'):,}/{len(result.records):,} rows matched"
        with st.expander(label, expanded=False, icon=":material/check_circle:" if result.verdict in PASSED else ":material/error_outline:"):
            render_file_details(result, run, index)


def render_preparation(prepared, database_upload):
    if prepared and prepared.database_error:
        reset_results()
        st.error(f"Database: {database_upload.name}: {prepared.database_error}")
    elif prepared and prepared.database_info:
        info = prepared.database_info
        worksheet = f"{info.sheet}, " if info.sheet else ""
        st.caption(f"Database: {database_upload.name} - {worksheet}{info.column} header found on row {info.header_row}. Header rows are excluded from account checks.")

    if prepared and prepared.files:
        summary = []
        for item in prepared.files:
            info = item.inspection
            balance_fields = (balance_viber_database_fields(prepared.database_info.headers)
                              if info and prepared.database_info and is_datagrid_balance_viber_blast(info.headers)
                              else {})
            summary.append({
                "Database": database_upload.name,
                "File": item.source.upload.name,
                "Worksheet": item.source.sheet,
                "Header row": (str(info.header_row) if info and item.source.has_headers
                               else "No headers" if info else ""),
                "Lookup column": item.source.column if info else "",
                "Database lookup column": (item.source.database_column
                                           or (find_header(prepared.database_info.headers, balance_fields["phone"][0])
                                               if balance_fields
                                               else prepared.database.column)
                                           if info and prepared.database else ""),
                "Check": ("Detected field mappings" if item.inferred
                          else "CIMB account, email, name and product" if prepared.database_info and is_cimb_database(prepared.database_info.headers)
                          else "SALAD master list account details" if prepared.database_info and is_salad_ml_database(prepared.database_info.headers)
                          else "SALAD email account details" if info and is_datagrid_salad_email_blast(info.headers)
                          else "Account number, name and email" if info and prepared.database_info
                          and is_salad_endo_pair(info.headers, prepared.database_info.headers, item.source.column, prepared.database.column)
                          else "Viber debtor ID, phone and name" if info and is_datagrid_pl_viber_blast(info.headers)
                          else "Viber phone, name and balance" if balance_fields and "level" not in balance_fields
                          else "Viber phone, name, balance and level" if info and is_datagrid_balance_viber_blast(info.headers)
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
            if item.mapping_evidence:
                with st.expander(f"Detected field mappings: {item.source.upload.name}"):
                    st.caption("These mappings use headers and matching values from the uploaded database. Results cover the mapped fields; other columns are not checked.")
                    st.dataframe(list(item.mapping_evidence), hide_index=True, width="stretch")
        file_errors = sum(bool(item.error) for item in prepared.files)
        if file_errors:
            st.warning(f"{file_errors} file(s) need attention. Details are shown above and included in the report.")
        if not prepared.database_error:
            ready_count = len(prepared.files) - file_errors
            st.caption(f"{ready_count} file(s) ready to check.")


def main():
    st.set_page_config(page_title="File Checker", page_icon=":material/fact_check:", layout="wide", initial_sidebar_state="collapsed")
    apply_style()
    with st.container(key="check_setup"):
        title, restart = st.columns([4, 1], vertical_alignment="center")
        with title:
            st.title("File checker")
        with restart:
            st.button("Start over", on_click=clear_session, key="clear_session", type="tertiary", width="stretch")
        epoch = st.session_state.get("upload_epoch", 0)
        left, right = st.columns(2)
        with left:
            with st.container(border=True, key="upload_database_card"):
                st.subheader("1. Add your databases")
                database_files = st.file_uploader("Database files", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
                                                 key=f"database_upload_{epoch}", label_visibility="collapsed")
        with right:
            with st.container(border=True, key="upload_files_card"):
                st.subheader("2. Add files to check")
                uploaded_files = st.file_uploader("Files to check", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
                                                  key=f"files_upload_{epoch}", label_visibility="collapsed")

        # Accept a single file object as well as Streamlit's multiple-upload list.
        if database_files is None:
            database_files = []
        elif not isinstance(database_files, (list, tuple)):
            database_files = [database_files]
        database_uploads = [Upload(file.name, file.getvalue()) for file in database_files]
        file_uploads = [Upload(file.name, file.getvalue()) for file in uploaded_files]
        current_inputs = (AUTO_DETECTION_VERSION, tuple(upload.fingerprint for upload in database_uploads),
                          tuple(upload.fingerprint for upload in file_uploads))
        if st.session_state.get("input_signature") != current_inputs:
            reset_results()
            st.session_state.pop("preparation", None)
            st.session_state.input_signature = current_inputs

        preparations = []
        if database_uploads:
            if "preparation" not in st.session_state:
                with st.spinner("Detecting worksheets, headers, and lookup columns..."):
                    st.session_state.preparation = [prepare_uploaded_check(database, file_uploads)
                                                    for database in database_uploads]
            preparations = st.session_state.preparation

        ready = bool(preparations and file_uploads and all(
            item.database and not item.database_error for item in preparations))
        with st.container(key="check_preparation"):
            for database_upload, prepared in zip(database_uploads, preparations):
                render_preparation(prepared, database_upload)
            if not database_uploads:
                st.caption("Add a database and your files to get started. Excel and CSV are supported.")
            elif preparations and not file_uploads:
                st.caption("Add at least one file to check.")

        options = Options(ignore_case=True, trim_spaces=True, mapped_fields=True,
                          require_flow_stage=True, allow_missing_flow_stage=True)
        signature = tuple(configuration_signature(check.database, [item.source for item in check.files], options)
                          for check in preparations) if ready else None
        if st.session_state.get("check_run") and st.session_state.get("check_signature") != signature:
            reset_results()
        if st.button("3. Check files", type="primary", key="run_check", disabled=not ready, width="stretch"):
            reset_results()
            progress = st.progress(0, text="Reading your database...")

            def update(index, total, result):
                progress.progress(index / total, text=f"Checking file {index} of {total}...")

            try:
                with st.spinner("Checking your files..."):
                    st.session_state.check_run = (run_prepared_check(preparations[0], update) if len(preparations) == 1
                                                  else run_prepared_checks(preparations, update))
                    st.session_state.check_signature = signature
            except Exception as error:
                st.error(f"Unable to complete the check: {error}")
            finally:
                progress.empty()
    if st.session_state.get("check_run"):
        with st.container(key="result_divider"):
            st.divider()
        render_results(st.session_state.check_run)


if __name__ == "__main__":
    main()

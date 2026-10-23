"""Run with: python -m streamlit run streamlit_app.py"""

from __future__ import annotations

import hashlib
from collections import defaultdict
from dataclasses import replace

import streamlit as st

from checker import ColumnMapping, Options, find_header, is_datagrid_pair, suggest_mappings
from web_checker import Upload, UploadedSource, configuration_signature, inspect_upload, run_uploaded_check

UPLOADED_ROWS = "Validate uploaded rows (mapped fields)"
FULL_COPY = "Compare complete database copy"
SKIP = "(Do not check)"
PREVIEW_LIMIT = 500
PASSED = {"EXACT MATCH", "MAPPED ROWS MATCH", "SAME ROWS"}


def clear_session():
    epoch = st.session_state.get("upload_epoch", 0) + 1
    st.session_state.clear()
    st.session_state.upload_epoch = epoch


def inspection(upload, sheet=None, *, has_headers=True):
    cache = st.session_state.setdefault("inspections", {})
    key = (upload.fingerprint, sheet, has_headers)
    if key not in cache:
        cache[key] = inspect_upload(upload, sheet, has_headers=has_headers)
    return cache[key]


def text(value):
    return "" if value is None else str(value)


def reset_results():
    st.session_state.pop("check_run", None)
    st.session_state.pop("check_signature", None)


def lookup_columns(info, *, database=False):
    names = {"chcode", "old i.c."} if database else {"chcode", "indexone_uid", "primary_key"}
    return [header for header in info.headers if header.casefold() in names or header == info.column]


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


def mapping_editor(headers, file_key, database_info, layout_id, count, *, has_headers=True):
    suggested = {item.file_column: item for item in suggest_mappings(
        list(headers), list(database_info.headers), file_key, database_info.column)} if has_headers else {}
    identifier = find_header(headers, "FLOW_STAGE")
    rows = []
    for header in headers:
        if header in {file_key, identifier}:
            continue
        item = suggested.get(header)
        rows.append({"File column": header, "Database column": item.database_column if item else SKIP,
                     "Compare as": item.comparison if item else "Text"})
    with st.popover(f"Edit comparison fields · {count} file(s)", width="stretch"):
        if not has_headers:
            st.caption(f"CHCODE in column B finds the database row via {database_info.column}. Other columns are unchecked until mapped.")
        elif not identifier and is_datagrid_pair(headers, database_info.headers, file_key, database_info.column):
            st.caption(f"{file_key} finds the database row via {database_info.column}. The filename identifies this DataGrid email blast.")
        else:
            st.caption(f"{file_key} finds the database row via {database_info.column}. FLOW_STAGE identifies the file.")
        edited = st.data_editor(
            rows, hide_index=True, num_rows="fixed", disabled=["File column"],
            column_config={
                "Database column": st.column_config.SelectboxColumn(options=[SKIP, *database_info.headers], required=True),
                "Compare as": st.column_config.SelectboxColumn(options=["Text", "Number", "Product"], required=True),
            }, key=f"mapping_{layout_id}", width="stretch",
        ) if rows else []
        mappings = [ColumnMapping(file_key, database_info.column)]
        for row in edited:
            if row["Database column"] != SKIP:
                mappings.append(ColumnMapping(row["File column"], row["Database column"], row["Compare as"]))
        checked = {item.file_column for item in mappings}
        ignored = [name for name in headers if name not in checked and name != identifier]
        st.caption("Number ignores comma separators. Product recognizes Credit Card / CARDS and Personal Loan / LOAN.")
        st.text(f"Not checked: {', '.join(ignored) or 'None'}")
    return tuple(mappings)


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
            st.caption("Only CHCODE in column B is checked by default for headerless files. Unmapped columns are not checked.")
        else:
            st.caption("Results cover the compared fields. Unmapped campaign fields are not checked.")
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
    st.set_page_config(page_title="File Checker", page_icon="📋", layout="wide", initial_sidebar_state="collapsed")
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
    current_inputs = (database_upload.fingerprint if database_upload else None, tuple(upload.fingerprint for upload in file_uploads))
    if st.session_state.get("input_signature") != current_inputs:
        reset_results()
        st.session_state.input_signature = current_inputs
        active = {upload.fingerprint for upload in file_uploads}
        if database_upload:
            active.add(database_upload.fingerprint)
        st.session_state.inspections = {key: value for key, value in st.session_state.get("inspections", {}).items() if key[0] in active}

    database_info = None
    database_sheet = ""
    database_error = ""
    sources = []
    layouts = defaultdict(list)
    file_summary = []
    with st.expander("Advanced options", expanded=False):
        mode = st.selectbox("Comparison", [UPLOADED_ROWS, FULL_COPY], key="check_mode")
        option_left, option_right = st.columns(2)
        ignore_case = option_left.checkbox("Ignore letter case", value=True, key="ignore_case")
        trim_spaces = option_right.checkbox("Trim outer spaces", value=True, key="trim_spaces")
        if mode == FULL_COPY:
            st.caption("Every database row and data column must be present. FLOW_STAGE remains the file identifier.")
        else:
            st.caption("Checks the rows in your files. They can be smaller subsets of the database.")
        if database_upload:
            try:
                base = inspection(database_upload)
                database_sheet = st.selectbox("Database worksheet", base.sheets,
                                               key=f"database_sheet_{database_upload.fingerprint}") if len(base.sheets) > 1 else base.sheet
                database_info = inspection(database_upload, database_sheet)
                if not database_info.column:
                    database_error = "The selected database worksheet must have a CHCODE header or the supported DataGrid headers: Name, Old I.C., Account No., and Email."
                else:
                    columns = lookup_columns(database_info, database=True)
                    database_column = st.selectbox(
                        "Database lookup column (CHCODE)", columns, index=columns.index(database_info.column),
                        key=f"database_column_{database_upload.fingerprint}_{database_sheet}",
                        help="Select the database column containing the CHCODE. DataGrid exports usually use Old I.C.",
                    )
                    database_info = replace(database_info, column=database_column)
            except Exception as error:
                database_error = str(error)
        for index, upload in enumerate(file_uploads):
            has_headers = not st.checkbox(f"No headers (CHCODE in column B): {upload.name}",
                                           key=f"headerless_{index}_{upload.fingerprint}",
                                           help="Leave unchecked when the file has headers, including DataGrid email blasts. Select only when every row is data and CHCODE values are in column B.")
            try:
                base = inspection(upload, has_headers=has_headers)
                sheet = st.selectbox(f"Worksheet: {upload.name}", base.sheets,
                                     key=f"file_sheet_{index}_{upload.fingerprint}") if len(base.sheets) > 1 else base.sheet
                info = inspection(upload, sheet, has_headers=has_headers)
                if has_headers and info.column:
                    columns = lookup_columns(info)
                    file_column = st.selectbox(
                        f"Lookup column (CHCODE): {upload.name}", columns, index=columns.index(info.column),
                        key=f"file_column_{index}_{upload.fingerprint}_{sheet}",
                        help="Select the file column containing the CHCODE. Choose indexone_uid or primary_key according to this file's layout.",
                    )
                    info = replace(info, column=file_column)
                source = UploadedSource(upload, sheet, info.column or "CHCODE", has_headers=has_headers)
                if not info.column:
                    status = "Missing CHCODE or supported DataGrid email-blast key header"
                elif not has_headers and mode == FULL_COPY:
                    status = "Headerless files need uploaded-row comparison"
                elif (has_headers and not find_header(info.headers, "FLOW_STAGE")
                      and not (mode == UPLOADED_ROWS and database_info and is_datagrid_pair(
                          info.headers, database_info.headers, info.column, database_info.column))):
                    status = "Missing FLOW_STAGE header"
                else:
                    status = "Ready"
                    layouts[(info.headers, info.column, has_headers)].append(index)
                file_summary.append({"File": upload.name, "Worksheet": sheet,
                                     "Header row": str(info.header_row) if info.header_row is not None else "No headers",
                                     "Status": status})
            except Exception as error:
                source = UploadedSource(upload, "", "CHCODE", has_headers=has_headers)
                file_summary.append({"File": upload.name, "Worksheet": "", "Header row": "", "Status": str(error)})
            sources.append(source)
        if file_summary:
            st.dataframe(file_summary, hide_index=True, width="stretch")
        mapped_fields = mode == UPLOADED_ROWS
        if mapped_fields and database_info and not database_error:
            for (headers, file_key, has_headers), indices in layouts.items():
                layout_id = hashlib.sha256(repr((database_upload.fingerprint, database_sheet, database_info.column, headers, file_key, has_headers)).encode("utf-8")).hexdigest()
                mappings = mapping_editor(headers, file_key, database_info, layout_id, len(indices), has_headers=has_headers)
                for index in indices:
                    sources[index] = replace(sources[index], mappings=mappings)
        st.caption("CHCODE campaign files with headers need one FLOW_STAGE value, such as NO_REPLY. Recognized DataGrid email blasts use their filenames in uploaded-row mode. Headerless files use column B as CHCODE and include row 1.")
        st.caption("CHCODE and supported DataGrid header rows are detected automatically; rows above them are skipped. Leave No headers unchecked when the file has headers.")
        st.caption("Check the lookup columns above. For DataGrid, use Old I.C. in the database and choose indexone_uid or primary_key in each email blast according to where its CHCODE is stored.")
        st.caption("Headerless files check CHCODE only by default. Edit comparison fields to map other columns. The database still needs headers.")
        st.caption("Uploads are processed on the app server in this session. Start over clears the current uploads and results.")

    if database_info and not database_error:
        st.caption(f"Database: {database_info.column} header found on row {database_info.header_row}.")
    if database_error:
        reset_results()
        st.error(f"Database: {database_error}")
    file_errors = sum(item["Status"] != "Ready" for item in file_summary)
    if file_errors:
        st.warning(f"{file_errors} file(s) need attention. See Advanced options for details.")
    ready = bool(database_upload and database_info and not database_error and sources)
    if ready:
        st.caption(f"{len(sources)} file(s) ready to check.")
    elif database_upload is None:
        reset_results()
        st.caption("Add a database and your files to get started. Excel and CSV are supported.")
    else:
        reset_results()
        st.caption("Add at least one file to check.")

    options = Options(ignore_case, trim_spaces, mapped_fields, require_flow_stage=True)
    database = UploadedSource(database_upload, database_sheet, database_info.column) if ready else None
    signature = configuration_signature(database, sources, options) if ready else None
    if st.session_state.get("check_run") and st.session_state.get("check_signature") != signature:
        reset_results()
        st.caption("Settings changed. Check the files again to update results.")
    if st.button("3. Check files", type="primary", key="run_check", disabled=not ready, width="stretch"):
        reset_results()
        progress = st.progress(0, text="Reading your database…")

        def update(index, total, result):
            progress.progress(index / total, text=f"Checking file {index} of {total}…")

        try:
            with st.spinner("Checking your files…"):
                st.session_state.check_run = run_uploaded_check(database, sources, options, update)
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

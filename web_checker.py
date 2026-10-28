"""Session-friendly upload adapter for the existing file comparison engine.

No Streamlit dependency here: each check gets its own temporary directory, and
the returned result uses uploaded filenames rather than temporary server paths.
"""

from __future__ import annotations

import hashlib
import json
import tempfile
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import Callable

from checker import (
    CHCODE_ALIASES, ColumnMapping, FileResult, Options, Source, Table, check_files, chcode_columns, export_report,
    find_header, is_datagrid_pl_viber_blast, is_datagrid_viber_blast, is_salad_endo_database, is_salad_endo_pair, match_column, normalize, phone_column, read_table, sheet_names,
)

AUTO_DETECTION_VERSION = 12


@dataclass(frozen=True)
class Upload:
    name: str
    data: bytes

    @property
    def fingerprint(self) -> str:
        return hashlib.sha256(self.name.encode("utf-8") + b"\0" + self.data).hexdigest()


@dataclass(frozen=True)
class Inspection:
    sheets: tuple[str, ...]
    sheet: str
    headers: tuple[str, ...]
    column: str
    header_row: int | None = 1
    has_headers: bool = True


@dataclass(frozen=True)
class UploadedSource:
    upload: Upload
    sheet: str
    column: str
    mappings: tuple[ColumnMapping, ...] | None = None
    has_headers: bool = True
    detection_error: str = ""


@dataclass(frozen=True)
class PreparedUpload:
    source: UploadedSource
    inspection: Inspection | None = None
    error: str = ""
    note: str = ""


@dataclass(frozen=True)
class PreparedCheck:
    database: UploadedSource | None
    database_info: Inspection | None
    files: tuple[PreparedUpload, ...] = ()
    database_error: str = ""


@dataclass
class CheckRun:
    database: Source
    options: Options
    results: list[FileResult]
    report: bytes


def clean_filename(name: str) -> str:
    """Remove client paths on both Windows and Linux, without losing extensions."""
    result = name.replace("\\", "/").rsplit("/", 1)[-1]
    if result in {"", ".", ".."} or "\0" in result:
        raise ValueError("The uploaded file has an invalid filename.")
    return result


def _write_upload(directory: Path, upload: Upload) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / clean_filename(upload.name)
    path.write_bytes(upload.data)
    return path


def inspect_upload(upload: Upload, sheet: str | None = None, *, has_headers: bool = True) -> Inspection:
    with tempfile.TemporaryDirectory(prefix="chcode_inspect_") as folder:
        path = _write_upload(Path(folder), upload)
        names = sheet_names(path)
        table = read_table(path, sheet or names[0], headers_only=True, has_headers=has_headers)
        column = match_column(table.headers)
        return Inspection(tuple(names), table.sheet, tuple(table.headers), column, table.header_row, has_headers)


def _key_column(table: Table, database_keys: set[str], database_headers=()) -> tuple[str, str]:
    """Choose one lookup for the entire file; unmatched rows never switch keys."""
    if is_salad_endo_pair(table.headers, database_headers):
        return match_column(table.headers, database_headers), "SBF endorsement layout detected (SALAD / PL); accountnumber matches caseAccountNo; email and chname are checked on the same account. Staff assignments are not checked"
    if (is_salad_endo_database(database_headers) and not find_header(database_headers, "CHCODE")
            and not find_header(table.headers, "CHCODE")):
        raise ValueError("SBF endorsement email files require accountnumber, chname, and email columns.")
    if is_datagrid_pl_viber_blast(table.headers):
        return match_column(table.headers), "PL Viber layout detected; Debtorid matches DataGrid Debtor ID; FullName and Phone are checked against Name and Last Called No. on the same row"
    if is_datagrid_viber_blast(table.headers):
        return match_column(table.headers), "SALAD Viber layout detected; Phone matches DataGrid Last Called No. and FullName is checked against Name on the same row"
    chcode = find_header(table.headers, "CHCODE")
    if chcode:
        return chcode, "CHCODE header detected"
    candidates = chcode_columns(table.headers)
    if not candidates:
        return "", ""
    values = {header: [normalize(row[table.headers.index(header)], Options()) for _, row in table.rows]
              for header in candidates}
    scores = {header: sum(bool(value) and value in database_keys for value in values[header])
              for header in candidates}
    best = max(scores.values())
    tied = [header for header in candidates if scores[header] == best]
    preferred = match_column(table.headers)
    # Identical columns are interchangeable for lookup. Prefer primary_key,
    # including the SBF layout, while still comparing both identifier fields.
    if len(tied) == 2 and best > 0 and values[tied[0]] == values[tied[1]]:
        selected = tied[0]
    else:
        selected = preferred if preferred in tied else tied[0]
    return selected, f"{selected}: {scores[selected]:,} of {len(table.rows):,} rows found in the database"


def _detect_file_table(path: Path, sheet: str, database: Table,
                       database_keys: set[str]) -> tuple[Table, bool]:
    """Inspect both readings so an actual first data row is never discarded."""
    headered = raw = None
    header_error = raw_error = ""
    try:
        headered = read_table(path, sheet)
    except ValueError as error:
        header_error = str(error)
    try:
        raw = read_table(path, sheet, has_headers=False)
    except ValueError as error:
        raw_error = str(error)

    known_headers = {name.casefold() for name in database.headers} | set(CHCODE_ALIASES) | {
        "flow_stage", "email", "cc", "bcc",
        "name", "chname", "accountnumber", "financingnumber", "formatted mobile",
        "lastname", "producttype", "product_type", "chname_format", "ob_at_endorsement",
    }
    if headered is not None:
        key = match_column(headered.headers, database.headers)
        recognized = sum(header.casefold() in known_headers or bool(phone_column([header]))
                         for header in headered.headers)
        first_key_matches = bool(raw and raw.rows
                                 and normalize(raw.rows[0][1][1], Options()) in database_keys)
        if key:
            earlier_data = bool(raw and any(
                number < headered.header_row
                and (normalize(values[1], Options()) in database_keys
                     or (recognized < 2 and bool(normalize(values[1], Options()))))
                for number, values in raw.rows))
            first_row_is_data = bool(raw and raw.rows and recognized < 2 and (
                first_key_matches
                or (headered.headers.index(key) == 1 and key.casefold() == "chcode")))
            if not earlier_data and not first_row_is_data:
                return headered, True
        elif (recognized >= 2 or phone_column(headered.headers)) and not first_key_matches:
            message = "The file has headers but no CHCODE, primary_key, or indexone_uid column."
            if phone_column(headered.headers):
                message += " CHCODE is required to compare phone numbers for the same account."
            raise ValueError(message)
    if raw is not None:
        # A malformed named header must not become data and bypass validation.
        first = next((values for _, values in raw.rows
                      if any(value is not None and str(value).strip() for value in values[1:])), ())
        named = {str(value).strip().casefold() for value in first if value is not None}
        if sum(name in known_headers or bool(phone_column([name])) for name in named) >= 2 and not (
                len(first) > 1 and normalize(first[1], Options()) in database_keys):
            raise ValueError(header_error or "The file has headers but no supported CHCODE column.")
        return raw, False
    raise ValueError(header_error or raw_error or "The selected worksheet is empty.")


def prepare_uploaded_check(database: Upload, files: list[Upload]) -> PreparedCheck:
    """Detect upload layouts against one database without retaining row data."""
    with tempfile.TemporaryDirectory(prefix="chcode_detect_") as folder:
        root = Path(folder)
        try:
            database_path = _write_upload(root / "database", database)
            database_sheets = tuple(sheet_names(database_path))
            candidates = []
            database_errors = []
            for sheet in database_sheets:
                try:
                    table = read_table(database_path, sheet)
                    column = match_column(table.headers)
                    if column:
                        candidates.append((table, column))
                except ValueError as error:
                    database_errors.append(str(error))
            if not candidates:
                raise ValueError(database_errors[0] if database_errors else
                                 "The database must have a CHCODE header (CHCODE, primary_key, or indexone_uid), supported DataGrid headers, or the SBF SALAD / PL endorsement columns: caseNo, caseAccountNo, caseCustomerName, caseCustomerEmail, and Collector.")
            if len(candidates) > 1:
                raise ValueError("Multiple database worksheets have supported headers. Upload a workbook with one database worksheet.")
            table, database_column = candidates[0]
            key_position = table.headers.index(database_column)
            database_keys = {normalize(values[key_position], Options()) for _, values in table.rows}
            if not database_keys or database_keys == {""}:
                raise ValueError("The database match column has no nonblank values.")
            if "" in database_keys:
                raise ValueError(f"The database has blank values in '{database_column}'. Fill these before comparing rows.")
            database_info = Inspection(database_sheets, table.sheet, tuple(table.headers),
                                       database_column, table.header_row)
            database_source = UploadedSource(database, table.sheet, database_column)
        except Exception as error:
            return PreparedCheck(None, None, database_error=str(error))

        prepared = []
        for index, upload in enumerate(files):
            try:
                path = _write_upload(root / "files" / str(index), upload)
                names = tuple(sheet_names(path))
                candidates = []
                errors = []
                for sheet in names:
                    try:
                        target, has_headers = _detect_file_table(path, sheet, table, database_keys)
                        candidates.append((target, has_headers))
                    except ValueError as error:
                        errors.append(str(error))
                if not candidates:
                    raise ValueError(errors[0] if errors else "No file data was found.")
                if len(candidates) > 1:
                    raise ValueError("Multiple worksheets contain file data. Upload a workbook with one data worksheet.")
                target, has_headers = candidates[0]
                column, note = _key_column(target, database_keys, table.headers) if has_headers else (
                    "CHCODE", "No headers detected; column B is CHCODE and row 1 is included")
                if (has_headers and phone_column(target.headers)
                        and not is_datagrid_viber_blast(target.headers) and not is_datagrid_pl_viber_blast(target.headers)
                        and not find_header(target.headers, "FLOW_STAGE")):
                    note += "; no FLOW_STAGE detected; phone numbers are compared for the same CHCODE"
                info = Inspection(names, target.sheet, tuple(target.headers), column,
                                  target.header_row, has_headers)
                source = UploadedSource(upload, target.sheet, column, has_headers=has_headers)
                prepared.append(PreparedUpload(source, info, note=note))
            except Exception as error:
                message = str(error)
                source = UploadedSource(upload, "", "CHCODE", detection_error=message)
                prepared.append(PreparedUpload(source, error=message))
        return PreparedCheck(database_source, database_info, tuple(prepared))


def run_prepared_check(prepared: PreparedCheck,
                       progress: Callable[[int, int, FileResult], None] | None = None) -> CheckRun:
    if prepared.database is None:
        raise ValueError(prepared.database_error or "Upload a valid database first.")
    return run_uploaded_check(prepared.database, [item.source for item in prepared.files],
                              Options(mapped_fields=True, require_flow_stage=True,
                                      allow_missing_flow_stage=True), progress)


def configuration_signature(database: UploadedSource, files: list[UploadedSource], options: Options) -> str:
    def describe(source):
        return {"upload": source.upload.fingerprint, "sheet": source.sheet, "column": source.column,
                "mappings": None if source.mappings is None else [asdict(item) for item in source.mappings],
                "has_headers": source.has_headers, "detection_error": source.detection_error}
    configuration = {"database": describe(database), "files": [describe(source) for source in files], "options": asdict(options)}
    return hashlib.sha256(json.dumps(configuration, sort_keys=True).encode("utf-8")).hexdigest()


def run_uploaded_check(database: UploadedSource, files: list[UploadedSource], options: Options,
                       progress: Callable[[int, int, FileResult], None] | None = None) -> CheckRun:
    if not files:
        raise ValueError("Upload at least one file to check.")
    with tempfile.TemporaryDirectory(prefix="chcode_check_") as folder:
        root = Path(folder)
        database_source = Source(_write_upload(root / "database", database.upload), database.sheet, database.column,
                                 has_headers=database.has_headers)
        sources = [Source(_write_upload(root / "files" / str(index), source.upload), source.sheet, source.column,
                          source.mappings, source.has_headers, source.detection_error)
                   for index, source in enumerate(files)]
        results = check_files(database_source, sources, options, progress)
        # The download describes client filenames, and remains valid after the
        # temporary source files have been deleted.
        displayed_database = replace(database_source, path=Path(clean_filename(database.upload.name)))
        for result, source in zip(results, files):
            result.source = replace(result.source, path=Path(clean_filename(source.upload.name)))
        report_path = root / "report" / "CHCODE_check_report.xlsx"
        report_path.parent.mkdir()
        export_report(report_path, displayed_database, options, results)
        return CheckRun(displayed_database, options, results, report_path.read_bytes())

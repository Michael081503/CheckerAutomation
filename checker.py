"""File readers and matching engine for the local database checker."""

from __future__ import annotations

import csv
from collections import defaultdict, deque
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from itertools import chain
from pathlib import Path
from typing import Callable, Iterable

from openpyxl import Workbook, load_workbook

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm"}


@dataclass
class Table:
    path: Path
    sheet: str
    headers: list[str]
    rows: list[tuple[int, tuple]]
    header_row: int | None = 1


@dataclass(frozen=True)
class ColumnMapping:
    file_column: str
    database_column: str
    comparison: str = "Text"


@dataclass(frozen=True)
class Source:
    path: Path
    sheet: str
    column: str
    mappings: tuple[ColumnMapping, ...] | None = None
    has_headers: bool = True


@dataclass(frozen=True)
class Options:
    ignore_case: bool = True
    trim_spaces: bool = True
    mapped_fields: bool = False
    require_flow_stage: bool = False


@dataclass
class FileResult:
    source: Source
    headers: list[str]
    # Source row, status, matching database rows, original source values.
    records: list[tuple[int, str, str, tuple]]
    error: str = ""
    database_headers: list[str] = field(default_factory=list)
    missing_rows: list[tuple[int, tuple]] = field(default_factory=list)
    differences: list[tuple[int | None, int | None, str, object, object]] = field(default_factory=list)
    missing_columns: list[str] = field(default_factory=list)
    extra_columns: list[str] = field(default_factory=list)
    comparison_mode: str = "Full database copy"
    mappings: list[ColumnMapping] = field(default_factory=list)
    ignored_columns: list[str] = field(default_factory=list)
    flow_stage_required: bool = False
    flow_stage_column: str = ""
    flow_stage: str = ""

    def count(self, status: str) -> int:
        if status == "MISSING FROM FILE":
            return len(self.missing_rows)
        return sum(record[1] == status for record in self.records)

    @property
    def verdict(self) -> str:
        if self.error:
            return "ERROR"
        if self.comparison_mode == "Uploaded rows":
            if self.records and all(row[1] == "MATCHED" for row in self.records):
                return "EXACT MATCH" if self.flow_stage_required else "MAPPED ROWS MATCH"
            return "ROW ISSUES"
        if (self.missing_rows or self.missing_columns or self.extra_columns
                or any(record[1] != "MATCHED" for record in self.records)):
            return "DIFFERENT ROWS"
        return "SAME ROWS"


class Cancelled(Exception):
    pass


def _validate(path: Path) -> None:
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError("Choose an .xlsx, .xlsm, or .csv file. Convert older .xls files to .xlsx first.")
    if not path.is_file():
        raise ValueError(f"File no longer exists: {path.name}")


def sheet_names(path: Path) -> list[str]:
    _validate(path)
    if path.suffix.lower() == ".csv":
        return ["CSV"]
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        return book.sheetnames
    finally:
        book.close()


def _has_chcode(values: Iterable) -> bool:
    return any(str(value).strip().casefold() == "chcode" for value in values if value is not None)


def _table(path: Path, sheet: str, rows: Iterable, *, has_headers: bool = True,
           headers_only: bool = False) -> Table:
    iterator = iter(rows)
    first = next(iterator, None)
    if first is None:
        raise ValueError(f"{path.name}: the selected sheet is empty.")
    if not has_headers:
        records = []
        width = 0
        for number, values in enumerate(chain((first,), iterator), 1):
            values = tuple(values)
            if any(value is not None and str(value).strip() for value in values):
                width = max(width, len(values))
                records.append((number, values))
        if not records:
            raise ValueError(f"{path.name}: the selected sheet is empty.")
        if width < 2:
            raise ValueError(f"{path.name}: headerless files need at least two columns; CHCODE must be in column B.")
        from openpyxl.utils import get_column_letter
        headers = ["CHCODE" if index == 2 else f"Column {get_column_letter(index)}"
                   for index in range(1, width + 1)]
        return Table(path, sheet, headers,
                     [] if headers_only else
                     [(number, values + (None,) * (width - len(values))) for number, values in records],
                     header_row=None)
    header_row = 1
    if not _has_chcode(first):
        preamble = []
        for number, values in enumerate(iterator, 2):
            if _has_chcode(values):
                first = values
                header_row = number
                break
            preamble.append(values)
        else:
            # Keep the generic reader's missing-key behavior when CHCODE is absent.
            iterator = iter(preamble)
    # Keep internal empty headers so column positions never shift.
    headers = [str(value).strip() if value is not None else "" for value in first]
    while headers and not headers[-1]:
        headers.pop()
    if not headers or any(not header for header in headers):
        raise ValueError(f"{path.name}: header row {header_row} must contain a name for every data column.")
    # Real database exports repeat names. Preserve every position and give later
    # occurrences visible suffixes; the first CHCODE remains the primary key.
    used = set()
    for index, name in enumerate(headers):
        candidate = name
        occurrence = 2
        while candidate.casefold() in used:
            candidate = f"{name} [{occurrence}]"
            occurrence += 1
        headers[index] = candidate
        used.add(candidate.casefold())
    records = []
    if headers_only:
        return Table(path, sheet, headers, records, header_row=header_row)
    for number, values in enumerate(iterator, header_row + 1):
        values = tuple(values)
        if any(value is not None and str(value).strip() for value in values):
            if any(value is not None and str(value).strip() for value in values[len(headers):]):
                raise ValueError(f"{path.name}: row {number} contains data under an unnamed column.")
            records.append((number, values[:len(headers)] + (None,) * max(0, len(headers) - len(values))))
    return Table(path, sheet, headers, records, header_row=header_row)


def read_table(path: Path, sheet: str | None = None, *, headers_only: bool = False,
               has_headers: bool = True) -> Table:
    path = Path(path)
    _validate(path)
    if path.suffix.lower() == ".csv":
        try:
            contents = path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            contents = path.read_text(encoding="cp1252")
        try:
            dialect = csv.Sniffer().sniff(contents[:65536], delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        import io
        delimiter = dialect.delimiter
        if has_headers:
            # Preambles with different row widths can defeat Sniffer. Recover
            # the delimiter using the actual CHCODE header, including later rows.
            candidates = [delimiter, *(value for value in ",;\t|" if value != delimiter)]
            for candidate in candidates:
                probe = csv.reader(io.StringIO(contents, newline=""), dialect, delimiter=candidate)
                if any(_has_chcode(values) for values in probe):
                    delimiter = candidate
                    break
        reader = csv.reader(io.StringIO(contents, newline=""), dialect, delimiter=delimiter)
        return _table(path, "CSV", reader, has_headers=has_headers, headers_only=headers_only)
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        selected = sheet or book.sheetnames[0]
        if selected not in book.sheetnames:
            raise ValueError(f"Sheet '{selected}' no longer exists in {path.name}.")
        reader = book[selected].iter_rows(values_only=True)
        return _table(path, selected, reader, has_headers=has_headers, headers_only=headers_only)
    finally:
        book.close()


def normalize(value, options: Options) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        result = str(value)
    elif isinstance(value, (int, float, Decimal)):
        result = format(Decimal(str(value)).normalize(), "f")
    elif isinstance(value, (datetime, date)):
        result = value.isoformat()
    else:
        result = str(value)
    if options.trim_spaces:
        result = result.strip()
    return result.casefold() if options.ignore_case else result


def suggest_mappings(file_headers: list[str], database_headers: list[str],
                     file_key: str, database_key: str) -> list[ColumnMapping]:
    """Direct header matches and explicit HSBC export aliases; never infer by position."""
    database_names = {name.casefold(): name for name in database_headers}
    aliases = {
        "email": ("Email address", "Text"),
        "primary_key": (database_key, "Text"),
        "chname": ("Name", "Text"),
        "chname_format": ("Last Name", "Text"),
        "product_type": ("PROD TYPE", "Product"),
        "ob_at_endorsement": ("OB", "Number"),
    }
    mappings = []
    for name in file_headers:
        if name == file_key:
            mappings.append(ColumnMapping(name, database_key))
        elif name.casefold() in database_names:
            mappings.append(ColumnMapping(name, database_names[name.casefold()]))
        elif name.casefold() in aliases:
            expected, comparison = aliases[name.casefold()]
            if expected.casefold() in database_names:
                mappings.append(ColumnMapping(name, database_names[expected.casefold()], comparison))
    return mappings


def compare_value(value, comparison: str, options: Options) -> str:
    if comparison == "Number" and value is not None and str(value).strip():
        try:
            from decimal import InvalidOperation
            number = Decimal(str(value).strip().replace(",", ""))
            if number.is_finite():
                return format(number, "f").rstrip("0").rstrip(".") if "." in format(number, "f") else format(number, "f")
        except InvalidOperation:
            pass
    if comparison == "Product":
        product = str(value or "").strip().casefold()
        family = {"credit card": "cards", "cards": "cards", "personal loan": "loan", "loan": "loan"}
        if product in family:
            return family[product]
    return normalize(value, options)


def find_header(headers: Iterable[str], name: str) -> str:
    return next((header for header in headers if header.casefold() == name.strip().casefold()), "")


def _identify_flow_stage(target: Table, result: FileResult) -> None:
    result.flow_stage_column = find_header(target.headers, "FLOW_STAGE")
    if not result.flow_stage_column:
        raise ValueError("The uploaded file must have a FLOW_STAGE header to identify the file type.")
    position = target.headers.index(result.flow_stage_column)
    stages = {}
    blank_rows = []
    for number, values in target.rows:
        value = "" if values[position] is None else str(values[position]).strip()
        if not value:
            blank_rows.append(number)
        else:
            stages.setdefault(value.casefold(), value.upper())
    if blank_rows:
        raise ValueError(f"FLOW_STAGE is blank at source row(s) {', '.join(map(str, blank_rows[:20]))}. Every row needs the file identifier.")
    if len(stages) > 1:
        raise ValueError(f"The file contains multiple FLOW_STAGE identifiers ({', '.join(stages.values())}). Upload a separate file for each stage.")
    if not stages:
        raise ValueError("The file has no data rows with a FLOW_STAGE identifier.")
    result.flow_stage = next(iter(stages.values()))


def _compare_mapped(table: Table, target: Table, database: Source, source: Source,
                    result: FileResult, options: Options, lookup: dict,
                    cancelled: Callable[[], bool]) -> None:
    result.comparison_mode = "Uploaded rows"
    if source.mappings is not None:
        mappings = list(source.mappings)
    elif source.has_headers:
        mappings = suggest_mappings(target.headers, table.headers, source.column, database.column)
    else:
        mappings = [ColumnMapping(source.column, database.column)]
    if result.flow_stage_required:
        if source.mappings is not None and any(item.file_column == result.flow_stage_column for item in mappings):
            raise ValueError("FLOW_STAGE is the file identifier and must not be mapped to a database column.")
        mappings = [item for item in mappings if item.file_column != result.flow_stage_column]
    if not any(item.file_column == source.column and item.database_column == database.column for item in mappings):
        raise ValueError("The CHCODE-to-CHCODE mapping is required.")
    if len({item.file_column for item in mappings}) != len(mappings):
        raise ValueError("Each file column can have only one database mapping.")
    for item in mappings:
        if item.file_column not in target.headers or item.database_column not in table.headers:
            raise ValueError(f"Mapping no longer exists: {item.file_column} → {item.database_column}. Review field mappings.")
        if item.comparison not in {"Text", "Number", "Product"}:
            raise ValueError(f"Unknown comparison type: {item.comparison}")
    result.mappings = mappings
    result.ignored_columns = [name for name in target.headers if name not in {item.file_column for item in mappings}
                              and name != result.flow_stage_column]
    if not target.rows:
        raise ValueError("The selected file has column headers but no data rows to validate.")
    positions = [(target.headers.index(item.file_column), table.headers.index(item.database_column), item) for item in mappings]
    database_rows = dict(table.rows)
    # An uploaded row is validated independently; database rows may be reused by
    # repeated uploaded CHCODEs and by separate campaign files.
    key_position = target.headers.index(source.column)
    database_signatures = {}
    exact = {}
    for key, numbers in lookup.items():
        if cancelled():
            raise Cancelled()
        for number in numbers:
            values = database_rows[number]
            signature = tuple(compare_value(values[db_pos], item.comparison, options) for _, db_pos, item in positions)
            database_signatures[number] = signature
            exact.setdefault((key, signature), number)
    for row_number, values in target.rows:
        if cancelled():
            raise Cancelled()
        key = normalize(values[key_position], options)
        signature = tuple(compare_value(values[file_pos], item.comparison, options) for file_pos, _, item in positions)
        matched = exact.get((key, signature))
        if not key:
            status = "BLANK KEY"
        elif key not in lookup:
            status = "NOT FOUND"
        elif matched is not None:
            status = "MATCHED"
        else:
            status = "DIFFERENT"
            # Show the candidate with the fewest differing mapped values.
            matched = min(lookup[key], key=lambda number: sum(a != b for a, b in zip(signature, database_signatures[number])))
            expected = database_rows[matched]
            for file_pos, db_pos, item in positions:
                actual, wanted = values[file_pos], expected[db_pos]
                if compare_value(actual, item.comparison, options) != compare_value(wanted, item.comparison, options):
                    label = item.file_column if item.file_column == item.database_column else f"{item.file_column} → {item.database_column}"
                    result.differences.append((row_number, matched, label, actual, wanted))
        result.records.append((row_number, status, str(matched) if matched is not None else "", values))


def check_files(database: Source, files: list[Source], options: Options,
                progress: Callable[[int, int, FileResult], None] | None = None,
                cancelled: Callable[[], bool] = lambda: False) -> list[FileResult]:
    if not database.has_headers:
        raise ValueError("The database must have a header row containing CHCODE.")
    table = read_table(database.path, database.sheet)
    if database.column not in table.headers:
        raise ValueError(f"Database column '{database.column}' no longer exists.")
    comparison_headers = [name for name in table.headers if not (options.require_flow_stage and name.casefold() == "flow_stage")]
    comparison_positions = [table.headers.index(name) for name in comparison_headers]
    position = table.headers.index(database.column)
    lookup: dict[str, list[int]] = {}
    database_rows = dict(table.rows)
    signatures = {}
    blank_rows = []
    for row_number, values in table.rows:
        if cancelled():
            raise Cancelled()
        key = normalize(values[position], options)
        if key:
            lookup.setdefault(key, []).append(row_number)
            if not options.mapped_fields:
                signatures[row_number] = tuple(normalize(values[index], options) for index in comparison_positions)
        else:
            blank_rows.append(row_number)
    if not lookup:
        raise ValueError("The database match column has no nonblank values.")
    if blank_rows:
        raise ValueError(f"The database has blank CHCODEs at row(s) {', '.join(map(str, blank_rows[:20]))}. Fill these before comparing rows.")
    results = []
    for index, source in enumerate(files, 1):
        if cancelled():
            raise Cancelled()
        result = FileResult(source, [], [], database_headers=table.headers,
                            comparison_mode="Uploaded rows" if options.mapped_fields else "Full database copy",
                            flow_stage_required=options.require_flow_stage and source.has_headers)
        try:
            if not source.has_headers and not options.mapped_fields:
                raise ValueError("Headerless files support uploaded-row checks only. Choose Validate uploaded rows.")
            target = read_table(source.path, source.sheet, has_headers=source.has_headers)
            result.headers = target.headers
            if source.column not in target.headers:
                raise ValueError(f"Match column '{source.column}' no longer exists.")
            if result.flow_stage_required:
                _identify_flow_stage(target, result)
            if options.mapped_fields:
                _compare_mapped(table, target, database, source, result, options, lookup, cancelled)
                results.append(result)
                if progress:
                    progress(index, len(files), result)
                continue
            target_position = target.headers.index(source.column)
            target_columns = {name.casefold(): i for i, name in enumerate(target.headers) if name != result.flow_stage_column}
            database_columns = {name.casefold(): i for i, name in enumerate(table.headers)}
            result.missing_columns = [name for name in comparison_headers if name.casefold() not in target_columns]
            result.extra_columns = [name for name in target.headers if name.casefold() not in database_columns and name != result.flow_stage_column]
            schema_matches = not result.missing_columns and not result.extra_columns
            # Reserve identical rows before pairing modified rows with the same CHCODE.
            # This keeps duplicate CHCODEs independent of source row order.
            available = defaultdict(deque)
            for number, signature in signatures.items():
                if cancelled():
                    raise Cancelled()
                available[signature].append(number)
            consumed = set()
            exact = {}
            if schema_matches:
                for number, values in target.rows:
                    if cancelled():
                        raise Cancelled()
                    signature = tuple(normalize(values[target_columns[name.casefold()]], options) for name in comparison_headers)
                    candidates = available.get(signature)
                    if candidates:
                        matched = candidates.popleft()
                        exact[number] = matched
                        consumed.add(matched)
            remaining = {key: deque(number for number in numbers if number not in consumed)
                         for key, numbers in lookup.items()}
            for row_number, values in target.rows:
                if cancelled():
                    raise Cancelled()
                key = normalize(values[target_position], options)
                matched = exact.get(row_number)
                if matched is not None:
                    status = "MATCHED"
                elif not key:
                    status = "BLANK KEY"
                elif key not in lookup:
                    status = "NOT FOUND"
                elif remaining[key]:
                    matched = remaining[key].popleft()
                    consumed.add(matched)
                    status = "DIFFERENT"
                    expected = database_rows[matched]
                    for name in comparison_headers:
                        name_key = name.casefold()
                        expected_value = expected[database_columns[name_key]]
                        if name_key not in target_columns:
                            result.differences.append((row_number, matched, name, "(missing column)", expected_value))
                        else:
                            actual = values[target_columns[name_key]]
                            if normalize(actual, options) != normalize(expected_value, options):
                                result.differences.append((row_number, matched, name, actual, expected_value))
                    for name in result.extra_columns:
                        result.differences.append((row_number, matched, name, values[target_columns[name.casefold()]], "(no database column)"))
                else:
                    status = "EXTRA ROW"
                result.records.append((row_number, status, str(matched) if matched is not None else "", values))
            result.missing_rows = [(number, values) for number, values in table.rows if number not in consumed]
        except Cancelled:
            raise
        except Exception as error:
            result.error = str(error)
        results.append(result)
        if progress:
            progress(index, len(files), result)
    return results


def export_report(path: Path, database: Source, options: Options, results: list[FileResult]) -> None:
    from openpyxl.cell import WriteOnlyCell

    book = Workbook(write_only=True)

    def append(sheet, values):
        cells = []
        for value in values:
            cell = WriteOnlyCell(sheet, value=value)
            if isinstance(value, str):
                cell.data_type = "s"
            cells.append(cell)
        sheet.append(cells)

    summary = book.create_sheet("Summary")
    append(summary, ["Database", str(database.path)])
    append(summary, ["Database sheet", database.sheet])
    append(summary, ["Database match column", database.column])
    append(summary, ["Ignore letter case", str(options.ignore_case)])
    append(summary, ["Trim outer spaces", str(options.trim_spaces)])
    append(summary, ["Created", datetime.now().isoformat(timespec="seconds")])
    append(summary, [])
    append(summary, ["File", "Full path", "Sheet", "Match column", "Rows", "Matched rows", "Not found", "Blank key", "Error", "Report sheet",
                     "Comparison", "Different rows", "Extra rows", "Missing database rows", "Missing columns", "Extra columns", "Differences sheet", "Missing rows sheet",
                     "Scope", "Fields checked", "Unchecked file columns", "Mappings sheet",
                     "FLOW_STAGE", "File identifier column"])
    for index, result in enumerate(results, 1):
        name = f"File {index:03d}"
        append(summary, [result.source.path.name, str(result.source.path), result.source.sheet,
                         result.source.column, len(result.records), result.count("MATCHED"),
                         result.count("NOT FOUND"), result.count("BLANK KEY"), result.error, name,
                         result.verdict, result.count("DIFFERENT"), result.count("EXTRA ROW"), len(result.missing_rows),
                         ", ".join(result.missing_columns), ", ".join(result.extra_columns), f"Differences {index:03d}",
                         f"Missing {index:03d}" if result.comparison_mode != "Uploaded rows" else "",
                         result.comparison_mode,
                         "; ".join(f"{item.file_column} → {item.database_column} ({item.comparison})" for item in result.mappings),
                         ", ".join(result.ignored_columns), f"Mappings {index:03d}" if result.comparison_mode == "Uploaded rows" else "",
                         result.flow_stage, result.flow_stage_column])
        detail = book.create_sheet(name)
        detail.freeze_panes = "D2"
        append(detail, ["Source row", "Result", "Database row(s)", *result.headers])
        for number, status, matches, values in result.records:
            append(detail, [number, status, matches, *values])
        if result.error:
            append(detail, ["ERROR", result.error])
        differences = book.create_sheet(f"Differences {index:03d}")
        differences.freeze_panes = "A2"
        append(differences, ["Source row", "Database row", "Column", "File value", "Database value"])
        for difference in result.differences:
            append(differences, difference)
        if result.comparison_mode == "Uploaded rows":
            mapping_sheet = book.create_sheet(f"Mappings {index:03d}")
            mapping_sheet.freeze_panes = "A2"
            append(mapping_sheet, ["File column", "Database column", "Comparison", "Coverage", "Rule"])
            mapped = {item.file_column: item for item in result.mappings}
            for header in result.headers:
                item = mapped.get(header)
                if header == result.flow_stage_column:
                    append(mapping_sheet, [header, "", "", "FILE IDENTIFIER", "Labels the file; row data is compared against database records by CHCODE"])
                elif item is None:
                    append(mapping_sheet, [header, "", "", "NOT CHECKED", "No database mapping selected"])
                else:
                    rule = {"Text": "Uses case and space options", "Number": "Numeric equality; commas ignored",
                            "Product": "Credit Card = CARDS; Personal Loan = LOAN"}[item.comparison]
                    append(mapping_sheet, [header, item.database_column, item.comparison, "CHECKED", rule])
        else:
            missing = book.create_sheet(f"Missing {index:03d}")
            missing.freeze_panes = "B2"
            append(missing, ["Database row", *result.database_headers])
            for number, values in result.missing_rows:
                append(missing, [number, *values])
    summary.freeze_panes = "A9"
    # Save beside the destination first, then replace atomically.
    import os
    import tempfile
    destination = Path(path)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=destination.parent, suffix=".xlsx", delete=False) as stream:
            temporary = Path(stream.name)
        book.save(temporary)
        os.replace(temporary, destination)
    finally:
        book.close()
        if temporary is not None:
            temporary.unlink(missing_ok=True)

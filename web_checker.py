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

from checker import ColumnMapping, FileResult, Options, Source, check_files, export_report, match_column, read_table, sheet_names


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


@dataclass(frozen=True)
class UploadedSource:
    upload: Upload
    sheet: str
    column: str
    mappings: tuple[ColumnMapping, ...] | None = None
    has_headers: bool = True


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
        return Inspection(tuple(names), table.sheet, tuple(table.headers), column, table.header_row)


def configuration_signature(database: UploadedSource, files: list[UploadedSource], options: Options) -> str:
    def describe(source):
        return {"upload": source.upload.fingerprint, "sheet": source.sheet, "column": source.column,
                "mappings": None if source.mappings is None else [asdict(item) for item in source.mappings],
                "has_headers": source.has_headers}
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
                          source.mappings, source.has_headers)
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

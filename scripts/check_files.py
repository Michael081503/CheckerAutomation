"""Run the same local checker from the command line and write an Excel report."""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from checker import Options, Source, check_files, export_report, find_header, match_column, read_table


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database", type=Path, required=True)
    parser.add_argument("--files", type=Path, nargs="+", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--database-sheet", default=None)
    parser.add_argument("--database-key", help="Database lookup-column header; defaults to automatic detection")
    parser.add_argument("--file-key", help="Lookup-column header to use in every file; defaults to automatic detection")
    parser.add_argument("--full-copy", action="store_true", help="Require every database row and every column")
    arguments = parser.parse_args()
    for name in ("database_key", "file_key"):
        value = getattr(arguments, name)
        if value is not None and not value.strip():
            parser.error(f"--{name.replace('_', '-')} must name a column header.")

    def source(path, sheet=None, key=None):
        table = read_table(path, sheet, headers_only=True)
        # Preserve a missing explicit column so the checker reports an error,
        # rather than silently falling back to a different lookup key.
        column = (find_header(table.headers, key) or key.strip()) if key is not None else (match_column(table.headers) or "CHCODE")
        return Source(path, table.sheet, column)

    database = source(arguments.database, arguments.database_sheet, arguments.database_key)
    files = []
    for path in arguments.files:
        try:
            files.append(source(path, key=arguments.file_key))
        except Exception:
            # Let the checker report one bad file while continuing the others.
            files.append(Source(path, "", arguments.file_key.strip() if arguments.file_key is not None else "CHCODE"))
    if any(arguments.output.resolve() == item.path.resolve() for item in [database, *files]):
        parser.error("The output must differ from every input file.")
    options = Options(mapped_fields=not arguments.full_copy, require_flow_stage=True)
    results = check_files(database, files, options)
    arguments.output.parent.mkdir(parents=True, exist_ok=True)
    export_report(arguments.output, database, options, results)
    for result in results:
        print(json.dumps({"file": result.source.path.name, "identifier": result.file_identifier,
                          "file_key": result.source.column, "database_key": result.database_match_column or database.column,
                          "FLOW_STAGE": result.flow_stage, "rows": len(result.records),
                          "matched": result.count("MATCHED"), "changed": result.count("DIFFERENT"),
                          "not_found": result.count("NOT FOUND"), "blank_key": result.count("BLANK KEY"),
                          "verdict": result.verdict, "checked_fields": len(result.mappings),
                          "unchecked_fields": result.ignored_columns, "error": result.error}))
    print(f"Report: {arguments.output.resolve()}")
    return 1 if any(result.error for result in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())

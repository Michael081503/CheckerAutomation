"""Run the same local checker from the command line and write an Excel report."""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from checker import Options, Source, check_files, export_report, read_table


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database", type=Path, required=True)
    parser.add_argument("--files", type=Path, nargs="+", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--database-sheet", default=None)
    parser.add_argument("--full-copy", action="store_true", help="Require every database row and every column")
    arguments = parser.parse_args()

    def source(path, sheet=None):
        table = read_table(path, sheet, headers_only=True)
        column = next((name for name in table.headers if name.casefold() == "chcode"), "CHCODE")
        return Source(path, table.sheet, column)

    database = source(arguments.database, arguments.database_sheet)
    files = []
    for path in arguments.files:
        try:
            files.append(source(path))
        except Exception:
            # Let the checker report one bad file while continuing the others.
            files.append(Source(path, "", "CHCODE"))
    if any(arguments.output.resolve() == item.path.resolve() for item in [database, *files]):
        parser.error("The output must differ from every input file.")
    options = Options(mapped_fields=not arguments.full_copy, require_flow_stage=True)
    results = check_files(database, files, options)
    arguments.output.parent.mkdir(parents=True, exist_ok=True)
    export_report(arguments.output, database, options, results)
    for result in results:
        print(json.dumps({"file": result.source.path.name, "FLOW_STAGE": result.flow_stage, "rows": len(result.records),
                          "matched": result.count("MATCHED"), "changed": result.count("DIFFERENT"),
                          "not_found": result.count("NOT FOUND"), "blank_key": result.count("BLANK KEY"),
                          "verdict": result.verdict, "checked_fields": len(result.mappings),
                          "unchecked_fields": result.ignored_columns, "error": result.error}))
    print(f"Report: {arguments.output.resolve()}")
    return 1 if any(result.error for result in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())

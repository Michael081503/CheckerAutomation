"""Conservative, local schema inference for unfamiliar account exports.

Identifiers establish identity first. Other fields are inferred only from the
same identified rows; a value appearing elsewhere in the database is no proof.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from checker import ColumnMapping, Options, Table, compare_value, is_customer_name_column, normalize, phone_column

NON_KEY_FIELDS = {"no", "number", "batchno", "serviceno", "statuscode", "collector", "email", "emailaddress",
                  "row", "rownumber", "index", "sequence", "seq", "count", "serial", "serialno",
                  "accounttype", "accountstatus", "accountbalance", "accountname", "accountcategory",
                  "firstname", "agentcode", "flowstage", "ob", "monthly", "terms", "promo", "level",
                  "levelendo", "status", "amount"}


def _non_key(header: str) -> bool:
    name = token(header)
    return (is_customer_name_column(header) or bool(phone_column([header])) or name in NON_KEY_FIELDS
            or name.endswith(("name", "type", "status", "balance", "amount", "date", "category"))
            or name.startswith(("staff", "agent", "collector")))


def token(header: str) -> str:
    return re.sub(r"[^a-z0-9]", "", header.casefold())


def is_assignment_template(headers) -> bool:
    """Staff CHCode/FirstName fields distinguish these from account CHCODEs."""
    names = {token(header) for header in headers}
    return ({"fullname", "firstname", "lastname", "phone", "chcode", "ob", "monthly",
             "terms", "promo", "level"}.issubset(names) and "debtorid" not in names)


def key_group(header: str) -> str:
    name = token(header)
    if name in {"debtor", "debtorid", "debtoruid"}:
        return "debtor"
    if name in {"customerid", "clientid", "customerkey", "clientkey"}:
        return "customer"
    if name in {"accountno", "accountnumber", "accountid", "caseaccountno"}:
        return "account"
    if name in {"chcode", "primarykey", "indexoneuid", "oldic"}:
        return "chcode"
    return name


def _identifier(header: str) -> bool:
    name = token(header)
    return (not _non_key(header)
            and (name in {"id", "uid", "key", "ref", "debtor", "oldic", "newic", "financingcardno"}
                 or key_group(header) == "account"
                 or any(part in name for part in ("reference", "primarykey", "chcode"))
                 or name.endswith(("id", "uid", "key", "code", "number"))))


def has_identifier_header(headers) -> bool:
    return any(_identifier(header) for header in headers)


def database_key_candidates(table: Table) -> list[str]:
    """Only complete, unique identifier columns are automatic lookup choices."""
    result = []
    for index, header in enumerate(table.headers):
        if not _identifier(header):
            continue
        values = [normalize(row[index], Options()) for _, row in table.rows]
        if values and all(values) and len(set(values)) == len(values):
            result.append(header)
    return sorted(result, key=lambda header: (
        {"debtor": 0, "customer": 1, "account": 2, "chcode": 3}.get(key_group(header), 4),
        table.headers.index(header)))


@dataclass(frozen=True)
class InferredLayout:
    file_key: str
    database_key: str
    mappings: tuple[ColumnMapping, ...]
    evidence: tuple[dict, ...]
    note: str


@dataclass
class _Key:
    file_column: str
    database_column: str
    aligned: dict[int, tuple[int, tuple]]
    hits: int
    populated: int
    related: bool


def _source_key(header: str) -> bool:
    return not _non_key(header)


def _equivalent_rows(left: _Key, right: _Key) -> bool:
    common = left.aligned.keys() & right.aligned.keys()
    return bool(common) and all(left.aligned[row][0] == right.aligned[row][0] for row in common)


def _semantic_target(header: str, database_headers: list[str]) -> str:
    name = token(header)
    groups = (
        {"fullname", "name", "chname", "customername", "clientname", "debtorname", "casecustomername"},
        {"phone", "mobile", "mobilephone", "mobilenumber", "phonenumber", "lastcalledno", "contactnumber"},
        {"email", "emailaddress", "casecustomeremail"},
        {"accountnumber", "accountno", "accountid", "caseaccountno"},
    )
    same = next((value for value in database_headers if token(value) == name), "")
    if same:
        return same
    for group in groups:
        if name in group:
            return next((value for value in database_headers if token(value) in group), "")
    return ""


def infer_layout(target: Table, database: Table) -> InferredLayout:
    """Return one whole-file lookup and mappings, or explain unsafe ambiguity."""
    db_keys = database_key_candidates(database)
    candidates = []
    options = Options()
    assignment = is_assignment_template(target.headers)
    assignment_fields = {"fullname": "Name", "lastname": "Account No.", "phone": "Email"}
    established_assignment = assignment and all(
        any(token(header) == token(expected) for header in database.headers)
        for expected in (*assignment_fields.values(), "Debtor ID"))
    unnamed = [header for header in target.headers if token(header).startswith("unnamedcolumn")]
    for database_column in db_keys:
        db_pos = database.headers.index(database_column)
        lookup = {normalize(values[db_pos], options): (number, values) for number, values in database.rows}
        for file_pos, file_column in enumerate(target.headers):
            if not _source_key(file_column):
                continue
            values = [(number, normalize(row[file_pos], options)) for number, row in target.rows]
            populated = [value for _, value in values if value]
            distinct = set(populated)
            related = (key_group(file_column) == key_group(database_column)
                       or (established_assignment and key_group(database_column) == "debtor"
                           and unnamed == [file_column]))
            if assignment and token(file_column) in {"chcode", "firstname"}:
                continue
            if not populated and not related:
                continue
            # Repeated staff/level codes and sequence counters must not identify customers.
            if not related and len(distinct) / len(populated) < 0.5:
                continue
            hits = distinct & lookup.keys()
            if not related and (len(hits) < 3 or len(hits) / len(distinct) < 0.8):
                continue
            aligned = {number: lookup[value] for number, value in values if value in lookup}
            candidates.append(_Key(file_column, database_column, aligned, len(aligned), len(populated), related))
    if not candidates:
        raise ValueError("No reliable identifier match was found. Use a database with matching debtor, customer, or account IDs; names, staff codes, and repeated values cannot establish identity.")
    candidates.sort(key=lambda candidate: (
        -int(candidate.related), -candidate.hits / max(1, candidate.populated),
        db_keys.index(candidate.database_column), target.headers.index(candidate.file_column)))
    selected = candidates[0]
    for other in candidates[1:]:
        if (other.hits >= 3 and selected.hits >= 3
                and other.hits / max(1, other.populated) >= 0.8
                and selected.hits / max(1, selected.populated) >= 0.8
                and not _equivalent_rows(selected, other)):
            raise ValueError(f"Ambiguous identifier matches: {selected.file_column} / {selected.database_column} and {other.file_column} / {other.database_column} point to different records. Supply a file with one consistent account identifier.")
    mappings = [ColumnMapping(selected.file_column, selected.database_column)]
    evidence = [{"File column": selected.file_column, "Database column": selected.database_column,
                 "Comparison": "Text", "Evidence": f"Lookup: {selected.hits:,} of {len(target.rows):,} rows found; database IDs are complete and unique"}]
    warnings = []
    for file_pos, file_column in enumerate(target.headers):
        if file_column == selected.file_column or token(file_column) == "flowstage":
            continue
        if assignment and token(file_column) in {"firstname", "chcode", "ob", "monthly", "terms", "promo", "level"}:
            continue
        established = assignment_fields.get(token(file_column)) if established_assignment else None
        semantic = (next((header for header in database.headers if token(header) == token(established)), "")
                    if established else _semantic_target(file_column, database.headers))
        # Preserve header-established fields even when every uploaded value is wrong.
        # A stronger same-record value match can reveal misleading source headers.
        comparison = "Phone" if phone_column([file_column]) else "Text"
        scored = []
        for db_pos, database_column in enumerate(database.headers):
            samples = []
            for number, row in target.rows:
                if number not in selected.aligned:
                    continue
                expected = selected.aligned[number][1][db_pos]
                actual = row[file_pos]
                # Blank/blank pairs never support an inferred mapping.
                if not normalize(actual, options) or not normalize(expected, options):
                    continue
                actual_value = compare_value(actual, comparison, options,
                                             customer_name=is_customer_name_column(database_column))
                expected_value = compare_value(expected, comparison, options,
                                               customer_name=is_customer_name_column(database_column))
                samples.append((actual_value, expected_value))
            matches = sum(left == right for left, right in samples)
            diversity = len({left for left, right in samples if left == right})
            # Include blank/missing values in the denominator, so a sparse
            # accidental intersection cannot look like perfect evidence.
            total = len(selected.aligned)
            if total and diversity >= 3 and matches / total >= 0.8:
                scored.append((matches / total, database_column, matches, total))
        scored.sort(key=lambda item: (-item[0], item[1] != semantic,
                                      {"accountno": 0, "name": 1, "email": 2}.get(token(item[1]), 3),
                                      database.headers.index(item[1])))
        chosen = semantic if established else scored[0][1] if scored else semantic
        if not chosen:
            continue
        if established:
            rule = "Established staff-assignment template field; every value is checked on the same debtor record"
        elif scored:
            tied = [item[1] for item in scored if item[0] == scored[0][0]]
            chosen_pos = database.headers.index(chosen)
            equivalent = all(all(normalize(values[chosen_pos], options) == normalize(values[database.headers.index(other)], options)
                                 for _, values in selected.aligned.values()) for other in tied)
            if len(tied) > 1 and not equivalent and semantic not in tied:
                warnings.append(f"{file_column} is unchecked because several database fields match equally well")
                continue
            rule = f"{scored[0][2]:,} of {scored[0][3]:,} identified rows agree"
            if len(tied) > 1:
                rule += "; equivalent candidates: " + ", ".join(tied)
        else:
            rule = "Matching header or established field alias; every value is still checked"
        if comparison == "Text" and phone_column([chosen]) and phone_column([file_column]):
            comparison = "Phone"
        mappings.append(ColumnMapping(file_column, chosen, comparison))
        evidence.append({"File column": file_column, "Database column": chosen,
                         "Comparison": comparison, "Evidence": rule})
        if chosen != _semantic_target(file_column, database.headers):
            warnings.append(f"{file_column} maps to {chosen}: " + (
                "established template mapping" if established else "values agree on the same identified record"))
    if len(mappings) == 1:
        warnings.append("Only the identifier could be checked; other fields remain unchecked")
    note = (f"Detected lookup {selected.file_column} to {selected.database_column}; "
            f"{selected.hits:,} of {len(target.rows):,} rows found. " + "; ".join(warnings))
    return InferredLayout(selected.file_column, selected.database_column, tuple(mappings), tuple(evidence), note.rstrip(". ;"))

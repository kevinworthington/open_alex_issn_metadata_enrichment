#!/usr/bin/env python3
"""Fetch OpenAlex source metadata for ISSNs and maintain an Excel workbook."""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

API_URL = "https://api.openalex.org/sources"
HEADERS = [
    "input_issns", "openalex_id", "display_name", "issn_l", "issn",
    "publisher", "type", "country_code", "homepage_url", "works_count",
    "cited_by_count", "is_oa", "is_in_doaj", "created_date", "updated_date",
    "abbreviated_title", "alternate_titles", "counts_by_year",
    "summary_stats", "status", "last_checked_utc", "error",
]
CORE_FIELDS = (
    "openalex_id", "display_name", "publisher", "type", "works_count",
)
ALIASES = {
    "openalex_id": ("openalex_id", "id", "OpenAlex ID", "OpenAlex Id"),
    "display_name": ("display_name", "source_name", "name", "Source Name"),
    "issn_l": ("issn_l", "ISSN-L", "issn_l"),
    "issn": ("issn", "ISSNs", "issn_list"),
    "publisher": ("publisher", "Publisher"),
    "type": ("type", "source_type", "Source Type"),
    "country_code": ("country_code", "country", "Country"),
    "homepage_url": ("homepage_url", "homepage", "Homepage"),
    "works_count": ("works_count", "Works Count"),
    "cited_by_count": ("cited_by_count", "Cited By Count"),
    "is_oa": ("is_oa", "Open Access"),
    "is_in_doaj": ("is_in_doaj", "DOAJ"),
    "created_date": ("created_date", "Created Date"),
    "updated_date": ("updated_date", "Updated Date"),
    "abbreviated_title": ("abbreviated_title", "Abbreviated Title"),
    "alternate_titles": ("alternate_titles", "Alternate Titles"),
    "counts_by_year": ("counts_by_year", "Counts by Year"),
    "summary_stats": ("summary_stats", "Summary Stats"),
    "input_issns": ("input_issns", "input_issn", "ISSN input", "queried_issn"),
    "status": ("status", "Status"),
    "last_checked_utc": ("last_checked_utc", "last_checked", "Last Checked"),
    "error": ("error", "Error"),
}
LOG = logging.getLogger("openalex_issn")


def normalize_issn(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip().upper()
    if not text:
        return None
    digits = re.sub(r"[^0-9X]", "", text)
    if len(digits) == 8:
        return f"{digits[:4]}-{digits[4:]}"
    return text


def read_issns(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    if path.suffix.lower() == ".csv":
        import csv
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        if not rows:
            return []
        start = 1 if rows[0] and rows[0][0].strip().lower() == "issn" else 0
        values = [row[0] for row in rows[start:] if row]
    else:
        values = path.read_text(encoding="utf-8-sig").splitlines()
    result, seen = [], set()
    for value in values:
        issn = normalize_issn(value)
        if issn and issn not in seen:
            result.append(issn)
            seen.add(issn)
    return result


def json_cell(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return value


def api_record_to_row(record: dict[str, Any], query_issn: str) -> dict[str, Any]:
    return {
        "input_issns": query_issn,
        "openalex_id": record.get("id"),
        "display_name": record.get("display_name"),
        "issn_l": record.get("issn_l"),
        "issn": json_cell(record.get("issn")),
        "publisher": record.get("host_organization_name"),
        "type": record.get("type"),
        "country_code": record.get("country_code"),
        "homepage_url": record.get("homepage_url"),
        "works_count": record.get("works_count"),
        "cited_by_count": record.get("cited_by_count"),
        "is_oa": record.get("is_oa"),
        "is_in_doaj": record.get("is_in_doaj"),
        "created_date": record.get("created_date"),
        "updated_date": record.get("updated_date"),
        "abbreviated_title": json_cell(record.get("abbreviated_title")),
        "alternate_titles": json_cell(record.get("alternate_titles")),
        "counts_by_year": json_cell(record.get("counts_by_year")),
        "summary_stats": json_cell(record.get("summary_stats")),
        "status": "found",
        "last_checked_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "error": None,
    }


def row_issns(row: dict[str, Any]) -> set[str]:
    values: set[str] = set()
    for key in ("input_issns", "issn_l", "issn"):
        value = row.get(key)
        if value is None:
            continue
        if isinstance(value, str):
            try:
                parsed = json.loads(value)
                candidates = parsed if isinstance(parsed, list) else [value]
            except (json.JSONDecodeError, TypeError):
                candidates = re.split(r"[,;|]", value)
        elif isinstance(value, list):
            candidates = value
        else:
            candidates = [value]
        for candidate in candidates:
            normalized = normalize_issn(candidate)
            if normalized:
                values.add(normalized)
    return values


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def is_incomplete(row: dict[str, Any]) -> bool:
    return any(is_blank(row.get(field)) for field in CORE_FIELDS)


def load_existing(path: Path) -> tuple[list[dict[str, Any]], list[str]]:
    if not path.exists():
        return [], list(HEADERS)
    workbook = load_workbook(path, data_only=True)
    sheet = workbook.active
    raw_headers = [cell.value for cell in sheet[1]]
    headers = [str(value).strip() if value is not None else "" for value in raw_headers]
    rows: list[dict[str, Any]] = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not any(value is not None for value in values):
            continue
        old = {headers[i]: values[i] for i in range(min(len(headers), len(values))) if headers[i]}
        row = {column: None for column in HEADERS}
        for target, aliases in ALIASES.items():
            for alias in aliases:
                if alias in old and not is_blank(old[alias]):
                    row[target] = json_cell(old[alias])
                    break
        # Preserve columns not managed by this version.
        for key, value in old.items():
            if key not in row and key not in HEADERS:
                row[key] = value
        rows.append(row)
    extras = []
    for row in rows:
        for key in row:
            if key not in HEADERS and key not in extras:
                extras.append(key)
    return rows, list(HEADERS) + extras


def find_existing(rows: list[dict[str, Any]], record: dict[str, Any]) -> dict[str, Any] | None:
    record_id = record.get("openalex_id")
    if record_id:
        for row in rows:
            if str(row.get("openalex_id") or "").strip().lower() == str(record_id).strip().lower():
                return row
    candidates = row_issns(record)
    if candidates:
        for row in rows:
            if candidates.intersection(row_issns(row)):
                return row
    return None


def merge_record(rows: list[dict[str, Any]], fresh: dict[str, Any]) -> None:
    existing = find_existing(rows, fresh)
    if existing is None:
        rows.append(fresh)
        return
    for key, value in fresh.items():
        if key == "input_issns":
            all_issns = row_issns(existing) | row_issns(fresh)
            existing[key] = "; ".join(sorted(all_issns))
        elif is_blank(existing.get(key)) and not is_blank(value):
            existing[key] = value
    # A successful lookup resolves an old error/status without overwriting metadata.
    existing["status"] = "found"
    existing["last_checked_utc"] = fresh["last_checked_utc"]
    existing["error"] = None


def fetch_sources(session: requests.Session, issn: str, api_key: str | None) -> list[dict[str, Any]]:
    params: dict[str, Any] = {"filter": f"issn:{issn}", "per-page": 100}
    if api_key:
        params["api_key"] = api_key
    last_error: Exception | None = None
    for attempt in range(4):
        try:
            response = session.get(API_URL, params=params, timeout=30)
            if response.status_code == 429 or response.status_code >= 500:
                response.raise_for_status()
            response.raise_for_status()
            payload = response.json()
            return payload.get("results", [])
        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            if attempt == 3:
                break
            wait = min(2 ** attempt, 8)
            LOG.warning("Request for %s failed (%s); retrying in %ss.", issn, exc, wait)
            time.sleep(wait)
    raise RuntimeError(str(last_error))


def issns_from_existing_incomplete(rows: list[dict[str, Any]]) -> set[str]:
    return {
        issn
        for row in rows
        if is_incomplete(row)
        for issn in row_issns(row)
    }


def safe_excel_value(value: Any) -> Any:
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def save_workbook(path: Path, rows: list[dict[str, Any]], headers: list[str]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "OpenAlex Sources"
    sheet.append(headers)
    for row in rows:
        sheet.append([safe_excel_value(row.get(header)) for header in headers])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for index, header in enumerate(headers, start=1):
        width = min(max(len(header) + 2, 14), 42)
        sheet.column_dimensions[get_column_letter(index)].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def configure_logging(level: str) -> None:
    logging.basicConfig(
        level={"quiet": logging.ERROR, "normal": logging.INFO,
               "verbose": logging.DEBUG, "debug": logging.DEBUG}[level],
        format="%(levelname)s: %(message)s",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Look up ISSNs in OpenAlex and create or update an Excel workbook."
    )
    parser.add_argument("input", type=Path, help="Text file (one ISSN per line) or CSV")
    parser.add_argument("-o", "--output", type=Path, default=Path("openalex_metadata.xlsx"),
                        help="Output workbook; an existing workbook is retained and completed")
    parser.add_argument("--api-key", default=None,
                        help="Optional OpenAlex API key (or set OPENALEX_API_KEY)")
    parser.add_argument("--verbosity", choices=("quiet", "normal", "verbose", "debug"),
                        default="normal", help="Terminal output detail (default: normal)")
    parser.add_argument("-v", "--verbose", action="count", default=0,
                        help="Increase verbosity; repeat for debug (overridden by --verbosity)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    level = args.verbosity
    if args.verbose:
        level = "verbose" if args.verbose == 1 else "debug"
    configure_logging(level)
    try:
        input_issns = read_issns(args.input)
        rows, headers = load_existing(args.output)
    except (OSError, ValueError) as exc:
        LOG.error("%s", exc)
        return 2

    api_key = args.api_key or os.getenv("OPENALEX_API_KEY")
    if not input_issns and not rows:
        LOG.error("No ISSNs found in the input file.")
        return 2

    # Always query requested ISSNs; also revisit incomplete saved rows, even if
    # they are not in this run's input list.
    targets = list(dict.fromkeys(input_issns + sorted(issns_from_existing_incomplete(rows))))
    LOG.info("Loaded %d input ISSN(s); %d existing record(s).", len(input_issns), len(rows))
    if not api_key:
        LOG.debug("No API key supplied; using the public OpenAlex endpoint.")

    session = requests.Session()
    failed = 0
    for index, issn in enumerate(targets, start=1):
        LOG.info("[%d/%d] Looking up %s", index, len(targets), issn)
        try:
            records = fetch_sources(session, issn, api_key)
        except RuntimeError as exc:
            failed += 1
            LOG.error("Lookup failed for %s: %s", issn, exc)
            for row in rows:
                if issn in row_issns(row) and is_incomplete(row):
                    row["error"] = str(exc)
                    row["status"] = "error"
            continue

        if records:
            LOG.debug("%s returned %d source record(s).", issn, len(records))
            for record in records:
                merge_record(rows, api_record_to_row(record, issn))
        else:
            LOG.warning("No source found for %s.", issn)
            if not any(issn in row_issns(row) and row.get("status") == "not_found" for row in rows):
                rows.append({
                    **{header: None for header in headers},
                    "input_issns": issn,
                    "status": "not_found",
                    "last_checked_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                })

    try:
        save_workbook(args.output, rows, headers)
    except OSError as exc:
        LOG.error("Could not write workbook: %s", exc)
        return 2

    LOG.info("Saved %d record row(s) to %s.", len(rows), args.output)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())

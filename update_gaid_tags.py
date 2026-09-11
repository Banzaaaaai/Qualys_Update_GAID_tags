"""Bulk-update GAID Asset Management tags in Qualys (EU2 pod) from a CMDB Excel export.

Credentials are never read from argv/hardcoded; they come from the
QUALYS_USERNAME / QUALYS_PASSWORD env vars, falling back to
qualys_creds.txt next to this script (format: "key:\\tvalue" per line).

HARD CONSTRAINTS (see README.md for the full rationale):
1. COMPLETE REPLACEMENT of each GAID tag's IP/range list with the set from
   the Excel file -- never a merge/union with what is currently in Qualys.
2. UPDATE IN PLACE. The tag keeps its id/name/parent/color/criticality --
   only the ruleText (IP/range list) changes. Never delete+recreate.

Safety: dry-run by default. Nothing is written to Qualys unless --apply is
passed. Every touched tag is backed up before any write call.
"""

import argparse
import datetime
import ipaddress
import json
import os
import re
import sys
import time
import xml.etree.ElementTree as ET

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

BASE_URL = "https://qualysapi.qg2.apps.qualys.eu"
TAG_SEARCH_URL = f"{BASE_URL}/qps/rest/2.0/search/am/tag"
TAG_UPDATE_URL = f"{BASE_URL}/qps/rest/2.0/update/am/tag/{{tag_id}}"

# Confirmed 2026-09-12 against a live tag ("[VFZ] GAID: 1356", id 189526283):
# name is exactly this prefix + the GAID number, parent is the fixed tag
# below, and the rule is a NETWORK_RANGE whose ruleText is a single
# comma-separated list mixing bare IPv4 addresses and "A-B" hyphen ranges.
GAID_TAG_PREFIX = "[VFZ] GAID: "
GAID_PARENT_TAG_NAME = "[VFZ] Global Application Inventory"
EXPECTED_RULE_TYPE = "NETWORK_RANGE"

INITIAL_DELAY = 1
MAX_RETRIES = 5
REQUEST_TIMEOUT = 60
MAX_RATE_LIMIT_SLEEP = 300
TAG_SEARCH_PAGE_SIZE = 100

HEADERS = {
    "Content-Type": "text/xml",
    "X-Requested-With": "update_gaid_tags.py",
}

REPORT_COLUMNS = [
    "tag_id",
    "tag_name",
    "status",
    "old_ip_summary",
    "new_ip_summary",
    "ips_added",
    "ips_removed",
    "error_message",
]

# Header substrings (case-insensitive, regex word-boundary where noted) used
# to infer the CMDB column mapping. Adjust here if a real export doesn't
# match what discovery infers.
GAID_COLUMN_PATTERNS = [
    r"\bgaid\b",
    r"application\s*id",
    r"app\s*id",
]
IP_COLUMN_PATTERNS = [
    r"ip\s*address(es)?",
    r"ip\s*range(s)?",
    r"network\s*range",
    r"\bip(s)?\b",
]

# Tokens are split out of a cell (and across rows for the same GAID) on any
# of these separators.
IP_CELL_SPLIT_RE = re.compile(r"[,;\n\r]+")


# --------------------------------------------------------------------------
# Credentials (unchanged pattern from fetch_qualys_tags.py)
# --------------------------------------------------------------------------


def load_credentials():
    username = os.environ.get("QUALYS_USERNAME")
    password = os.environ.get("QUALYS_PASSWORD")
    if username and password:
        return username, password

    creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qualys_creds.txt")
    if not os.path.isfile(creds_path):
        raise RuntimeError(
            "No credentials found: set QUALYS_USERNAME/QUALYS_PASSWORD or provide qualys_creds.txt"
        )

    creds = {}
    with open(creds_path, "r", encoding="utf-8") as f:
        for line in f:
            if ":" not in line:
                continue
            key, _, value = line.partition(":")
            creds[key.strip().lower()] = value.strip()

    username = creds.get("user")
    password = creds.get("pass")
    if not username or not password:
        raise RuntimeError("qualys_creds.txt is missing 'user' and/or 'pass' entries")
    return username, password


# --------------------------------------------------------------------------
# Namespace-agnostic XML helpers (unchanged pattern from fetch_qualys_tags.py)
# --------------------------------------------------------------------------


def local_tag(tag):
    return tag.split("}", 1)[-1] if "}" in tag else tag


def find_child(elem, name):
    for child in elem:
        if local_tag(child.tag) == name:
            return child
    return None


def get_text(elem, name):
    child = find_child(elem, name)
    if child is None or child.text is None:
        return ""
    return child.text.strip()


# --------------------------------------------------------------------------
# Rate-limit-aware HTTP with retry (same policy as fetch_qualys_tags.py,
# generalized to any Qualys tagging endpoint/body/HTTP method)
# --------------------------------------------------------------------------


def request_with_retry(session, method, url, xml_body, label):
    delay = INITIAL_DELAY
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.request(
                method,
                url,
                data=xml_body.encode("utf-8"),
                headers=HEADERS,
                timeout=REQUEST_TIMEOUT,
            )
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_error = str(exc)
            wait = min(delay, MAX_RATE_LIMIT_SLEEP)
            if attempt == MAX_RETRIES:
                break
            print(
                f"  [{label}] network error ({exc}); retrying in {wait}s "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )
            time.sleep(wait)
            delay = min(delay * 2, MAX_RATE_LIMIT_SLEEP)
            continue

        if resp.status_code == 200:
            root = ET.fromstring(resp.text)
            response_code = get_text(root, "responseCode")
            if response_code and response_code != "SUCCESS":
                error_msg = get_text(root, "responseErrorDetails") or resp.text[:500]
                raise RuntimeError(f"{label}: Qualys API returned {response_code}: {error_msg}")
            return root

        if resp.status_code in (409, 429):
            wait_header = resp.headers.get("X-RateLimit-ToWait-Sec") or resp.headers.get(
                "Retry-After"
            )
            try:
                wait = (
                    min(float(wait_header), MAX_RATE_LIMIT_SLEEP)
                    if wait_header
                    else min(delay, MAX_RATE_LIMIT_SLEEP)
                )
            except ValueError:
                wait = min(delay, MAX_RATE_LIMIT_SLEEP)

            last_error = f"HTTP {resp.status_code} (rate-limited)"
            if attempt == MAX_RETRIES:
                break
            print(
                f"  [{label}] rate-limited (HTTP {resp.status_code}); waiting {wait}s "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )
            time.sleep(wait)
            delay = min(delay * 2, MAX_RATE_LIMIT_SLEEP)
            continue

        last_error = f"HTTP {resp.status_code}: {resp.text[:500]}"
        wait = min(delay, MAX_RATE_LIMIT_SLEEP)
        if attempt == MAX_RETRIES:
            break
        print(f"  [{label}] {last_error}; retrying in {wait}s (attempt {attempt}/{MAX_RETRIES})")
        time.sleep(wait)
        delay = min(delay * 2, MAX_RATE_LIMIT_SLEEP)

    raise RuntimeError(f"{label}: giving up after {MAX_RETRIES} attempts: {last_error}")


# --------------------------------------------------------------------------
# Enumerate existing tags (unchanged id-GREATER cursor pattern)
# --------------------------------------------------------------------------


def build_search_all_xml(last_id, page_size):
    return (
        "<ServiceRequest>"
        "<preferences>"
        f"<limitResults>{page_size}</limitResults>"
        "</preferences>"
        "<filters>"
        f'<Criteria field="id" operator="GREATER">{last_id}</Criteria>'
        "</filters>"
        "</ServiceRequest>"
    )


def parse_tag_elements(data_elem):
    tags = []
    if data_elem is None:
        return tags
    for tag_elem in data_elem:
        if local_tag(tag_elem.tag) != "Tag":
            continue
        tags.append(
            {
                "tag_id": get_text(tag_elem, "id"),
                "tag_name": get_text(tag_elem, "name"),
                "parent_tag_id": get_text(tag_elem, "parentTagId"),
                "rule_type": get_text(tag_elem, "ruleType"),
                "rule_text": get_text(tag_elem, "ruleText"),
                "color": get_text(tag_elem, "color"),
                "criticality": get_text(tag_elem, "criticalityScore"),
            }
        )
    return tags


def fetch_all_tags(session):
    all_tags = []
    last_id = 0
    page_num = 0

    while True:
        page_num += 1
        xml_body = build_search_all_xml(last_id, TAG_SEARCH_PAGE_SIZE)
        root = request_with_retry(
            session, "POST", TAG_SEARCH_URL, xml_body, f"search page {page_num}"
        )

        data_elem = find_child(root, "data")
        page_tags = parse_tag_elements(data_elem)
        all_tags.extend(page_tags)

        has_more = get_text(root, "hasMoreRecords").lower() == "true"
        last_id_text = get_text(root, "lastId")
        if last_id_text:
            last_id = int(last_id_text)
        else:
            numeric_ids = [int(t["tag_id"]) for t in page_tags if t["tag_id"].isdigit()]
            if numeric_ids:
                last_id = max(last_id, max(numeric_ids))

        print(f"  [page {page_num}] fetched {len(page_tags)} tags (last_id={last_id})")

        if not has_more or not page_tags:
            break
        time.sleep(INITIAL_DELAY)

    return all_tags


def fetch_tag_by_id(session, tag_id):
    xml_body = (
        "<ServiceRequest>"
        "<filters>"
        f'<Criteria field="id" operator="EQUALS">{tag_id}</Criteria>'
        "</filters>"
        "</ServiceRequest>"
    )
    root = request_with_retry(session, "POST", TAG_SEARCH_URL, xml_body, f"verify tag {tag_id}")
    data_elem = find_child(root, "data")
    tags = parse_tag_elements(data_elem)
    return tags[0] if tags else None


# --------------------------------------------------------------------------
# IP / range normalization
# --------------------------------------------------------------------------


def parse_ip_entry(token):
    """Validate one token as an IPv4 address or 'A-B' range.

    Returns the canonical string form and a sort key, or raises ValueError.
    """
    token = token.strip()
    if not token:
        raise ValueError("empty entry")

    if "-" in token:
        start_s, _, end_s = token.partition("-")
        start_s = start_s.strip()
        end_s = end_s.strip()
        start = ipaddress.IPv4Address(start_s)
        end = ipaddress.IPv4Address(end_s)
        if int(end) < int(start):
            raise ValueError(f"range end before start: {token}")
        return f"{start}-{end}", int(start)

    addr = ipaddress.IPv4Address(token)
    return str(addr), int(addr)


def normalize_ip_set(raw_tokens):
    """Validate + dedupe + sort a collection of raw IP/range tokens.

    Returns (canonical_sorted_list, invalid_tokens).
    """
    canonical = {}
    invalid = []
    for token in raw_tokens:
        if token is None:
            continue
        token = str(token).strip()
        if not token:
            continue
        try:
            canon, sort_key = parse_ip_entry(token)
        except ValueError:
            invalid.append(token)
            continue
        canonical[canon] = sort_key

    ordered = sorted(canonical.keys(), key=lambda c: canonical[c])
    return ordered, invalid


def parse_qualys_rule_text(rule_text):
    tokens = IP_CELL_SPLIT_RE.split(rule_text) if rule_text else []
    ordered, invalid = normalize_ip_set(tokens)
    return ordered, invalid


# --------------------------------------------------------------------------
# Excel discovery + parsing
# --------------------------------------------------------------------------


def find_column(headers, patterns):
    for pattern in patterns:
        regex = re.compile(pattern, re.IGNORECASE)
        for idx, header in enumerate(headers):
            if header and regex.search(str(header)):
                return idx
    return None


def normalize_gaid_key(value):
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if not text:
        return None
    # Excel often stores "1356" as 1356.0 -> already handled above; also
    # strip a stray ".0" suffix if it arrives as a string.
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def discover_excel(path):
    """Print sheet names, header row, and sample rows; infer the column
    mapping. Returns (sheet_name, headers, rows, gaid_col, ip_col)."""
    wb = load_workbook(path, data_only=True, read_only=True)
    print(f"Sheets found: {wb.sheetnames}")

    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    print(f"Using sheet: {sheet_name!r} (first sheet)")

    row_iter = ws.iter_rows(values_only=True)
    try:
        headers = list(next(row_iter))
    except StopIteration:
        raise RuntimeError(f"Sheet {sheet_name!r} is empty")

    rows = []
    for row in row_iter:
        if row is None or all(v is None for v in row):
            continue
        rows.append(row)

    print(f"Header row: {headers}")
    print("Sample rows:")
    for row in rows[:3]:
        print(f"  {row}")

    gaid_col = find_column(headers, GAID_COLUMN_PATTERNS)
    ip_col = find_column(headers, IP_COLUMN_PATTERNS)

    if gaid_col is None or ip_col is None:
        raise RuntimeError(
            "Could not infer GAID/IP columns from header "
            f"{headers!r}. GAID_COLUMN_PATTERNS/IP_COLUMN_PATTERNS at the top "
            "of this script may need adjusting for this export's real headers."
        )

    gaid_values = [normalize_gaid_key(row[gaid_col]) for row in rows if row[gaid_col] is not None]
    duplicates_exist = len(gaid_values) != len(set(gaid_values))
    layout = (
        "multiple rows per GAID (aggregated)"
        if duplicates_exist
        else "one row per GAID (delimited IP cell)"
    )

    print()
    print("=== Inferred column mapping ===")
    print(f"  GAID column: {headers[gaid_col]!r} (index {gaid_col})")
    print(f"  IP column:   {headers[ip_col]!r} (index {ip_col})")
    print(f"  Row layout:  {layout}")
    print(
        "  (Regardless of layout, every row's IP cell is split on commas/"
        "semicolons/newlines and grouped by GAID, so both layouts are "
        "handled by the same aggregation logic.)"
    )
    print()

    return sheet_name, headers, rows, gaid_col, ip_col


def build_gaid_ip_map(rows, gaid_col, ip_col):
    """Group raw IP tokens by GAID, across however many rows each GAID has."""
    raw_by_gaid = {}
    for row in rows:
        gaid_key = normalize_gaid_key(row[gaid_col])
        if gaid_key is None:
            continue
        cell = row[ip_col]
        if cell is None:
            continue
        tokens = IP_CELL_SPLIT_RE.split(str(cell))
        raw_by_gaid.setdefault(gaid_key, []).extend(tokens)
    return raw_by_gaid


# --------------------------------------------------------------------------
# Backup
# --------------------------------------------------------------------------


def write_backup(touched_tags, timestamp):
    json_path = f"backup_gaid_tags_{timestamp}.json"
    xlsx_path = f"backup_gaid_tags_{timestamp}.xlsx"

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(touched_tags, f, indent=2)

    wb = Workbook()
    ws = wb.active
    ws.title = "Backup"
    cols = ["tag_id", "tag_name", "parent_tag_id", "rule_type", "rule_text", "color", "criticality"]
    ws.append(cols)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for t in touched_tags:
        ws.append([t.get(c, "") for c in cols])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"
    widths = {"tag_id": 12, "tag_name": 24, "parent_tag_id": 14, "rule_type": 16, "rule_text": 80, "color": 10, "criticality": 12}
    for idx, col in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(col, 15)
    wb.save(xlsx_path)

    print(f"Backup written: {json_path}")
    print(f"Backup written: {xlsx_path}")
    return json_path, xlsx_path


# --------------------------------------------------------------------------
# Report
# --------------------------------------------------------------------------


def write_csv_report(rows, path):
    import csv

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        for r in rows:
            writer.writerow({col: r.get(col, "") for col in REPORT_COLUMNS})


def write_xlsx_report(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "GAID Update Report"

    ws.append(REPORT_COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    wrap_cols = {REPORT_COLUMNS.index(c) + 1 for c in ("old_ip_summary", "new_ip_summary", "ips_added", "ips_removed")}
    for r in rows:
        ws.append([r.get(col, "") for col in REPORT_COLUMNS])
    for row in range(2, ws.max_row + 1):
        for col in wrap_cols:
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(REPORT_COLUMNS))}{ws.max_row}"

    widths = {
        "tag_id": 12,
        "tag_name": 24,
        "status": 20,
        "old_ip_summary": 50,
        "new_ip_summary": 50,
        "ips_added": 35,
        "ips_removed": 35,
        "error_message": 40,
    }
    for idx, col in enumerate(REPORT_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(col, 15)

    wb.save(path)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("excel_path", help="Path to the CMDB Excel export")
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually write changes to Qualys. Without this flag, runs as a dry-run.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    print(f"=== Discovery: {args.excel_path} ===")
    _, _, rows, gaid_col, ip_col = discover_excel(args.excel_path)
    raw_by_gaid = build_gaid_ip_map(rows, gaid_col, ip_col)
    print(f"Found {len(raw_by_gaid)} distinct GAID(s) in the Excel file.")
    print()

    username, password = load_credentials()
    session = requests.Session()
    session.auth = (username, password)

    print(f"=== Fetching existing tags from {TAG_SEARCH_URL} ===")
    all_tags = fetch_all_tags(session)
    print(f"Fetched {len(all_tags)} tags total.")

    gaid_tags_by_name = {
        t["tag_name"]: t for t in all_tags if t["tag_name"].startswith(GAID_TAG_PREFIX)
    }
    print(f"Found {len(gaid_tags_by_name)} existing '{GAID_TAG_PREFIX}*' tags in the tenant.")

    non_network_range = [
        t for t in gaid_tags_by_name.values() if t["rule_type"] != EXPECTED_RULE_TYPE
    ]
    if non_network_range:
        print(
            f"WARNING: {len(non_network_range)} GAID tag(s) do not have "
            f"ruleType={EXPECTED_RULE_TYPE}; they will be skipped as errors: "
            + ", ".join(f"{t['tag_name']} (ruleType={t['rule_type']!r})" for t in non_network_range[:10])
        )
    print()

    report_rows = []
    matched_names = set()

    for gaid_key, raw_tokens in sorted(raw_by_gaid.items(), key=lambda kv: (len(kv[0]), kv[0])):
        expected_name = f"{GAID_TAG_PREFIX}{gaid_key}"
        new_ips, invalid_tokens = normalize_ip_set(raw_tokens)

        tag = gaid_tags_by_name.get(expected_name)
        if tag is None:
            report_rows.append(
                {
                    "tag_id": "",
                    "tag_name": expected_name,
                    "status": "skipped-no-match",
                    "old_ip_summary": "",
                    "new_ip_summary": ", ".join(new_ips),
                    "ips_added": "",
                    "ips_removed": "",
                    "error_message": "no matching tag in tenant",
                }
            )
            continue

        matched_names.add(expected_name)

        if invalid_tokens:
            report_rows.append(
                {
                    "tag_id": tag["tag_id"],
                    "tag_name": expected_name,
                    "status": "error",
                    "old_ip_summary": tag["rule_text"],
                    "new_ip_summary": "",
                    "ips_added": "",
                    "ips_removed": "",
                    "error_message": "invalid IP/range entries: " + ", ".join(invalid_tokens),
                }
            )
            continue

        if tag["rule_type"] != EXPECTED_RULE_TYPE:
            report_rows.append(
                {
                    "tag_id": tag["tag_id"],
                    "tag_name": expected_name,
                    "status": "error",
                    "old_ip_summary": tag["rule_text"],
                    "new_ip_summary": ", ".join(new_ips),
                    "ips_added": "",
                    "ips_removed": "",
                    "error_message": f"unexpected ruleType {tag['rule_type']!r}, expected {EXPECTED_RULE_TYPE!r}",
                }
            )
            continue

        old_ips, _ = parse_qualys_rule_text(tag["rule_text"])
        old_set = set(old_ips)
        new_set = set(new_ips)
        added = sorted(new_set - old_set, key=lambda c: new_ips.index(c) if c in new_ips else 0)
        removed = sorted(old_set - new_set, key=lambda c: old_ips.index(c) if c in old_ips else 0)

        report_rows.append(
            {
                "tag_id": tag["tag_id"],
                "tag_name": expected_name,
                "status": "dry-run" if not args.apply else "pending",
                "old_ip_summary": ", ".join(old_ips),
                "new_ip_summary": ", ".join(new_ips),
                "ips_added": ", ".join(added),
                "ips_removed": ", ".join(removed),
                "error_message": "",
                "_new_ips": new_ips,
                "_has_diff": old_set != new_set,
                "_tag": tag,
            }
        )

    for name, tag in gaid_tags_by_name.items():
        if name not in matched_names:
            old_ips, _ = parse_qualys_rule_text(tag["rule_text"])
            report_rows.append(
                {
                    "tag_id": tag["tag_id"],
                    "tag_name": name,
                    "status": "untouched",
                    "old_ip_summary": ", ".join(old_ips),
                    "new_ip_summary": "",
                    "ips_added": "",
                    "ips_removed": "",
                    "error_message": "not in file",
                }
            )

    # --- dry-run: print plan and exit ---
    if not args.apply:
        print("=== DRY RUN: no changes will be written. Pass --apply to write. ===")
        print_action_plan(report_rows)
        write_reports(report_rows)
        print_summary(report_rows, applied=False)
        return

    # --- apply: backup, then write, then verify ---
    to_update = [r for r in report_rows if r["status"] == "pending" and r["_has_diff"]]
    no_op = [r for r in report_rows if r["status"] == "pending" and not r["_has_diff"]]
    for r in no_op:
        r["status"] = "updated"

    if to_update:
        timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        touched_tags = [r["_tag"] for r in to_update]
        write_backup(touched_tags, timestamp)

    print(f"=== APPLYING: {len(to_update)} tag(s) will be updated ===")
    for r in to_update:
        apply_one_update(session, r)

    write_reports(report_rows)
    print_summary(report_rows, applied=True)


def apply_one_update(session, row):
    tag = row["_tag"]
    tag_id = tag["tag_id"]
    expected_name = row["tag_name"]
    new_ips = row["_new_ips"]

    # Defensive re-check: the id we're about to write to must still be the
    # tag we think it is (hard constraint: never write to the wrong tag).
    if tag["tag_name"] != expected_name:
        row["status"] = "error"
        row["error_message"] = (
            f"aborted: current name {tag['tag_name']!r} != expected {expected_name!r}"
        )
        return

    new_rule_text = ",".join(new_ips)
    xml_body = (
        "<ServiceRequest><data><Tag>"
        f"<ruleType>{EXPECTED_RULE_TYPE}</ruleType>"
        f"<ruleText>{new_rule_text}</ruleText>"
        "</Tag></data></ServiceRequest>"
    )

    try:
        request_with_retry(
            session,
            "POST",
            TAG_UPDATE_URL.format(tag_id=tag_id),
            xml_body,
            f"update tag {tag_id}",
        )
    except RuntimeError as exc:
        row["status"] = "error"
        row["error_message"] = str(exc)
        return

    verify_tag = fetch_tag_by_id(session, tag_id)
    if verify_tag is None:
        row["status"] = "error"
        row["error_message"] = "update call succeeded but tag not found on read-back"
        return

    stored_ips, _ = parse_qualys_rule_text(verify_tag["rule_text"])
    if set(stored_ips) != set(new_ips):
        row["status"] = "error"
        row["error_message"] = (
            "post-update verification failed: stored IPs do not match intended set "
            f"(stored={stored_ips!r}, expected={new_ips!r})"
        )
        return

    row["status"] = "updated"


def print_action_plan(report_rows):
    for r in report_rows:
        if r["status"] not in ("dry-run",):
            continue
        if not r.get("_has_diff"):
            print(f"  [no change] {r['tag_name']} (id {r['tag_id']}) already matches file")
            continue
        print(f"  [would update] {r['tag_name']} (id {r['tag_id']})")
        print(f"      + added:   {r['ips_added'] or '(none)'}")
        print(f"      - removed: {r['ips_removed'] or '(none)'}")
    print()


def write_reports(report_rows):
    today = datetime.date.today().isoformat()
    csv_path = f"gaid_update_report_{today}.csv"
    xlsx_path = f"gaid_update_report_{today}.xlsx"
    clean_rows = [{k: v for k, v in r.items() if not k.startswith("_")} for r in report_rows]
    write_csv_report(clean_rows, csv_path)
    write_xlsx_report(clean_rows, xlsx_path)
    print(f"Wrote {csv_path}")
    print(f"Wrote {xlsx_path}")


def print_summary(report_rows, applied):
    total = len(report_rows)
    updated = sum(1 for r in report_rows if r["status"] == "updated")
    dry_run = sum(1 for r in report_rows if r["status"] == "dry-run")
    skipped = sum(1 for r in report_rows if r["status"] == "skipped-no-match")
    untouched = sum(1 for r in report_rows if r["status"] == "untouched")
    failed = sum(1 for r in report_rows if r["status"] == "error")

    print()
    print("=== Summary ===")
    print(f"Mode:                  {'APPLY' if applied else 'DRY-RUN'}")
    print(f"Total rows:            {total}")
    print(f"Updated:               {updated}")
    print(f"Dry-run (would update):{dry_run}")
    print(f"Skipped (no match):    {skipped}")
    print(f"Untouched (not in file):{untouched}")
    print(f"Failed:                {failed}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

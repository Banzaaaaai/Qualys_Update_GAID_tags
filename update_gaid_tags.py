"""Reconcile GAID Asset Management tags in Qualys (EU2 pod) with a CMDB Excel export.

The Excel file is authoritative. For each "[VFZ] GAID: <n>" tag the script
updates, converts, creates or deletes so the tenant matches the file.

Credentials are never read from argv/hardcoded; they come from the
QUALYS_USERNAME / QUALYS_PASSWORD env vars, falling back to
qualys_creds.txt next to this script (format: "key:\\tvalue" per line).

WHAT IT DOES (see README.md for the rationale behind each rule):
  * In file + in tenant (NETWORK_RANGE) -> ruleText replaced with EXACTLY the
    file's IP set. Complete replacement, never a merge.
  * In file + in tenant (static)        -> converted in place to a dynamic
    NETWORK_RANGE tag, same id and name.
  * In file + not in tenant             -> created under
    "[VFZ] Global Application Inventory".
  * Not in file, or every row for the GAID is Out of Service -> deleted.
  * NAME_CONTAINS tags are never written to; other unexpected rule types are
    refused rather than guessed at.

Addresses in 169.254.0.0/16 and 192.168.0.0/16 are stripped from the file's
IP set (see EXCLUDED_IP_NETWORKS). The filter is applied to the file only,
never to what Qualys holds, so any such addresses already stored in a tag
show up as removals and get cleaned out.

HARD CONSTRAINTS:
1. COMPLETE REPLACEMENT of each GAID tag's IP/range list with the set from
   the Excel file -- never a merge/union with what is currently in Qualys.
2. EXISTING TAGS ARE UPDATED IN PLACE. The tag keeps its
   id/name/parent/color/criticality -- only ruleText (and, for a static tag,
   ruleType) changes. Never delete+recreate, because Qualys access scope is
   bound to tag identity.

RUN MODE: this script performs the REAL run by default -- it writes to
Qualys, deletions included. Pass --dry-run to preview without writing.
Every tag that will be modified or deleted is backed up to
backup_gaid_tags_<UTC>.json/.xlsx before the first write call, and every
write is verified by reading the tag back.
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

from xml.sax.saxutils import escape as xml_escape

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
TAG_CREATE_URL = f"{BASE_URL}/qps/rest/2.0/create/am/tag"
TAG_DELETE_URL = f"{BASE_URL}/qps/rest/2.0/delete/am/tag/{{tag_id}}"

# Confirmed 2026-09-12 against a live tag ("[VFZ] GAID: 1356", id 189526283):
# name is exactly this prefix + the GAID number, parent is the fixed tag
# below, and the rule is a NETWORK_RANGE whose ruleText is a single
# comma-separated list mixing bare IPv4 addresses and "A-B" hyphen ranges.
GAID_TAG_PREFIX = "[VFZ] GAID: "
GAID_PARENT_TAG_NAME = "[VFZ] Global Application Inventory"
EXPECTED_RULE_TYPE = "NETWORK_RANGE"

# Rule types that are never written to, whatever the file says. A
# NAME_CONTAINS tag matches assets by hostname pattern (one in this tenant
# holds a regex); overwriting it with an IP list would silently destroy that
# rule, so it is reported and skipped.
NEVER_UPDATE_RULE_TYPES = {"NAME_CONTAINS"}

# A static tag has no rule at all -- Qualys returns either no ruleType
# element (empty string here) or "STATIC". When the file supplies IPs for
# such a GAID, the tag is converted in place to a dynamic NETWORK_RANGE tag;
# when it does not, the tag is left alone.
STATIC_RULE_TYPES = {"", "STATIC"}

# Attributes given to newly created GAID tags, mirroring the convention of
# tags that already exist in the tenant (e.g. "Toolbox (GAID: 1356)"). The
# parent tag id is resolved at runtime by looking up GAID_PARENT_TAG_NAME,
# so it is never hardcoded/stale.
#
# Color is deliberately left empty. Existing GAID tags in this tenant store
# "#FF", but the live tag.xsd restricts color to #RGB or #RRGGBB
# (pattern "#[0-9a-fA-F]{3}([0-9a-fA-F]{3})?"), so sending "#FF" back would
# be rejected -- that stored value predates or bypasses validation. Rather
# than invent a color the tenant never chose, the element is omitted and
# Qualys applies its own default. Set this to a valid 3- or 6-digit hex
# value (e.g. "#FFFFFF") if new tags should have a specific color.
NEW_TAG_COLOR = ""
NEW_TAG_DESCRIPTION_TEMPLATE = "{asset} (GAID: {gaid})"
ASSET_NAME_COLUMN_PATTERNS = [r"^\s*asset\s*$", r"asset\s*name", r"application\s*name"]

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
    "file_resource_status",
    "old_ip_summary",
    "new_ip_summary",
    "ips_added",
    "ips_removed",
    "ips_excluded",
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

# A row's IP only counts toward a GAID's new set if this column (when
# present) has one of these values (case-insensitive). Rows whose resource
# is e.g. "Out of Service" or "Planned Decommission" are excluded -- an
# explicit choice confirmed against a real CMDB export on 2026-09-12, where
# ~12% of IP-bearing rows were "Out of Service". If no column matches these
# patterns, filtering is skipped (all rows included) and a warning is
# printed.
RESOURCE_STATUS_COLUMN_PATTERNS = [r"resource\s*status"]
INCLUDE_RESOURCE_STATUSES = {"in service"}

# A GAID that is present in the file is treated as decommissioned -- and its
# tag deleted -- only when EVERY one of its rows carries one of these
# statuses. Requiring every row (not just the IP-bearing ones) is deliberate:
# in a real export, 7 GAIDs had all their IP-bearing rows "Out of Service"
# while still having "In Service" resources that simply carry no IP address
# (databases, load balancers). Those applications are live, and deleting
# their tags would have been wrong and irreversible. A blank/unknown status
# is not in this set, so it also blocks deletion.
DELETE_ON_RESOURCE_STATUSES = {"out of service"}

# Tokens are split out of a cell (and across rows for the same GAID) on any
# of these separators.
IP_CELL_SPLIT_RE = re.compile(r"[,;\n\r]+")

# Safety cap on how many addresses a single "A-B" range may expand to, so a
# garbage/typo'd range (e.g. a full /8) can't blow up memory -- real GAID
# ranges observed in the tenant are at most a few hundred addresses wide.
MAX_RANGE_SIZE = 1_000_000

# Address blocks that are never scoped into a GAID tag. 169.254.0.0/16 is
# link-local/APIPA and 192.168.0.0/16 is re-used verbatim across sites, so
# neither identifies a real asset in this estate. Addresses in these blocks
# are stripped from the file's IP set; they are NOT stripped from what
# Qualys currently holds, so any that are already stored show up as
# removals in the diff and get cleaned out on the next run.
EXCLUDED_IP_NETWORKS = ("169.254.0.0/16", "192.168.0.0/16")
_EXCLUDED_IP_RANGES = tuple(
    (int(net.network_address), int(net.broadcast_address))
    for net in (ipaddress.IPv4Network(c) for c in EXCLUDED_IP_NETWORKS)
)


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
                "description": get_text(tag_elem, "description"),
                "created": get_text(tag_elem, "created"),
                "modified": get_text(tag_elem, "modified"),
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
        f'<Criteria field="id" operator="EQUALS">{xml_escape(str(tag_id))}</Criteria>'
        "</filters>"
        "</ServiceRequest>"
    )
    root = request_with_retry(session, "POST", TAG_SEARCH_URL, xml_body, f"verify tag {tag_id}")
    data_elem = find_child(root, "data")
    tags = parse_tag_elements(data_elem)
    return tags[0] if tags else None


def fetch_tag_by_name(session, name):
    xml_body = (
        "<ServiceRequest>"
        "<filters>"
        f'<Criteria field="name" operator="EQUALS">{xml_escape(name)}</Criteria>'
        "</filters>"
        "</ServiceRequest>"
    )
    root = request_with_retry(session, "POST", TAG_SEARCH_URL, xml_body, f"verify tag {name!r}")
    data_elem = find_child(root, "data")
    tags = parse_tag_elements(data_elem)
    return tags[0] if tags else None


# --------------------------------------------------------------------------
# IP / range normalization
# --------------------------------------------------------------------------


def expand_entry_to_ints(token):
    """Validate one token as an IPv4 address or 'A-B' range.

    Returns the set of individual address integers it covers, or raises
    ValueError. Expanding to individual addresses (rather than keeping the
    token as an opaque string) is what lets us correctly compare a tag whose
    current ruleText uses compressed ranges against a CMDB export that lists
    the same hosts as separate single-IP rows -- same address set, different
    surface form.
    """
    token = token.strip()
    if not token:
        raise ValueError("empty entry")

    if "-" in token:
        start_s, _, end_s = token.partition("-")
        start = ipaddress.IPv4Address(start_s.strip())
        end = ipaddress.IPv4Address(end_s.strip())
        if int(end) < int(start):
            raise ValueError(f"range end before start: {token}")
        if int(end) - int(start) + 1 > MAX_RANGE_SIZE:
            raise ValueError(f"range too large ({token})")
        return set(range(int(start), int(end) + 1))

    return {int(ipaddress.IPv4Address(token))}


def compact_ints_to_entries(int_set):
    """Merge a set of address integers into the minimal sorted list of
    single-IP / 'A-B' range strings covering exactly that set."""
    if not int_set:
        return []
    ordered = sorted(int_set)
    runs = []
    start = prev = ordered[0]
    for n in ordered[1:]:
        if n == prev + 1:
            prev = n
            continue
        runs.append((start, prev))
        start = prev = n
    runs.append((start, prev))

    entries = []
    for a, b in runs:
        if a == b:
            entries.append(str(ipaddress.IPv4Address(a)))
        else:
            entries.append(f"{ipaddress.IPv4Address(a)}-{ipaddress.IPv4Address(b)}")
    return entries


def normalize_ip_set(raw_tokens):
    """Validate every token, union them at the individual-address level, and
    recompact into the minimal sorted range representation.

    Returns (canonical_sorted_list, invalid_tokens). Recompacting from
    scratch (rather than trusting each token's own single/range shape) means
    the result is identical regardless of whether the input expressed a
    given set of hosts as one range or as many single IPs.
    """
    all_ints = set()
    invalid = []
    for token in raw_tokens:
        if token is None:
            continue
        token = str(token).strip()
        if not token:
            continue
        try:
            all_ints |= expand_entry_to_ints(token)
        except ValueError:
            invalid.append(token)
            continue

    ordered = compact_ints_to_entries(all_ints)
    return ordered, invalid


def drop_excluded_networks(entries):
    """Strip addresses in EXCLUDED_IP_NETWORKS from canonical IP entries.

    Applied to the file's IP set only -- never to what Qualys currently
    holds, so addresses already stored in a tag surface as removals in the
    diff instead of being masked on both sides.

    Returns (kept_entries, dropped_entries), both recompacted.
    """
    kept, dropped = set(), set()
    for entry in entries:
        for addr in expand_entry_to_ints(entry):
            if any(lo <= addr <= hi for lo, hi in _EXCLUDED_IP_RANGES):
                dropped.add(addr)
            else:
                kept.add(addr)
    return compact_ints_to_entries(kept), compact_ints_to_entries(dropped)


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
    status_col = find_column(headers, RESOURCE_STATUS_COLUMN_PATTERNS)
    asset_col = find_column(headers, ASSET_NAME_COLUMN_PATTERNS)

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
    if asset_col is not None:
        print(
            f"  Asset name column: {headers[asset_col]!r} (index {asset_col}) "
            "-- used for the description of newly created tags."
        )
    else:
        print(
            "  WARNING: no asset-name column matched "
            f"{ASSET_NAME_COLUMN_PATTERNS!r}; new tags will be created without "
            "a description."
        )
    if status_col is not None:
        print(
            f"  Resource status column: {headers[status_col]!r} (index {status_col}) "
            f"-- only rows with status in {sorted(INCLUDE_RESOURCE_STATUSES)!r} "
            "contribute IPs; others are excluded."
        )
    else:
        print(
            "  WARNING: no resource-status column matched "
            f"{RESOURCE_STATUS_COLUMN_PATTERNS!r}; all rows will contribute "
            "IPs regardless of status."
        )
    print()

    return sheet_name, headers, rows, gaid_col, ip_col, status_col, asset_col


def collect_file_gaids(rows, gaid_col, asset_col=None):
    """Every GAID appearing in the file's GAID column, regardless of whether
    it has IP rows or survives the status filter, plus its asset name.

    Deletion keys off this set -- "absent from the file" must mean the GAID
    number is nowhere in the export, NOT merely that it contributed no IPs.
    """
    all_gaids = set()
    asset_by_gaid = {}
    for row in rows:
        gaid_key = normalize_gaid_key(row[gaid_col])
        if gaid_key is None:
            continue
        all_gaids.add(gaid_key)
        if asset_col is not None and gaid_key not in asset_by_gaid:
            asset = row[asset_col]
            if asset is not None and str(asset).strip():
                asset_by_gaid[gaid_key] = str(asset).strip()
    return all_gaids, asset_by_gaid


def build_gaid_ip_map(rows, gaid_col, ip_col, status_col=None):
    """Group raw IP tokens by GAID, across however many rows each GAID has.

    Rows are excluded from contributing IPs when status_col is given and the
    row's value there is not in INCLUDE_RESOURCE_STATUSES (case-insensitive).
    Returns (raw_by_gaid, excluded_row_count, excluded_ip_token_count,
    fully_filtered_gaids) -- the last is the set of GAIDs that had at least
    one IP-bearing row in the file, but every single one was excluded by the
    status filter, so they carry zero IPs into the update logic even though
    they're technically "in the file".
    """
    raw_by_gaid = {}
    seen_any_ip_row = set()
    excluded_rows = 0
    excluded_tokens = 0
    for row in rows:
        gaid_key = normalize_gaid_key(row[gaid_col])
        if gaid_key is None:
            continue
        cell = row[ip_col]
        if cell is None:
            continue
        seen_any_ip_row.add(gaid_key)

        if status_col is not None:
            status_value = row[status_col]
            status_norm = str(status_value).strip().lower() if status_value is not None else ""
            if status_norm not in INCLUDE_RESOURCE_STATUSES:
                excluded_rows += 1
                excluded_tokens += len(IP_CELL_SPLIT_RE.split(str(cell)))
                continue

        tokens = IP_CELL_SPLIT_RE.split(str(cell))
        raw_by_gaid.setdefault(gaid_key, []).extend(tokens)

    fully_filtered_gaids = seen_any_ip_row - set(raw_by_gaid.keys())
    return raw_by_gaid, excluded_rows, excluded_tokens, fully_filtered_gaids


def find_decommissioned_gaids(rows, gaid_col, status_col):
    """GAIDs whose EVERY row is in DELETE_ON_RESOURCE_STATUSES.

    All rows count here, including rows with no IP address -- a single
    still-in-service resource anywhere under a GAID means the application is
    alive and its tag must not be deleted.
    """
    if status_col is None:
        return set()

    statuses_by_gaid = {}
    for row in rows:
        gaid_key = normalize_gaid_key(row[gaid_col])
        if gaid_key is None:
            continue
        status_value = row[status_col]
        label = str(status_value).strip().lower() if status_value is not None else ""
        statuses_by_gaid.setdefault(gaid_key, set()).add(label)

    return {
        gaid_key
        for gaid_key, labels in statuses_by_gaid.items()
        if labels and labels <= DELETE_ON_RESOURCE_STATUSES
    }


def build_status_summary(rows, gaid_col, ip_col, status_col):
    """Per GAID, a readable breakdown of RESOURCE STATUS across the rows that
    actually carry an IP -- e.g. "In Service: 5, Out of Service: 2".

    Only IP-bearing rows are counted, because rows without an IP cannot
    affect the tag's address set either way; counting them would just
    obscure why a GAID's IPs changed (or why it ended up untouched).
    """
    if status_col is None:
        return {}

    counts_by_gaid = {}
    for row in rows:
        gaid_key = normalize_gaid_key(row[gaid_col])
        if gaid_key is None or row[ip_col] is None:
            continue
        status_value = row[status_col]
        label = str(status_value).strip() if status_value is not None else "(blank)"
        counts = counts_by_gaid.setdefault(gaid_key, {})
        counts[label] = counts.get(label, 0) + 1

    summary = {}
    for gaid_key, counts in counts_by_gaid.items():
        ordered = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
        summary[gaid_key] = ", ".join(f"{label}: {n}" for label, n in ordered)
    return summary


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
    cols = [
        "tag_id",
        "tag_name",
        "parent_tag_id",
        "rule_type",
        "rule_text",
        "color",
        "criticality",
        "description",
        "created",
        "modified",
    ]
    ws.append(cols)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for t in touched_tags:
        ws.append([t.get(c, "") for c in cols])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"
    widths = {
        "tag_id": 12,
        "tag_name": 24,
        "parent_tag_id": 14,
        "rule_type": 16,
        "rule_text": 80,
        "color": 10,
        "criticality": 12,
        "description": 40,
        "created": 22,
        "modified": 22,
    }
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


REPORT_COLUMN_WIDTHS = {
    "tag_id": 12,
    "tag_name": 24,
    "status": 20,
    "file_resource_status": 38,
    "old_ip_summary": 60,
    "new_ip_summary": 60,
    "ips_added": 40,
    "ips_removed": 40,
    "ips_excluded": 40,
    "error_message": 44,
}
WRAP_COLUMNS = (
    "file_resource_status",
    "old_ip_summary",
    "new_ip_summary",
    "ips_added",
    "ips_removed",
    "ips_excluded",
    "error_message",
)


def _add_report_sheet(wb, title, rows, columns=None):
    """Append one styled sheet: bold/frozen header, autofilter, widths."""
    columns = columns or REPORT_COLUMNS
    ws = wb.create_sheet(title)

    ws.append(columns)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for r in rows:
        ws.append([r.get(col, "") for col in columns])

    wrap_idx = {columns.index(c) + 1 for c in WRAP_COLUMNS if c in columns}
    for row in range(2, ws.max_row + 1):
        for col in wrap_idx:
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"
    for idx, col in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = REPORT_COLUMN_WIDTHS.get(col, 15)
    return ws


def _split_rows_by_bucket(rows):
    """Split the flat report into the sheets a reviewer actually works from.

    Dry-run rows are split by whether there is a real address-level diff,
    because "matched and already correct" and "matched and will change" are
    completely different review tasks.
    """
    dry = [r for r in rows if r["status"] == "dry-run-update"]
    return {
        "Would Update": [r for r in dry if r.get("ips_added") or r.get("ips_removed")],
        "No Change": [r for r in dry if not (r.get("ips_added") or r.get("ips_removed"))],
        "Would Convert": [r for r in rows if r["status"] == "dry-run-convert"],
        "Would Create": [r for r in rows if r["status"] == "dry-run-create"],
        "Would Delete": [r for r in rows if r["status"] == "dry-run-delete"],
        "Updated": [r for r in rows if r["status"] == "updated"],
        "Converted": [r for r in rows if r["status"] == "converted"],
        "Created": [r for r in rows if r["status"] == "created"],
        "Deleted": [r for r in rows if r["status"] == "deleted"],
        "Skipped NAME_CONTAINS": [r for r in rows if r["status"] == "skipped-name-contains"],
        "Errors": [r for r in rows if r["status"] == "error"],
        "Skipped No Match": [r for r in rows if r["status"] == "skipped-no-match"],
        "Untouched": [r for r in rows if r["status"] == "untouched"],
    }


def write_xlsx_report(rows, path, run_meta=None):
    wb = Workbook()
    wb.remove(wb.active)

    buckets = _split_rows_by_bucket(rows)

    # --- Summary sheet ---
    ws = wb.create_sheet("Summary")
    mode = (run_meta or {}).get("Run mode", "")
    is_dry = "DRY" in mode.upper()
    ws.append([f"Qualys GAID Tag Update - {'Dry Run' if is_dry else 'REAL RUN (changes applied)'} Report"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])

    for key, value in (run_meta or {}).items():
        ws.append([key, value])
    ws.append([])

    ws.append(["Outcome", "Tags", "What it means"])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    explanations = {
        "Would Update": "Matched a live tag; IP set differs -> would be rewritten on --apply",
        "No Change": "Matched a live tag; IP set already identical -> no API call needed",
        "Would Convert": (
            "Static tag with IPs in the file -> would become a dynamic "
            "NETWORK_RANGE tag on --apply"
        ),
        "Converted": "Static tag converted in place to dynamic NETWORK_RANGE and verified",
        "Skipped NAME_CONTAINS": (
            "Hostname-pattern tag -- never updated by policy; existing rule left intact"
        ),
        "Would Create": "GAID in file has no tag in tenant -> would be created on --apply",
        "Would Delete": (
            "GAID absent from file entirely -> would be DELETED "
            "(needs --apply --allow-delete)"
        ),
        "Updated": "Rewritten in place in Qualys and verified by read-back",
        "Created": "Newly created in Qualys and verified by read-back",
        "Deleted": "Permanently deleted from Qualys and verified gone",
        "Errors": "Refused: unexpected ruleType or malformed IP content; never written",
        "Skipped No Match": "GAID in file but no valid IPs; no empty tag created",
        "Untouched": "GAID is in the file but contributed no IPs; left alone",
    }
    for name, bucket_rows in buckets.items():
        if not bucket_rows:
            continue
        ws.append([name, len(bucket_rows), explanations.get(name, "")])

    ws.append([])
    ws.append(["Total rows", len(rows)])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 76

    # --- One sheet per non-empty outcome, then the full flat report ---
    for name, bucket_rows in buckets.items():
        if not bucket_rows:
            continue
        _add_report_sheet(wb, name, bucket_rows)

    _add_report_sheet(wb, "All Rows", rows)

    wb.save(path)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("excel_path", help="Path to the CMDB Excel export")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help=(
            "Preview only: compute and report every change but make NO write "
            "calls to Qualys. Without this flag the script performs the real "
            "run -- updates, conversions, creates AND deletions."
        ),
    )
    # Accepted so the previous invocation still works; the real run is now
    # the default, so neither flag changes anything.
    parser.add_argument("--apply", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--allow-delete", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args()


def main():
    args = parse_args()

    # The real run is the default: writes happen unless --dry-run is given.
    # The rest of the script still reads args.apply / args.allow_delete, so
    # derive them here rather than threading a new flag through everything.
    args.apply = not args.dry_run
    args.allow_delete = not args.dry_run

    # One timestamp per run, shared by this run's backup and its report so
    # the two files are trivially paired after the fact.
    run_timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    if args.apply:
        print("*" * 72)
        print("REAL RUN - changes WILL be written to Qualys, including DELETIONS.")
        print("Every tag that will be modified or deleted is backed up first.")
        print("Use --dry-run to preview without writing.")
        print("*" * 72)
        print()

    print(f"=== Discovery: {args.excel_path} ===")
    _, _, rows, gaid_col, ip_col, status_col, asset_col = discover_excel(args.excel_path)
    raw_by_gaid, excluded_rows, excluded_tokens, fully_filtered_gaids = build_gaid_ip_map(
        rows, gaid_col, ip_col, status_col
    )
    all_gaids_in_file, asset_by_gaid = collect_file_gaids(rows, gaid_col, asset_col)
    status_summary_by_gaid = build_status_summary(rows, gaid_col, ip_col, status_col)
    decommissioned_gaids = find_decommissioned_gaids(rows, gaid_col, status_col)
    print(
        f"Found {len(all_gaids_in_file)} distinct GAID(s) in the file "
        f"({len(raw_by_gaid)} of them contribute IPs)."
    )
    if status_col is not None:
        print(
            f"Excluded {excluded_rows} row(s) ({excluded_tokens} IP token(s)) "
            "for not being in an included resource status."
        )
        if fully_filtered_gaids:
            print(
                f"NOTE: {len(fully_filtered_gaids)} GAID(s) had IP rows in the "
                "file but ALL were filtered out by resource status, so they "
                "carry zero IPs and are treated as 'untouched' (not cleared): "
                + ", ".join(f"{GAID_TAG_PREFIX}{g}" for g in sorted(fully_filtered_gaids))
            )
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

    run_meta = {
        "Run mode": "REAL RUN (writes to Qualys)" if args.apply else "DRY RUN (no writes)",
        "Run at (UTC)": datetime.datetime.now(datetime.timezone.utc).strftime(
            "%Y-%m-%d %H:%M:%S UTC"
        ),
        "Source file": os.path.basename(args.excel_path),
        "Qualys pod": BASE_URL,
        "Tag name format": f"{GAID_TAG_PREFIX}<number>",
        "Parent tag": GAID_PARENT_TAG_NAME,
        "Expected ruleType": EXPECTED_RULE_TYPE,
        "Never updated (ruleType)": ", ".join(sorted(NEVER_UPDATE_RULE_TYPES)),
        "Static tags": "converted to dynamic NETWORK_RANGE when the file has IPs, else untouched",
        "Excluded IP blocks": ", ".join(EXCLUDED_IP_NETWORKS),
        "Resource status filter": (
            ", ".join(sorted(INCLUDE_RESOURCE_STATUSES)) if status_col is not None else "(none)"
        ),
        "Deletion enabled": "YES (--allow-delete)" if args.allow_delete else "no",
        "Deletes when": (
            "GAID absent from file, OR every row for the GAID is "
            f"{'/'.join(sorted(DELETE_ON_RESOURCE_STATUSES))}"
        ),
        "GAIDs fully out of service in file": len(decommissioned_gaids),
        "Missing tags are created": "yes",
        "New tag description format": NEW_TAG_DESCRIPTION_TEMPLATE,
        "GAIDs in file (total)": len(all_gaids_in_file),
        "GAIDs in file (with IPs)": len(raw_by_gaid),
        "Rows excluded by status filter": excluded_rows,
        "IP tokens excluded by status filter": excluded_tokens,
        "GAIDs fully filtered out (left untouched)": len(fully_filtered_gaids),
        "Tags fetched from tenant": len(all_tags),
        "Existing GAID tags in tenant": len(gaid_tags_by_name),
    }

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

    excluded_ip_total = 0
    gaids_emptied_by_exclusion = set()

    for gaid_key, raw_tokens in sorted(raw_by_gaid.items(), key=lambda kv: (len(kv[0]), kv[0])):
        expected_name = f"{GAID_TAG_PREFIX}{gaid_key}"
        new_ips, invalid_tokens = normalize_ip_set(raw_tokens)
        had_ips_before_exclusion = bool(new_ips)
        new_ips, excluded_ips = drop_excluded_networks(new_ips)
        if excluded_ips:
            excluded_ip_total += 1
        if had_ips_before_exclusion and not new_ips:
            gaids_emptied_by_exclusion.add(gaid_key)

        tag = gaid_tags_by_name.get(expected_name)
        if tag is None:
            # In the file but not in the tenant -> create it.
            if invalid_tokens:
                report_rows.append(
                    {
                        "tag_id": "",
                        "tag_name": expected_name,
                        "status": "error",
                        "old_ip_summary": "",
                        "new_ip_summary": "",
                        "ips_added": "",
                        "ips_removed": "",
                        "error_message": "invalid IP/range entries: " + ", ".join(invalid_tokens),
                    }
                )
                continue
            if not new_ips:
                report_rows.append(
                    {
                        "tag_id": "",
                        "tag_name": expected_name,
                        "status": "skipped-no-match",
                        "old_ip_summary": "",
                        "new_ip_summary": "",
                        "ips_added": "",
                        "ips_removed": "",
                        "error_message": "no valid IPs in file; not creating an empty tag",
                    }
                )
                continue

            asset = asset_by_gaid.get(gaid_key, "")
            description = (
                NEW_TAG_DESCRIPTION_TEMPLATE.format(asset=asset, gaid=gaid_key) if asset else ""
            )
            report_rows.append(
                {
                    "tag_id": "",
                    "tag_name": expected_name,
                    "status": "create" if args.apply else "dry-run-create",
                    "old_ip_summary": "",
                    "new_ip_summary": ", ".join(new_ips),
                    "ips_added": ", ".join(new_ips),
                    "ips_removed": "",
                    "ips_excluded": ", ".join(excluded_ips),
                    "error_message": "",
                    "_new_ips": new_ips,
                    "_description": description,
                }
            )
            continue

        matched_names.add(expected_name)

        # An empty desired set is never written: clearing a tag's scope is a
        # far stronger action than updating it and must not happen as a side
        # effect of filtering out irrelevant address blocks.
        if not new_ips and not invalid_tokens:
            old_ips, _ = parse_qualys_rule_text(tag["rule_text"])
            report_rows.append(
                {
                    "tag_id": tag["tag_id"],
                    "tag_name": expected_name,
                    "status": "untouched",
                    "old_ip_summary": ", ".join(old_ips),
                    "new_ip_summary": "",
                    "ips_added": "",
                    "ips_removed": "",
                    "ips_excluded": ", ".join(excluded_ips),
                    "error_message": (
                        "all IPs for this GAID are in excluded blocks "
                        f"({', '.join(EXCLUDED_IP_NETWORKS)}); tag left as-is, not cleared"
                        if excluded_ips
                        else "no usable IPs in file; tag left as-is, not cleared"
                    ),
                }
            )
            continue

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

        rule_type = tag["rule_type"]

        # Hostname-pattern tags are never touched.
        if rule_type in NEVER_UPDATE_RULE_TYPES:
            report_rows.append(
                {
                    "tag_id": tag["tag_id"],
                    "tag_name": expected_name,
                    "status": "skipped-name-contains",
                    "old_ip_summary": tag["rule_text"],
                    "new_ip_summary": ", ".join(new_ips),
                    "ips_added": "",
                    "ips_removed": "",
                    "error_message": (
                        f"ruleType {rule_type!r} is never updated by policy; "
                        "its existing rule is left intact"
                    ),
                }
            )
            continue

        is_conversion = rule_type in STATIC_RULE_TYPES

        # Anything that is neither a network-range tag nor a static tag
        # awaiting conversion is refused rather than guessed at.
        if not is_conversion and rule_type != EXPECTED_RULE_TYPE:
            report_rows.append(
                {
                    "tag_id": tag["tag_id"],
                    "tag_name": expected_name,
                    "status": "error",
                    "old_ip_summary": tag["rule_text"],
                    "new_ip_summary": ", ".join(new_ips),
                    "ips_added": "",
                    "ips_removed": "",
                    "error_message": f"unexpected ruleType {rule_type!r}, expected {EXPECTED_RULE_TYPE!r}",
                }
            )
            continue

        old_ips, _ = parse_qualys_rule_text(tag["rule_text"])
        old_set = set(old_ips)
        new_set = set(new_ips)
        added = sorted(new_set - old_set, key=lambda c: new_ips.index(c) if c in new_ips else 0)
        removed = sorted(old_set - new_set, key=lambda c: old_ips.index(c) if c in old_ips else 0)

        if args.apply:
            row_status = "pending"
        elif is_conversion:
            row_status = "dry-run-convert"
        else:
            row_status = "dry-run-update"

        report_rows.append(
            {
                "tag_id": tag["tag_id"],
                "tag_name": expected_name,
                "status": row_status,
                "old_ip_summary": ", ".join(old_ips),
                "new_ip_summary": ", ".join(new_ips),
                "ips_added": ", ".join(added),
                "ips_removed": ", ".join(removed),
                "ips_excluded": ", ".join(excluded_ips),
                "error_message": (
                    f"static tag -> converting to dynamic {EXPECTED_RULE_TYPE}"
                    if is_conversion
                    else ""
                ),
                "_new_ips": new_ips,
                # A static tag has no rule at all, so converting it always
                # counts as a change even though there is nothing to diff.
                "_has_diff": old_set != new_set or is_conversion,
                "_is_conversion": is_conversion,
                "_tag": tag,
            }
        )

    for name, tag in gaid_tags_by_name.items():
        if name in matched_names:
            continue

        old_ips, _ = parse_qualys_rule_text(tag["rule_text"])
        gaid_key = name[len(GAID_TAG_PREFIX):]

        # Two independent reasons to delete: the GAID is nowhere in the file,
        # or it is in the file but every one of its resources is out of
        # service. A GAID that is in the file and still has live resources is
        # never a deletion candidate, even if it contributed no IPs.
        delete_reason = ""
        if gaid_key not in all_gaids_in_file:
            delete_reason = "GAID absent from file"
        elif gaid_key in decommissioned_gaids:
            delete_reason = "all resources Out of Service in file"

        if delete_reason:
            report_rows.append(
                {
                    "tag_id": tag["tag_id"],
                    "tag_name": name,
                    "status": "delete" if args.apply else "dry-run-delete",
                    "old_ip_summary": ", ".join(old_ips),
                    "new_ip_summary": "",
                    "ips_added": "",
                    "ips_removed": ", ".join(old_ips),
                    "error_message": delete_reason,
                    "_tag": tag,
                }
            )
            continue

        note = (
            "in file with live resources, but all IP-bearing rows are out of service"
            if gaid_key in fully_filtered_gaids
            else "in file, but no IP rows to apply"
        )
        report_rows.append(
            {
                "tag_id": tag["tag_id"],
                "tag_name": name,
                "status": "untouched",
                "old_ip_summary": ", ".join(old_ips),
                "new_ip_summary": "",
                "ips_added": "",
                "ips_removed": "",
                "error_message": note,
            }
        )

    if excluded_ip_total:
        print(
            f"Excluded IP blocks {', '.join(EXCLUDED_IP_NETWORKS)}: stripped addresses "
            f"from {excluded_ip_total} GAID(s)."
        )
        if gaids_emptied_by_exclusion:
            print(
                f"NOTE: {len(gaids_emptied_by_exclusion)} GAID(s) had ALL their IPs in "
                "excluded blocks; those tags are left as-is, not cleared: "
                + ", ".join(
                    f"{GAID_TAG_PREFIX}{g}" for g in sorted(gaids_emptied_by_exclusion)
                )
            )
        print()

    # Annotate every row with the file's resource-status breakdown for that
    # GAID, so a reviewer can see at a glance whether an IP change (or a
    # GAID being untouched) is explained by resources going out of service.
    for r in report_rows:
        gaid_key = r["tag_name"][len(GAID_TAG_PREFIX):]
        r["file_resource_status"] = status_summary_by_gaid.get(gaid_key, "")

    # --- dry-run: print plan and exit ---
    if not args.apply:
        print("=== DRY RUN: no changes will be written. Omit --dry-run for the real run. ===")
        print_action_plan(report_rows)
        write_reports(report_rows, run_meta, applied=False, run_timestamp=run_timestamp)
        print_summary(report_rows, applied=False)
        return

    # --- apply: backup, then write, then verify ---
    to_update = [r for r in report_rows if r["status"] == "pending" and r["_has_diff"]]
    no_op = [r for r in report_rows if r["status"] == "pending" and not r["_has_diff"]]
    for r in no_op:
        r["status"] = "updated"

    to_create = [r for r in report_rows if r["status"] == "create"]
    to_delete = [r for r in report_rows if r["status"] == "delete"]

    # Deleting is irreversible and detaches the tag from every asset it is
    # applied to, so it needs its own opt-in on top of --apply.
    if to_delete and not args.allow_delete:
        for r in to_delete:
            r["status"] = "untouched"
            r["ips_removed"] = ""
            r["error_message"] = (
                "GAID absent from file; deletion skipped (pass --allow-delete to delete)"
            )
        print(
            f"NOTE: {len(to_delete)} tag(s) qualify for deletion but --allow-delete "
            "was not passed; they were left untouched."
        )
        to_delete = []

    # Back up everything that is about to be modified or destroyed, before
    # the first write call.
    touched_tags = [r["_tag"] for r in to_update] + [r["_tag"] for r in to_delete]
    if touched_tags:
        write_backup(touched_tags, run_timestamp)

    parent_tag_id = ""
    if to_create:
        parent_tag = next(
            (t for t in all_tags if t["tag_name"] == GAID_PARENT_TAG_NAME), None
        )
        if parent_tag is None:
            raise RuntimeError(
                f"Cannot create tags: parent tag {GAID_PARENT_TAG_NAME!r} not found in tenant"
            )
        parent_tag_id = parent_tag["tag_id"]
        print(f"Parent tag for new tags: {GAID_PARENT_TAG_NAME!r} (id {parent_tag_id})")

    print(
        f"=== APPLYING: {len(to_update)} update(s), {len(to_create)} create(s), "
        f"{len(to_delete)} delete(s) ==="
    )
    for r in to_update:
        apply_one_update(session, r)
    for r in to_create:
        apply_one_create(session, r, parent_tag_id)
    # Recomputed from the same inputs that classified each row, so the
    # pre-delete guard re-checks the decision rather than trusting the row.
    deletable_gaids = decommissioned_gaids | {
        r["tag_name"][len(GAID_TAG_PREFIX):]
        for r in to_delete
        if r["tag_name"][len(GAID_TAG_PREFIX):] not in all_gaids_in_file
    }
    for r in to_delete:
        apply_one_delete(session, r, deletable_gaids)

    write_reports(report_rows, run_meta, applied=True, run_timestamp=run_timestamp)
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
        f"<ruleText>{xml_escape(new_rule_text)}</ruleText>"
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

    # The tag must end up dynamic/NETWORK_RANGE -- this is the only check
    # that proves a static tag actually converted rather than silently
    # keeping its old form.
    if verify_tag["rule_type"] != EXPECTED_RULE_TYPE:
        row["status"] = "error"
        row["error_message"] = (
            "post-update verification failed: ruleType is "
            f"{verify_tag['rule_type']!r}, expected {EXPECTED_RULE_TYPE!r}"
        )
        return

    row["status"] = "converted" if row.get("_is_conversion") else "updated"


def build_create_xml(name, parent_tag_id, rule_text, description):
    parts = [
        f"<name>{xml_escape(name)}</name>",
        f"<parentTagId>{xml_escape(str(parent_tag_id))}</parentTagId>",
        f"<ruleType>{EXPECTED_RULE_TYPE}</ruleType>",
        f"<ruleText>{xml_escape(rule_text)}</ruleText>",
    ]
    if NEW_TAG_COLOR:
        parts.append(f"<color>{xml_escape(NEW_TAG_COLOR)}</color>")
    if description:
        parts.append(f"<description>{xml_escape(description)}</description>")
    return "<ServiceRequest><data><Tag>" + "".join(parts) + "</Tag></data></ServiceRequest>"


def apply_one_create(session, row, parent_tag_id):
    """Create a missing GAID tag, then read it back and verify its IP set."""
    name = row["tag_name"]
    new_ips = row["_new_ips"]

    # Never create a tag that already exists -- a duplicate name would make
    # subsequent runs ambiguous about which id to update.
    existing = fetch_tag_by_name(session, name)
    if existing is not None:
        row["status"] = "error"
        row["error_message"] = (
            f"aborted: tag {name!r} already exists (id {existing['tag_id']}) -- not creating a duplicate"
        )
        return

    xml_body = build_create_xml(name, parent_tag_id, ",".join(new_ips), row.get("_description", ""))

    try:
        request_with_retry(session, "POST", TAG_CREATE_URL, xml_body, f"create tag {name!r}")
    except RuntimeError as exc:
        row["status"] = "error"
        row["error_message"] = str(exc)
        return

    verify_tag = fetch_tag_by_name(session, name)
    if verify_tag is None:
        row["status"] = "error"
        row["error_message"] = "create call succeeded but tag not found on read-back"
        return

    stored_ips, _ = parse_qualys_rule_text(verify_tag["rule_text"])
    if set(stored_ips) != set(new_ips):
        row["status"] = "error"
        row["error_message"] = (
            "post-create verification failed: stored IPs do not match intended set "
            f"(stored={stored_ips!r}, expected={new_ips!r})"
        )
        return

    row["tag_id"] = verify_tag["tag_id"]
    row["status"] = "created"


def apply_one_delete(session, row, deletable_gaids):
    """Delete a GAID tag that the file says should no longer exist.

    Deletion is irreversible and detaches the tag from every asset it is
    applied to, so the target is re-verified against live state immediately
    before the call: the id must still carry the expected name, and that
    name's GAID must still be in the approved deletion set.
    """
    tag_id = row["tag_id"]
    expected_name = row["tag_name"]

    live_tag = fetch_tag_by_id(session, tag_id)
    if live_tag is None:
        row["status"] = "error"
        row["error_message"] = "aborted: tag not found immediately before delete"
        return
    if live_tag["tag_name"] != expected_name:
        row["status"] = "error"
        row["error_message"] = (
            f"aborted: id {tag_id} now carries name {live_tag['tag_name']!r}, "
            f"expected {expected_name!r}"
        )
        return

    gaid_key = expected_name[len(GAID_TAG_PREFIX):]
    if gaid_key not in deletable_gaids:
        row["status"] = "error"
        row["error_message"] = (
            f"aborted: GAID {gaid_key} is no longer an approved deletion candidate"
        )
        return

    try:
        request_with_retry(
            session,
            "POST",
            TAG_DELETE_URL.format(tag_id=tag_id),
            "<ServiceRequest></ServiceRequest>",
            f"delete tag {tag_id}",
        )
    except RuntimeError as exc:
        row["status"] = "error"
        row["error_message"] = str(exc)
        return

    if fetch_tag_by_id(session, tag_id) is not None:
        row["status"] = "error"
        row["error_message"] = "delete call succeeded but tag still present on read-back"
        return

    row["status"] = "deleted"


def print_action_plan(report_rows):
    for r in report_rows:
        status = r["status"]
        if status == "dry-run-update":
            if not r.get("_has_diff"):
                print(f"  [no change] {r['tag_name']} (id {r['tag_id']}) already matches file")
                continue
            print(f"  [would update] {r['tag_name']} (id {r['tag_id']})")
            print(f"      + added:   {r['ips_added'] or '(none)'}")
            print(f"      - removed: {r['ips_removed'] or '(none)'}")
        elif status == "dry-run-convert":
            print(f"  [would CONVERT to dynamic] {r['tag_name']} (id {r['tag_id']})")
            print(f"      static tag -> {EXPECTED_RULE_TYPE} with: {r['new_ip_summary']}")
        elif status == "dry-run-create":
            print(f"  [would CREATE] {r['tag_name']}")
            print(f"      description: {r.get('_description', '') or '(none)'}")
            print(f"      IPs: {r['new_ip_summary']}")
        elif status == "dry-run-delete":
            print(f"  [would DELETE] {r['tag_name']} (id {r['tag_id']}) - GAID absent from file")
            print(f"      current IPs: {r['old_ip_summary'] or '(none)'}")
    print()


def _write_with_fallback(writer, path):
    """Write, falling back to a timestamped name if the file is locked.

    On Windows the previous report is often still open in Excel, which holds
    an exclusive lock. Losing a completed run's output to that would be
    daft, so write beside it instead.
    """
    try:
        writer(path)
        return path
    except PermissionError:
        stem, ext = os.path.splitext(path)
        stamp = datetime.datetime.now().strftime("%H%M%S")
        alt = f"{stem}_{stamp}{ext}"
        writer(alt)
        print(f"NOTE: {path} was locked (open in another program); wrote {alt} instead.")
        return alt


def write_reports(report_rows, run_meta=None, applied=False, run_timestamp=None):
    """Write the change report.

    A real run's report is stamped with that run's UTC timestamp -- the same
    one as its backup, so the pair is obvious -- and therefore can never be
    overwritten by a later run. It is the only record of what was changed,
    and changes here are irreversible. A dry-run report is just a preview, so
    it keeps the plain dated name and is allowed to be replaced.
    """
    if applied:
        stem = f"gaid_update_report_{run_timestamp}_applied"
    else:
        stem = f"gaid_update_report_{datetime.date.today().isoformat()}"

    clean_rows = [{k: v for k, v in r.items() if not k.startswith("_")} for r in report_rows]

    csv_path = _write_with_fallback(lambda p: write_csv_report(clean_rows, p), f"{stem}.csv")
    xlsx_path = _write_with_fallback(
        lambda p: write_xlsx_report(clean_rows, p, run_meta), f"{stem}.xlsx"
    )
    print(f"Wrote {csv_path}")
    print(f"Wrote {xlsx_path}")


def print_summary(report_rows, applied):
    count = lambda s: sum(1 for r in report_rows if r["status"] == s)
    dry_run_rows = [r for r in report_rows if r["status"] == "dry-run-update"]
    dry_run_real_diff = sum(1 for r in dry_run_rows if r.get("ips_added") or r.get("ips_removed"))

    print()
    print("=== Summary ===")
    print(f"Mode:                     {'APPLY' if applied else 'DRY-RUN'}")
    print(f"Total rows:               {len(report_rows)}")
    if applied:
        print(f"Updated:                  {count('updated')}")
        print(f"Converted to dynamic:     {count('converted')}")
        print(f"Created:                  {count('created')}")
        print(f"Deleted:                  {count('deleted')}")
    else:
        print(f"Would update (real diff): {dry_run_real_diff}")
        print(f"Matched, no change:       {len(dry_run_rows) - dry_run_real_diff}")
        print(f"Would convert to dynamic: {count('dry-run-convert')}")
        print(f"Would create:             {count('dry-run-create')}")
        print(f"Would DELETE:             {count('dry-run-delete')}")
    print(f"Skipped NAME_CONTAINS:    {count('skipped-name-contains')}")
    print(f"Skipped (no valid IPs):   {count('skipped-no-match')}")
    print(f"Untouched (in file, no IPs): {count('untouched')}")
    print(f"Failed:                   {count('error')}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

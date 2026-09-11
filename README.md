# Qualys GAID Tag Updater

Bulk-updates the IP/range content of `[VFZ] GAID: <n>` Asset Management tags
in Qualys (EU2 pod), using a CMDB Excel export as the authoritative source.

## What a GAID tag is

Confirmed against a live tag in the tenant (`[VFZ] GAID: 1356`, id
`189526283`) on 2026-09-12:

- Name: `[VFZ] GAID: <number>` — exact prefix `"[VFZ] GAID: "` (note the space
  after the colon).
- Parent: `[VFZ] Global Application Inventory` (tag id `175197329`, a static
  root tag).
- Rule: dynamic, `ruleType = NETWORK_RANGE`. `ruleText` is a single
  comma-separated list mixing bare IPv4 addresses and `A-B` hyphen ranges,
  e.g. `172.20.221.158-172.20.221.159,172.22.150.136,172.22.160.65-172.22.160.66`.
- **The only thing this tool ever changes is `ruleText`.** Name, parent,
  color, and criticality are left untouched.

**Tenant reality check:** as of 2026-09-12, 115 of the 464 existing
`[VFZ] GAID: *` tags do **not** have `ruleType = NETWORK_RANGE` (some are
`NAME_CONTAINS`, `CLOUD_ASSET`, or have no rule at all — likely static or
differently-configured tags). If a GAID in your Excel file matches one of
these, the script reports it as `error` ("unexpected ruleType") and never
touches it, rather than assuming every GAID tag is a network-range tag.

## The two hard constraints

1. **Complete replacement, not merge.** Each GAID tag's IP/range list is set
   to *exactly* the set found in the Excel file for that GAID. The old list
   is discarded entirely — nothing is unioned or appended. After every write,
   the tag is read back and the stored IP list is asserted to equal the
   intended set; a mismatch is reported as an `error` for that row.
2. **Update in place, never delete-and-recreate.** The script only ever
   calls `POST /qps/rest/2.0/update/am/tag/{id}` against the *existing* tag
   id, sending only `<ruleType>` (unchanged, resent because Qualys pairs it
   with `ruleText`) and the new `<ruleText>`. It never creates or deletes a
   tag. Before writing, it re-checks that the target id's current name still
   equals the expected `[VFZ] GAID: <n>` and aborts that row if not — Qualys
   access scope is bound to tag identity, so recreating a tag would silently
   break scoped access.

## Matching

- Excel GAID `<n>` is matched to the Qualys tag named exactly
  `[VFZ] GAID: <n>`.
- No matching tag in the tenant → **not created**; reported as
  `skipped-no-match`.
- Existing GAID tag not present in the Excel file → **not touched**;
  reported as `untouched`.
- Only tags named in the file that already exist in Qualys are ever written.

## Credential setup

Never pass credentials on the command line. The script reads them, in order:

1. `QUALYS_USERNAME` / `QUALYS_PASSWORD` environment variables, or
2. `qualys_creds.txt` next to the script, formatted as:
   ```
   user:	your_username
   pass:	your_password
   ```
   (tab or spaces after the colon both work). This file is git-ignored —
   never commit it.

## Usage

```
python update_gaid_tags.py <path_to_cmdb_export.xlsx>          # dry-run (default)
python update_gaid_tags.py <path_to_cmdb_export.xlsx> --apply  # writes to Qualys
```

### Dry-run (default, no `--apply`)

- Prints the Excel discovery output: sheet names, header row, 3 sample rows,
  and the inferred GAID/IP column mapping (see `GAID_COLUMN_PATTERNS` /
  `IP_COLUMN_PATTERNS` at the top of the script if your export's headers
  don't match what's inferred — adjust those lists and re-run).
- Fetches every tag in the tenant (paginated, same id-`GREATER` cursor
  approach as `fetch_qualys_tags.py`) and matches GAID tags by exact name.
- Computes and prints the per-tag diff (IPs added/removed) and the full
  action plan.
- Writes `gaid_update_report_<date>.csv` / `.xlsx`.
- **Makes no write calls to Qualys.**

### Apply (`--apply`)

- Re-runs discovery and diffing exactly as above.
- Before any write: dumps the full current state (id, name, parent, rule,
  color, criticality) of every tag that will actually be changed to
  `backup_gaid_tags_<UTC timestamp>.json` and `.xlsx`, for rollback.
- For each tag with a real diff: updates `ruleText` in place, then reads the
  tag back and asserts the stored IP set matches exactly. A tag whose new
  set equals its current set is left alone (no API call) and reported as
  `updated` with no diff.
- Writes the same dated CSV/XLSX report, now with final `updated` / `error`
  statuses.

## Excel input assumptions

The script auto-detects, per run (see console output to confirm):

- Which column holds the GAID number (header matching `GAID`,
  `Application ID`, `App ID`, etc.)
- Which column holds IP/range content (header matching `IP address`,
  `IP range`, `Network range`, or `IP`).
- Whether the file has **one row per GAID** with a delimited IP cell, or
  **multiple rows per GAID** (one IP/range per row). Both are handled by the
  same logic: every row's IP cell is split on commas/semicolons/newlines and
  grouped by GAID, so the layout doesn't need to be uniform.

Each resulting IP/range token is validated as a valid IPv4 address or an
`A-B` range (`A <= B`); the whole GAID's update is rejected (status
`error`, not partially applied) if any token is malformed.

## Report columns

`tag_id, tag_name, status, old_ip_summary, new_ip_summary, ips_added,
ips_removed, error_message`

Status values: `updated`, `dry-run`, `skipped-no-match`, `untouched`,
`error`.

## Rate limiting / retries

Same policy as `fetch_qualys_tags.py`: 5 attempts, exponential backoff
capped at 300s, honors `X-RateLimit-ToWait-Sec` then `Retry-After` on
`409`/`429`, retries timeouts/connection errors, and fails loud (exit 1) if
all retries are exhausted.

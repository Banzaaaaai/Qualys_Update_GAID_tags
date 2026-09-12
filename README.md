# Qualys GAID Tag Updater

Bulk-reconciles the `[VFZ] GAID: <n>` Asset Management tags in a Qualys
tenant (EU2 pod) against a CMDB Excel export, which is treated as the
authoritative source. Updates, converts, creates and deletes tags so the
tenant matches the file.

> **⚠ The default run is REAL.** Running the script with no flags writes to
> Qualys — updates, conversions, creates **and deletions**. Use `--dry-run`
> to preview.

---

## Usage

```bash
# REAL RUN (default): updates + conversions + creates + deletions
python update_gaid_tags.py asset-resource-owner.xlsx

# preview only: no write calls of any kind, produces the same report
python update_gaid_tags.py asset-resource-owner.xlsx --dry-run
```

- The Excel path is `argv[1]`.
- `--apply` and `--allow-delete` are still accepted so older invocations keep
  working, but they no longer change anything — the real run is the default.
- **Deletions are irreversible.** They detach the tag from every asset it is
  applied to and break any Qualys access scope bound to that tag's identity.
  For a static tag, the backup captures the definition but not its manual
  asset assignments, so those cannot be restored.
- Before changing anything the script writes
  `backup_gaid_tags_<UTC>.json` / `.xlsx` covering every tag it will modify
  or delete, and it verifies every write by reading the tag back.
- Run `--dry-run` first and read the report whenever the source file or the
  tenant has changed materially.

### Credentials

- Never passed on the command line and never hardcoded.
- Read from `QUALYS_USERNAME` / `QUALYS_PASSWORD` environment variables.
- Falling back to `qualys_creds.txt` next to the script:
  `user:<tab>your_username` / `pass:<tab>your_password` (one per line).
- That file is git-ignored — never commit it.

---

## What a GAID tag is

Confirmed against a live tag (`[VFZ] GAID: 1356`, id `189526283`):

- **Name:** `[VFZ] GAID: <number>` — exact prefix `"[VFZ] GAID: "`, note the
  space after the colon.
- **Parent:** `[VFZ] Global Application Inventory` (id `175197329`),
  resolved at runtime by name so it is never a stale hardcoded id.
- **Rule:** dynamic, `ruleType = NETWORK_RANGE`.
- **Content:** `ruleText` — one comma-separated list mixing bare IPv4
  addresses and `A-B` hyphen ranges, e.g.
  `172.20.221.158-172.20.221.159,172.22.150.136,172.22.160.65-172.22.160.66`.
- **`ruleText` is the only field this tool ever modifies on an existing tag.**
  Name, parent, color, criticality and description are left untouched.

Live rule-type distribution across the 464 GAID tags in the tenant:
`NETWORK_RANGE` 349, static/blank 113, `NAME_CONTAINS` 1, `CLOUD_ASSET` 1.

---

## Reading the Excel file

- **Sheet:** the first sheet is used; all sheet names are printed.
- **Header:** row 1. The header, three sample rows and the inferred column
  mapping are printed on every run for confirmation.
- **Column inference** (driven by the `*_COLUMN_PATTERNS` constants at the
  top of the script — adjust there if an export's headers differ):
  - GAID column — matches `GAID`, `Application ID`, `App ID`.
  - IP column — matches `IP address`, `IP range`, `Network range`, `IP`.
  - Resource status column — matches `RESOURCE STATUS`.
  - Asset name column — matches `ASSET`, `Asset name`, `Application name`;
    used only for the description of newly created tags.
- **Row layout is auto-detected and irrelevant to the result.** One row per
  GAID with a delimited IP cell, and many rows per GAID with one IP each,
  are handled by the same logic: every row's IP cell is split on commas,
  semicolons and newlines, then grouped by GAID.
- **GAID normalisation:** Excel floats (`1356.0`) and stray `.0` suffixes are
  normalised to `1356` so they match tag names.

### Resource status filter

- Only rows whose `RESOURCE STATUS` is in `INCLUDE_RESOURCE_STATUSES`
  (currently just `"In Service"`) contribute IP addresses.
- Rows marked `Out of Service` or `Planned Decommission` are dropped before
  a GAID's IP set is built. In the reference export this excluded 742 rows
  and 1,537 IP tokens (~12% of IP-bearing rows).
- If no status column is found, filtering is skipped and a warning is
  printed.
- A GAID whose rows all get filtered out is **not** treated as "clear this
  tag" — see the untouched rules below.

### IP validation and normalisation

- Every token must be a valid IPv4 address or an `A-B` range with `A <= B`.
- **If any token for a GAID is malformed, that whole GAID is rejected**
  (status `error`) rather than partially applied — no half-correct IP list
  is ever written.
- Tokens are expanded to individual addresses, de-duplicated, unioned, then
  recompacted into the minimal sorted set of ranges before writing.
- **Range vs single-IP formatting is not a difference.** Qualys stores
  ranges compressed (`10.0.0.1-10.0.0.3`); a CMDB that lists one IP per host
  row (`10.0.0.1`, `10.0.0.2`, `10.0.0.3`) describes the identical set. Both
  sides are expanded to individual addresses before comparison, so this
  produces no diff. Comparing the raw strings instead would have generated
  hundreds of false "changes" and written bloated, unranged rule text.
- A range wider than `MAX_RANGE_SIZE` (1,000,000 addresses) is rejected as
  garbage rather than expanded.

### Excluded address blocks

- Addresses in `EXCLUDED_IP_NETWORKS` — `169.254.0.0/16` (link-local/APIPA)
  and `192.168.0.0/16` (re-used verbatim across sites) — never identify a
  real asset in this estate and are stripped from the file's IP set.
- The filter applies to the **file only, never to what Qualys currently
  holds**. That is deliberate: filtering both sides would mask addresses
  already stored in a tag and they would never be cleaned up. Leaving the
  Qualys side unfiltered makes them appear as removals in the diff and they
  get stripped on the next run.
- Ranges straddling a block boundary are split, not dropped wholesale —
  `192.167.255.254-192.168.0.2` keeps `192.167.255.254-192.167.255.255` and
  drops the rest.
- The `ips_excluded` report column lists exactly what was stripped per GAID.
- **A GAID whose addresses are *all* excluded is left untouched, not
  cleared.** Wiping a tag's scope is a far stronger action than updating it
  and must never happen as a side effect of this filter.

---

## Decision logic

Each GAID is matched to the Qualys tag named exactly
`"[VFZ] GAID: " + <gaid>`. Existing tags are enumerated with the
`id`-`GREATER` cursor (`hasMoreRecords` / `lastId`) over
`POST /qps/rest/2.0/search/am/tag`.

### Update — in the file, in the tenant, `NETWORK_RANGE`

- `ruleText` is replaced with **exactly** the file's IP set.
- **Complete replacement, never a merge** — the old list is discarded
  entirely, never unioned or appended to.
- **Updated in place, never delete-and-recreate.** Only
  `POST /qps/rest/2.0/update/am/tag/{id}` against the existing id, sending
  only `<ruleType>` (unchanged — Qualys pairs it with `ruleText`) and the new
  `<ruleText>`. Tag id, name, parent, color and criticality are preserved,
  because Qualys access scope is bound to tag identity.
- Before writing, the target id's current name is re-checked against the
  expected `[VFZ] GAID: <n>`; a mismatch aborts that row.
- A tag whose set already matches is left alone with **no API call** and
  reported as `updated` with an empty diff.

### Convert — static tag with IPs in the file

- A static tag (`ruleType` blank or `STATIC`) whose GAID has IPs in the file
  is **converted in place to a dynamic `NETWORK_RANGE` tag**, keeping the
  same id and name.
- The payload is identical to a normal update; the post-write read-back
  additionally asserts `ruleType` really is `NETWORK_RANGE` afterwards,
  which is the only proof the conversion took effect rather than silently
  no-opping.

### Create — in the file, not in the tenant

- A new tag is created via `POST /qps/rest/2.0/create/am/tag` with:
  - `name` = `[VFZ] GAID: <n>`
  - `parentTagId` = id of `[VFZ] Global Application Inventory`, resolved by
    name at runtime
  - `ruleType` = `NETWORK_RANGE`, `ruleText` = the file's IP set
  - `color` — omitted by default. Existing GAID tags store `#FF`, but the
    live XSD restricts color to `#RGB`/`#RRGGBB`, so echoing `#FF` back
    would be rejected. Set `NEW_TAG_COLOR` to a valid 3- or 6-digit hex
    value if new tags should have a specific color.
  - `description` = `<ASSET> (GAID: <n>)`, mirroring existing tags such as
    `Toolbox (GAID: 1356)`
- A GAID with no valid IPs is **never created as an empty tag** — reported
  `skipped-no-match`.
- Before creating, the name is re-checked; if a tag with that name already
  exists the row aborts rather than creating a duplicate, which would make
  later runs ambiguous about which id to update.

### Delete — happens on a real run

Two independent reasons:

- **The GAID appears nowhere in the file's GAID column.**
- **Every row for that GAID is `Out of Service`**
  (`DELETE_ON_RESOURCE_STATUSES`) — the application is decommissioned.

Important qualifications:

- "Absent from the file" means the **GAID number is nowhere in the GAID
  column** — *not* merely that it contributed no IPs. In the reference
  export, 152 tenant tags had no IPs to apply but only 22 had a GAID
  genuinely missing from the file.
- The decommissioned rule requires **every** row, including rows with no IP
  address. Checking only IP-bearing rows would be wrong: 7 GAIDs had all
  their IP-bearing rows `Out of Service` while still having `In Service`
  resources that simply carry no IP (databases, load balancers). Those
  applications are live and their tags must not be deleted.
- A blank/unknown status is not in the delete set, so it also blocks
  deletion.
- Tag type does not protect a tag from deletion — deletion keys off the
  GAID, so a static tag whose GAID is absent from the file is still deleted.
- Immediately before each delete the script re-verifies that the id still
  carries the expected name **and** that the GAID is still in the approved
  deletion set, re-derived from the source data.

### Never updated

- **`NAME_CONTAINS` tags are never written to**, whatever the file says
  (`NEVER_UPDATE_RULE_TYPES`). These match assets by hostname pattern — one
  in this tenant holds a regex — and overwriting that with an IP list would
  silently destroy the rule. Reported as `skipped-name-contains` with the
  rule it retains.
- Any other unexpected rule type (e.g. `CLOUD_ASSET`) with IPs in the file
  is **refused** as an `error` rather than guessed at.

### Untouched

- A static tag whose GAID has **no** IPs in the file.
- A GAID in the file that contributed no IPs — either no IP rows at all, or
  every IP-bearing row filtered out by resource status — while still having
  live resources. These are **not** cleared: wiping a tag's scope is a
  stronger action than an update and deserves a human decision.

### Rule-type summary

| Existing `ruleType` | File has IPs | Action |
| --- | --- | --- |
| `NETWORK_RANGE` | yes | Update `ruleText` in place |
| blank / `STATIC` | yes | Convert to dynamic `NETWORK_RANGE` |
| blank / `STATIC` | no | Untouched |
| `NAME_CONTAINS` | either | Never updated — skipped |
| other (e.g. `CLOUD_ASSET`) | yes | Refused as `error` |

---

## Rule evaluation ("Evaluate Rule on Creation")

The UI checkbox has **no API equivalent, and none is needed** — Qualys does
it automatically for exactly the operations this script performs.

- The live `tag.xsd` on this pod has no settable evaluation field. It
  exposes only the read-only `reEvalStatus` and `reEvalStatusProgress`.
- The 392-page *Asset Management & Tagging API v2* guide contains no
  `reevaluateTagOnUpdate` / "evaluate rule" request parameter anywhere.
- The old explicit endpoint, `POST /qps/rest/2.0/evaluate/am/tag/<id>`, is
  **deprecated** and now returns `INVALID_REQUEST`:

  > "The Evaluate Tag API is now deprecated... now tags are automatically
  > queued for evaluation when their dynamic rule is updated or a new
  > dynamic tag is created."

So an updated `ruleText`, a static→dynamic conversion, and a newly created
dynamic tag are each queued for evaluation by the platform on write. Adding
an evaluation element to the request body would not enable anything; it
would risk the whole payload being rejected as invalid XML. Progress can be
observed afterwards via each tag's `reEvalStatus` field.

## Verification, backups and safety

- **Every write is verified by read-back.** After an update or conversion
  the tag is re-fetched and the stored IP set asserted equal to the intended
  set (and `ruleType` asserted to be `NETWORK_RANGE`); after a create the
  tag is re-fetched by name and its IP set asserted; after a delete the tag
  is confirmed gone. Any mismatch fails that row as an `error`.
- **Backup before any write.** The full current state of every tag that will
  be modified or deleted — id, name, parent, rule type, rule text, color,
  criticality, description, created/modified timestamps — is dumped to
  `backup_gaid_tags_<UTC timestamp>.json` and `.xlsx` before the first write
  call.
  - Caveat: for a **static** tag the backup captures the definition but not
    its manual asset assignments, so deleting one is not fully recoverable.
- **All XML values are escaped.** Asset names contain characters such as `&`
  (10 in the reference export) which would otherwise produce malformed XML.
- Failures are loud: unrecoverable errors exit 1.

### Rate limiting and retries

- 5 attempts with exponential backoff capped at 300s.
- Honours `X-RateLimit-ToWait-Sec`, then `Retry-After`, on HTTP 409/429.
- Retries timeouts and connection errors.
- Gives up loudly after the last attempt rather than silently continuing.

---

## Outputs

### Change report

Written every run as `gaid_update_report_<date>.csv` and `.xlsx`. If the
target file is locked (typically open in Excel), the script writes to a
timestamped name beside it rather than losing the run.

- **CSV** — one flat row per GAID.
- **XLSX** — `Summary` (run metadata, rules in force, outcome counts with
  explanations), one sheet per non-empty outcome, and `All Rows`.
- Columns: `tag_id`, `tag_name`, `status`, `file_resource_status`,
  `old_ip_summary`, `new_ip_summary`, `ips_added`, `ips_removed`,
  `ips_excluded`, `error_message`.
- `file_resource_status` shows the `RESOURCE STATUS` breakdown for that GAID
  across its IP-bearing rows (e.g. `In Service: 50, Out of Service: 1`), so
  every outcome explains itself without cross-referencing the source file.
- All sheets have bold frozen headers, autofilter and sized columns.

### Statuses

| Status | Meaning |
| --- | --- |
| `dry-run-update` / `updated` | IP set rewritten in place |
| `dry-run-convert` / `converted` | Static tag converted to dynamic `NETWORK_RANGE` |
| `dry-run-create` / `created` | New tag created |
| `dry-run-delete` / `deleted` | Tag deleted |
| `untouched` | In the file but no IPs to apply, or deletion skipped |
| `skipped-name-contains` | Hostname-pattern tag, never updated by policy |
| `skipped-no-match` | In the file but no valid IPs; no empty tag created |
| `error` | Refused or failed verification |

### Console summary

Totals per outcome, plus whether the run was a dry run or an apply.

---

## Configuration

Constants at the top of `update_gaid_tags.py`:

| Constant | Purpose |
| --- | --- |
| `BASE_URL` | EU2 pod endpoint |
| `GAID_TAG_PREFIX` | `"[VFZ] GAID: "` |
| `GAID_PARENT_TAG_NAME` | Parent tag, resolved to an id at runtime |
| `EXPECTED_RULE_TYPE` | `NETWORK_RANGE` |
| `NEVER_UPDATE_RULE_TYPES` | Rule types never written to (`NAME_CONTAINS`) |
| `STATIC_RULE_TYPES` | What counts as a static tag (`""`, `STATIC`) |
| `INCLUDE_RESOURCE_STATUSES` | Statuses whose rows contribute IPs |
| `DELETE_ON_RESOURCE_STATUSES` | Statuses that mark a GAID decommissioned |
| `NEW_TAG_COLOR` / `NEW_TAG_DESCRIPTION_TEMPLATE` | New tag appearance |
| `*_COLUMN_PATTERNS` | Excel header inference |
| `MAX_RANGE_SIZE` | Guard against absurd IP ranges |

## Files

- `update_gaid_tags.py` — the script.
- `qualys_creds.txt` — credentials (git-ignored).
- `asset-resource-owner.xlsx` — CMDB export (git-ignored; contains internal
  hostnames and IPs).
- `gaid_update_report_*` / `backup_gaid_tags_*` — generated (git-ignored).

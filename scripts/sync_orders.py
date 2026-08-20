#!/usr/bin/env python3
"""
Pulls the 'ORDERS' worksheet out of the SharePoint-hosted Excel workbook
via Microsoft Graph and writes the active (not completed, not cancelled)
rows to data/orders.json.

Run on a schedule (see .github/workflows/sync-orders.yml). Designed to be
safe to run every few minutes: it does a clean read of the whole sheet
each time, so it self-heals from any missed run.

Required environment variables (set as GitHub Actions secrets):
    AZURE_TENANT_ID       - Azure AD tenant ID
    AZURE_CLIENT_ID       - App registration (client) ID
    AZURE_CLIENT_SECRET   - App registration client secret

    SP_SHARE_URL           - the file's SharePoint link, exactly as you'd
                              copy it from "Share" or the browser address
                              bar (e.g. the Doc.aspx?sourcedoc=... link).
                              This is the easiest way to point the script
                              at the file -- no need to know the site path
                              or folder path.

    -- OR, if you'd rather point at a path instead of a link --
    SP_HOSTNAME             - e.g. "salleehorsevans.sharepoint.com"
    SP_SITE_PATH            - e.g. "/sites/Dispatch"
    SP_FILE_PATH            - path to the file from the site's drive root,
                               e.g. "General/ORDERS_LIST_2026.xlsx"

Optional:
    WORKSHEET_NAME          - defaults to "ORDERS " (note the trailing
                               space -- matches the tab name in the
                               original workbook; change if you rename it)
    OUTPUT_PATH              - defaults to "data/orders.json"

MULTI-HORSE ORDERS
-------------------
"HORSE NAME(S)" and " STALL SPACE" both support comma-separated lists for
orders covering more than one horse from the same origin, e.g.:
    HORSE NAME(S):  QUIRKY 25, MISS VEKOMA 25
     STALL SPACE:   1.5, 3
Entries are paired up by position (1st name <-> 1st stall value, etc).
A single horse still works exactly as before -- just one name, one
number, no comma needed. Each order's overall "stallSpace" field is the
sum across all its horses, so existing capacity/board math is unaffected.
If the two lists don't have the same count, that row is logged as a
warning (see WARNING lines in the run's output) rather than silently
guessed at.
"""

import base64
import hashlib
import json
import os
import sys
from datetime import datetime, timezone

import requests

GRAPH = "https://graph.microsoft.com/v1.0"


def env(name, default=None, required=True):
    val = os.environ.get(name, default)
    if required and not val:
        print(f"ERROR: missing required environment variable {name}", file=sys.stderr)
        sys.exit(1)
    return val


def get_token():
    tenant = env("AZURE_TENANT_ID")
    client_id = env("AZURE_CLIENT_ID")
    client_secret = env("AZURE_CLIENT_SECRET")
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    resp = requests.post(
        url,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def encode_share_url(url):
    """Turns any SharePoint/OneDrive sharing or browser URL into the
    'shares' token Graph needs to resolve it directly to a driveItem --
    see https://learn.microsoft.com/graph/api/shares-get"""
    b64 = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    b64 = b64.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b64


def resolve_via_share_url(token, share_url):
    """Preferred path: resolve the file directly from the link you'd
    copy out of SharePoint. Returns (drive_id, item_id)."""
    encoded = encode_share_url(share_url)
    url = f"{GRAPH}/shares/{encoded}/driveItem"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data["parentReference"]["driveId"], data["id"]


def resolve_via_path(token, hostname, site_path, file_path):
    """Fallback path: resolve by hostname + site path + file path."""
    site_url = f"{GRAPH}/sites/{hostname}:{site_path}"
    resp = requests.get(site_url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    site_id = resp.json()["id"]

    item_url = f"{GRAPH}/sites/{site_id}/drive/root:/{file_path}"
    resp = requests.get(item_url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data["parentReference"]["driveId"], data["id"]


def get_used_range(token, drive_id, item_id, worksheet_name):
    # Worksheet names with special characters need URL-safe quoting;
    # requests handles the query string, but the worksheet name sits in
    # the path segment inside parentheses, so we quote it directly.
    from urllib.parse import quote

    encoded_name = quote(f"'{worksheet_name}'")
    url = (
        f"{GRAPH}/drives/{drive_id}/items/{item_id}/workbook/worksheets("
        f"{encoded_name})/usedRange(valuesOnly=true)"
    )
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    return resp.json()["values"]


def excel_serial_to_date(value):
    """Excel sometimes returns dates as ISO strings via Graph, but if a
    raw serial number sneaks through, convert it."""
    if isinstance(value, (int, float)):
        # Excel's epoch, accounting for the 1900 leap-year bug
        from datetime import timedelta

        return (datetime(1899, 12, 30) + timedelta(days=value)).strftime("%Y-%m-%d")
    return value


def split_list(raw):
    """Splits a comma-separated cell value into trimmed, non-empty parts.
    Handles the value coming through as a plain number (single-horse
    orders with no comma) just as well as a comma-separated string."""
    if raw is None or raw == "":
        return []
    return [p.strip() for p in str(raw).split(",") if p.strip() != ""]


def parse_horses(names_raw, stalls_raw):
    """Pairs up HORSE NAME(S) with STALL SPACE by position. Returns
    (horses, mismatch). Handles a stall list with no names yet (creates
    that many unnamed horses) just as well as fully-named orders. A
    mismatch is only flagged when BOTH lists have entries but the counts
    genuinely disagree -- an empty side isn't treated as an error, since
    dispatch may enter stall counts before typing horse names."""
    names = split_list(names_raw)
    stall_strs = split_list(stalls_raw)

    stalls = []
    for s in stall_strs:
        try:
            stalls.append(float(s))
        except ValueError:
            stalls.append(None)

    if not names and not stalls:
        return [], False

    count = max(len(names), len(stalls))
    mismatch = len(names) > 0 and len(stalls) > 0 and len(names) != len(stalls)

    horses = []
    for i in range(count):
        name = names[i] if i < len(names) else None
        stall = stalls[i] if i < len(stalls) else None
        horses.append({"name": name, "stallSpace": stall})

    return horses, mismatch


def stable_id(row_dict):
    """Content-based ID so the same request keeps the same ID even if
    rows get reordered or resorted in the sheet."""
    key = "|".join(
        str(row_dict.get(k, ""))
        for k in [
            "dateTaken",
            "takenBy",
            "requestedBy",
            "originTrainer",
            "origin",
            "destination",
            "track",
            "stallSpace",
        ]
    )
    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:12]


# Column headers as they appear in the ORDERS sheet -> our field names.
# Matched by header text so column reordering in the sheet doesn't break
# this script. If you ever rename a column in Excel, update the key here
# to match -- the sync will print a warning listing any header it
# couldn't match, so a mismatch gets caught immediately next run instead
# of silently dropping data.
HEADER_MAP = {
    "DATE ORDER\nTAKEN": "dateTaken",
    "REQUESTED\n DATES": "requestedDates",
    " TAKEN BY": "takenBy",
    "REQUESTED BY": "requestedBy",
    "ORIGIN FARM": "originFarm",
    "ORIGIN TRAINER": "originTrainer",
    "ORIGIN TRACK": "originTrack",
    "ORIGIN": "origin",
    "DESTINATION": "destination",
    "DESTINATION TRACK": "track",
    "DESTINATION TRAINER": "destinationTrainer",
    "DESTINATION FARM": "destinationFarm",
    "HORSE NAME(S)": "horseNamesRaw",
    " STALL SPACE": "stallSpaceRaw",
    "COMPLETED": "completed",
    "TRIP DATE": "tripDate",
    "DRIVERS": "drivers",
    "CANCELED": "cancel",
}


def parse_rows(values):
    if not values:
        return []
    headers = list(values[0])
    normalized_map = {k.strip(): v for k, v in HEADER_MAP.items()}
    field_names = [
        normalized_map.get(h.strip()) if isinstance(h, str) else None
        for h in headers
    ]

    # Warn loudly (but don't fail the run) about any sheet column that
    # doesn't match anything in HEADER_MAP.
    unmatched = [
        h.strip() if isinstance(h, str) else h
        for h, f in zip(headers, field_names)
        if f is None and h not in (None, "")
    ]
    if unmatched:
        print(
            f"WARNING: {len(unmatched)} column header(s) in the sheet did not "
            f"match HEADER_MAP and will be dropped: {unmatched}",
            file=sys.stderr,
        )

    orders = []
    mismatch_count = 0
    for raw_row in values[1:]:
        row = {}
        for field, val in zip(field_names, raw_row):
            if field is None:
                continue
            row[field] = val

        if not row.get("origin") and not row.get("destination"):
            continue  # blank trailing row

        if row.get("dateTaken"):
            row["dateTaken"] = excel_serial_to_date(row["dateTaken"])
        if row.get("tripDate"):
            row["tripDate"] = excel_serial_to_date(row["tripDate"])
        if isinstance(row.get("requestedDates"), (int, float)):
            row["requestedDates"] = excel_serial_to_date(row["requestedDates"])

        completed = bool(row.get("completed"))
        cancelled = bool(row.get("cancel"))
        if completed or cancelled:
            continue  # only ship active requests to the board

        # ---- multi-horse parsing ----
        horses, mismatch = parse_horses(
            row.pop("horseNamesRaw", None), row.pop("stallSpaceRaw", None)
        )
        if mismatch:
            mismatch_count += 1
            print(
                f"WARNING: HORSE NAME(S) / STALL SPACE count mismatch on row "
                f"(dateTaken={row.get('dateTaken')!r}, origin={row.get('origin')!r}, "
                f"destination={row.get('destination')!r}) -- check that every "
                f"horse listed has a matching stall number, in the same order.",
                file=sys.stderr,
            )

        if not horses:
            # Backward-compatible fallback: nothing in HORSE NAME(S), so
            # treat it as a single unnamed horse using whatever raw
            # stall value (if any) was present, same as the old format.
            horses = [{"name": None, "stallSpace": None}]

        row["horses"] = horses
        row["stallSpace"] = sum(
            h["stallSpace"] for h in horses if h["stallSpace"] is not None
        )

        row["id"] = stable_id(row)
        row.pop("completed", None)
        row.pop("cancel", None)
        orders.append(row)

    if mismatch_count:
        print(
            f"WARNING: {mismatch_count} row(s) had a HORSE NAME(S)/STALL SPACE "
            f"mismatch this run -- see warnings above for which rows.",
            file=sys.stderr,
        )

    return orders


def main():
    worksheet_name = env("WORKSHEET_NAME", default="ORDERS ", required=False)
    output_path = env("OUTPUT_PATH", default="data/orders.json", required=False)

    token = get_token()

    share_url = env("SP_SHARE_URL", required=False)
    if share_url:
        drive_id, item_id = resolve_via_share_url(token, share_url)
    else:
        drive_id, item_id = resolve_via_path(
            token,
            env("SP_HOSTNAME"),
            env("SP_SITE_PATH"),
            env("SP_FILE_PATH"),
        )

    values = get_used_range(token, drive_id, item_id, worksheet_name)
    orders = parse_rows(values)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w") as f:
        json.dump(
            {
                "syncedAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
                "orders": orders,
            },
            f,
            indent=2,
        )

    print(f"Synced {len(orders)} active orders to {output_path}")


if __name__ == "__main__":
    main()

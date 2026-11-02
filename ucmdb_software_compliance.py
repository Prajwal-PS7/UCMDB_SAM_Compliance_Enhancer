#!/usr/bin/env python3
"""
UCMDB 25.4 - Installed Software Authorization & Compliance Report
(READ ONLY - no CI is ever modified)

Output (per computer, per installed software):
    Hostname | IP Address | Serial | Software Name | Version |
    Software Type | Subscription Type | Authorization | Compliance | Reason

Definitions (all VERSION based):
    Authorized    : software is in SOFTWARE_CATALOG AND the installed version
                    is one of its approved_versions.
    Unauthorized  : software not in catalog, or version not approved.
    Compliant     : software is in catalog AND installed version >= min_version.
    Non-Compliant : not in catalog, or installed version < min_version.

TQL REQUIREMENTS (software_version_compliance):
    Node  --(Composition)-->  InstalledSoftware      (required)
    Node  --(Containment)-->  IpAddress              (optional, for IPs)
  Query layout (attributes returned) must include:
    Node              : name, primary_dns_name, serial_number
    InstalledSoftware : name, display_label, version, discovered_vendor (publisher optional)
    IpAddress         : name
  Pagination root stays "InstalledSoftware".

Requires: requests, urllib3   (optional: openpyxl for the .xlsx report)
"""

import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from getpass import getpass
from pathlib import Path
from urllib.parse import quote

import requests
import urllib3

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False


# ============================================================
# CONFIGURATION
# ============================================================

SERVER_URL = "https://localhost.localdomain.com:8443"
API_PATH = "/rest-api"

USERNAME = "admin"
PASSWORD_ENV_VAR = "UCMDB_PASSWORD"   # optional: set to run unattended
CLIENT_CONTEXT = 1

TQL_NAME = "software_version_compliance"
PAGINATION_ROOT = "InstalledSoftware"   # TQL node label, not CI type
CHUNK_SIZE = 100
TIMEOUT = (15, 120)

VERIFY_SSL = False                      # production: True or CA bundle path
HIDE_CERTIFICATE_WARNINGS = True

SOFTWARE_CI_TYPE = "installed_software"
IP_CI_TYPE = "ip_address"

# Attribute names tried (in order) on each CI type.
SOFTWARE_NAME_ATTRS = ("display_label", "name")
SOFTWARE_VERSION_ATTRS = ("version", "software_version")
#SOFTWARE_PUBLISHER_ATTRS = ("publisher", "vendor")
SOFTWARE_PUBLISHER_ATTRS = ("discovered_vendor", "vendor")
HOST_NAME_ATTRS = ("name", "primary_dns_name", "display_label")
SERIAL_ATTRS = ("serial_number", "serial")
IP_ATTRS_ON_IP_CI = ("name", "ip_address", "display_label")
IP_ATTRS_ON_NODE = ("primary_ip_address", "ip_address")

EXCLUDE_IPV6 = True
EXCLUDE_IP_PREFIXES = ("127.", "169.254.")

# Software that is not in the catalog:
#   True  -> listed as Unauthorized / Non-Compliant
#   False -> skipped entirely (only catalog software is reported)
REPORT_UNLISTED_SOFTWARE = True

# Optional noise filter: case-insensitive substrings; matching software
# is skipped (e.g. "Update for", "Security Update", "Redistributable").
IGNORE_NAME_PATTERNS = []

UNCLASSIFIED = "Unclassified"

# ------------------------------------------------------------
# SOFTWARE CATALOG  (edit to match the customer's contracts)
#   name              : official name (also matched case-insensitively)
#   aliases           : other display names / prefixes seen in UCMDB
#   approved_versions : exact version strings => "Authorized"
#   min_version       : installed >= this => "Compliant"
#   software_type     : Freeware / Commercial / Licensed
#   subscription_type : Freeware / Perpetual / Annual Subscription / ...
#
# Name matching: exact (case-insensitive) OR installed name starts
# with the catalog name/alias (so "7-Zip 26.01 (x64)" matches "7-Zip").
# ------------------------------------------------------------
SOFTWARE_CATALOG = [
    {
        "name": "WPS Office",
        "aliases": [],
        "approved_versions": ["10,2,0,7549"],
        "min_version": "10,2,0,7549",
        "software_type": "Commercial",
        "subscription_type": "Annual Subscription",
    },
    {
        "name": "Microsoft Edge",
        "aliases": [],
        "approved_versions": ["153.0.4234.48"],
        "min_version": "153.0.4234.48",
        "software_type": "Freeware",
        "subscription_type": "Freeware",
    },
    {
        "name": "7-Zip",
        "aliases": [],
        "approved_versions": ["26.01"],
        "min_version": "26.01",
        "software_type": "Freeware",
        "subscription_type": "Freeware",
    },
    {
        "name": "Universal Discovery Agent",
        "aliases": [],
        "approved_versions": ["25.4.000.98"],
        "min_version": "25.4.000.98",
        "software_type": "Licensed",
        "subscription_type": "Perpetual",
    },
]

API_BASE = SERVER_URL.rstrip("/") + "/" + API_PATH.strip("/")
OUTPUT_DIR = Path(__file__).resolve().parent
RAW_DIR = OUTPUT_DIR / "raw_chunks"

if VERIFY_SSL is False and HIDE_CERTIFICATE_WARNINGS:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================
# FILE HELPERS
# ============================================================

def save_json(filename, data, folder=OUTPUT_DIR):
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / filename
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return path


# ============================================================
# HTTP HELPERS
# ============================================================

def request_json(session, method, path, **kwargs):
    endpoint = API_BASE + "/" + path.lstrip("/")
    label = (
        "topology/result/[result-id]/[chunk]"
        if path.startswith("topology/result/")
        else path
    )

    try:
        response = session.request(
            method=method, url=endpoint, timeout=TIMEOUT,
            verify=VERIFY_SSL, allow_redirects=False, **kwargs
        )
    except requests.exceptions.SSLError as exc:
        raise RuntimeError(
            "TLS certificate verification failed. Configure a trusted CA bundle."
        ) from exc
    except requests.exceptions.RequestException as exc:
        raise RuntimeError(f"{method} {label}: {type(exc).__name__}") from exc

    print(f"{method} {label} -> HTTP {response.status_code}")

    if not 200 <= response.status_code < 300:
        raise RuntimeError(
            f"{method} {label} failed with HTTP {response.status_code}. "
            "Check server logs, API permissions, and the pagination node label."
        )

    try:
        return json.loads(response.content.decode("utf-8-sig"))
    except (ValueError, UnicodeError) as exc:
        raise RuntimeError(f"{label}: response is not valid JSON.") from exc


def authenticate(session):
    password = os.environ.get(PASSWORD_ENV_VAR) or getpass("Enter UCMDB password: ")
    if not password:
        raise RuntimeError("Password cannot be empty.")

    result = request_json(
        session, "POST", "authenticate",
        json={
            "username": USERNAME,
            "password": password,
            "clientContext": CLIENT_CONTEXT,
        },
    )

    token = result.get("token") if isinstance(result, dict) else None
    if not isinstance(token, str) or not token.strip():
        raise RuntimeError("Authentication returned no valid token.")

    session.headers["Authorization"] = f"Bearer {token.strip()}"
    print("Authentication successful. Token is not displayed.")


# ============================================================
# ATTRIBUTE READING
# ============================================================

def extract_attributes(ci):
    """Read explicitly named attributes (properties / attributeDefinitions)."""
    attributes = {}

    properties = ci.get("properties")
    if isinstance(properties, dict):
        attributes.update(properties)

    definitions = ci.get("attributeDefinitions")

    def read_value(entry):
        if isinstance(entry, dict):
            for key in ("rawValue", "value"):
                if key in entry and entry[key] is not None:
                    return True, entry[key]
            return False, None
        if entry is None or isinstance(entry, (str, int, float, bool)):
            return True, entry
        return False, None

    def add_attribute(name, entry):
        if not isinstance(name, str) or not name:
            return
        present, value = read_value(entry)
        if not present:
            return
        if name not in attributes or attributes[name] in (None, ""):
            attributes[name] = value

    if isinstance(definitions, list):
        for entry in definitions:
            if isinstance(entry, dict):
                add_attribute(entry.get("name") or entry.get("attributeName"), entry)
    elif isinstance(definitions, dict):
        for name, entry in definitions.items():
            add_attribute(name, entry)

    return attributes


def first_text(attrs, names):
    """First non-empty string value among the given attribute names."""
    for name in names:
        value = attrs.get(name)
        if value is None:
            continue
        value = str(value).strip()
        if value:
            return value
    return ""


# ============================================================
# TOPOLOGY: SERVER CHUNKS (CIs + RELATIONS)
# ============================================================

def fetch_topology(session):
    """Returns (cis_by_id, relations_list). Relations are collected from
    ALL chunks, then linked afterwards (a relation may reference a CI that
    arrived in a different chunk)."""
    result = request_json(
        session, "GET", "topology",
        params={
            "queryName": TQL_NAME,
            "chunkSize": CHUNK_SIZE,
            "rootNodeForPagination": PAGINATION_ROOT,
            "includeEmptyLayoutProperties": "true",
        },
    )

    if not isinstance(result, dict):
        raise RuntimeError("Topology response must be a JSON object.")

    cis = {}
    relations = {}

    def collect(data, filename):
        if not isinstance(data, dict):
            raise RuntimeError("Topology result must be a JSON object.")

        records = data.get("cis")
        if not isinstance(records, list):
            raise RuntimeError(
                "Topology result has no valid cis list. "
                f"Returned keys: {list(data.keys())}"
            )

        save_json(filename, data, RAW_DIR)

        for ci in records:
            if not isinstance(ci, dict) or not ci.get("ucmdbId"):
                raise RuntimeError("A topology record is missing ucmdbId.")

            ci_id = str(ci["ucmdbId"])
            parsed = extract_attributes(ci)

            if ci_id not in cis:
                cis[ci_id] = {
                    "ucmdbId": ci_id,
                    "type": ci.get("type"),
                    "attributes": parsed,
                }
            else:
                existing = cis[ci_id]["attributes"]
                for key, value in parsed.items():
                    if key not in existing or value not in (None, ""):
                        existing[key] = value

        rel_records = data.get("relations") or []
        for rel in rel_records:
            if not isinstance(rel, dict):
                continue
            end1, end2 = rel.get("end1Id"), rel.get("end2Id")
            if not end1 or not end2:
                continue
            relations[(str(end1), str(end2), rel.get("type"))] = {
                "type": rel.get("type"),
                "end1": str(end1),
                "end2": str(end2),
            }

        print(
            f"Response CIs: {len(records)} | Relations: {len(rel_records)} | "
            f"Unique CIs: {len(cis)} | Unique relations: {len(relations)}"
        )

    # Small result returned directly
    if "numberOfChunks" not in result:
        print(f"\nServer returned a direct result (chunk size {CHUNK_SIZE}).")
        collect(result, "chunk_001.json")
        return cis, list(relations.values())

    try:
        chunk_count = int(result["numberOfChunks"])
    except (TypeError, ValueError) as exc:
        raise RuntimeError("Invalid numberOfChunks.") from exc

    if chunk_count < 0:
        raise RuntimeError("numberOfChunks cannot be negative.")

    if chunk_count == 0:
        if isinstance(result.get("cis"), list):
            collect(result, "chunk_001.json")
        return cis, list(relations.values())

    result_id = result.get("queryResultId")
    if not result_id:
        raise RuntimeError("Chunked response has no queryResultId.")

    encoded_id = quote(str(result_id), safe="")
    print(f"\nServer chunks: {chunk_count} | Requested size: {CHUNK_SIZE}")

    for number in range(1, chunk_count + 1):
        print(f"\nFetching server chunk {number}/{chunk_count}")
        data = request_json(
            session, "GET", f"topology/result/{encoded_id}/{number}",
            params={"includeEmptyLayoutProperties": "true"},
        )
        collect(data, f"chunk_{number:03d}.json")

    return cis, list(relations.values())


# ============================================================
# VERSION + CATALOG LOGIC
# ============================================================

def parse_version(text):
    """'10,2,0,7549' / '25.4.000.98' -> (10, 2, 0, 7549). None if no digits.
    Used ONLY for ordering; authorization uses the exact string."""
    if not isinstance(text, str):
        return None
    parts = re.findall(r"\d+", text)
    return tuple(int(p) for p in parts) if parts else None


def version_gte(installed, minimum):
    a, b = parse_version(installed), parse_version(minimum)
    if a is None or b is None:
        return None  # cannot decide
    size = max(len(a), len(b))
    a += (0,) * (size - len(a))
    b += (0,) * (size - len(b))
    return a >= b


def normalise_version(text):
    return re.sub(r"\s+", "", text or "").lower()


def build_catalog_index():
    index = []
    for entry in SOFTWARE_CATALOG:
        keys = [entry["name"]] + list(entry.get("aliases", []))
        index.append((sorted((k.lower() for k in keys), key=len, reverse=True), entry))
    return index


def find_catalog_entry(name, index):
    lowered = name.lower().strip()
    # exact match first
    for keys, entry in index:
        if lowered in keys:
            return entry
    # then prefix match (longest key wins)
    best, best_len = None, 0
    for keys, entry in index:
        for key in keys:
            if lowered.startswith(key) and len(key) > best_len:
                best, best_len = entry, len(key)
    return best


def evaluate(name, version, entry):
    """Returns authorization, compliance, type, subscription, reason."""
    if entry is None:
        return {
            "authorization": "Unauthorized",
            "compliance": "Non-Compliant",
            "software_type": UNCLASSIFIED,
            "subscription_type": UNCLASSIFIED,
            "reason": "Software not in approved catalog",
        }

    base = {
        "software_type": entry.get("software_type", UNCLASSIFIED),
        "subscription_type": entry.get("subscription_type", UNCLASSIFIED),
    }

    if not version:
        return {**base,
                "authorization": "Unauthorized",
                "compliance": "Non-Compliant",
                "reason": "Version missing - cannot verify"}

    approved = {normalise_version(v) for v in entry.get("approved_versions", [])}
    authorized = normalise_version(version) in approved

    min_version = entry.get("min_version")
    if min_version:
        gte = version_gte(version, min_version)
        compliant = bool(gte)
        comp_note = (
            "version cannot be compared" if gte is None
            else f"version {'>=' if gte else '<'} minimum {min_version}"
        )
    else:
        compliant = authorized
        comp_note = "no min_version defined; follows authorization"

    auth_note = (
        "approved version" if authorized
        else f"version not in approved list ({', '.join(entry.get('approved_versions', []))})"
    )

    return {
        **base,
        "authorization": "Authorized" if authorized else "Unauthorized",
        "compliance": "Compliant" if compliant else "Non-Compliant",
        "reason": f"{auth_note}; {comp_note}",
    }


# ============================================================
# BUILD COMPUTER -> SOFTWARE INVENTORY
# ============================================================

def clean_ip(value):
    if not value:
        return None
    value = value.strip()
    if EXCLUDE_IPV6 and ":" in value:
        return None
    if value.startswith(EXCLUDE_IP_PREFIXES):
        return None
    return value


def build_inventory(cis, relations):
    """Link software CIs to their computers using relations."""
    index = build_catalog_index()

    software_ids = {i for i, c in cis.items() if c.get("type") == SOFTWARE_CI_TYPE}
    ip_ids = {i for i, c in cis.items() if c.get("type") == IP_CI_TYPE}

    sw_to_host = {}
    host_to_ips = defaultdict(set)

    for rel in relations:
        e1, e2 = rel["end1"], rel["end2"]
        for a, b in ((e1, e2), (e2, e1)):
            # a = candidate computer, b = software / ip
            if a in cis and a not in software_ids and a not in ip_ids:
                if b in software_ids:
                    sw_to_host.setdefault(b, a)
                elif b in ip_ids:
                    host_to_ips[a].add(b)

    computers = {}
    unresolved = []
    counts = Counter()

    def computer_record(host_id):
        if host_id in computers:
            return computers[host_id]
        attrs = cis[host_id]["attributes"]

        ips = set()
        for ip_id in host_to_ips.get(host_id, ()):
            ip = clean_ip(first_text(cis[ip_id]["attributes"], IP_ATTRS_ON_IP_CI))
            if ip:
                ips.add(ip)
        if not ips:
            ip = clean_ip(first_text(attrs, IP_ATTRS_ON_NODE))
            if ip:
                ips.add(ip)

        computers[host_id] = {
            "ucmdbId": host_id,
            "hostname": first_text(attrs, HOST_NAME_ATTRS) or f"<unnamed {host_id[:8]}>",
            "ip_address": ", ".join(sorted(ips)),
            "serial": first_text(attrs, SERIAL_ATTRS),
            "software": [],
        }
        return computers[host_id]

    for sw_id in sorted(software_ids):
        attrs = cis[sw_id]["attributes"]
        name = first_text(attrs, SOFTWARE_NAME_ATTRS)
        version = first_text(attrs, SOFTWARE_VERSION_ATTRS)
        discovered_vendor = first_text(attrs, SOFTWARE_PUBLISHER_ATTRS)

        if not name:
            counts["missing_name"] += 1
            unresolved.append({
                "ucmdbId": sw_id,
                "reason": "software name missing",
                "available_attributes": attrs,
            })
            continue

        if any(p.lower() in name.lower() for p in IGNORE_NAME_PATTERNS):
            counts["ignored"] += 1
            continue

        host_id = sw_to_host.get(sw_id)
        if not host_id:
            counts["no_computer"] += 1
            unresolved.append({
                "ucmdbId": sw_id,
                "name": name,
                "version": version,
                "reason": "no computer relation found in query result",
            })
            continue

        entry = find_catalog_entry(name, index)
        if entry is None and not REPORT_UNLISTED_SOFTWARE:
            counts["unlisted_skipped"] += 1
            continue

        result = evaluate(name, version, entry)
        computer_record(host_id)["software"].append({
            "ucmdbId": sw_id,
            "name": name,
            "version": version,
            "discovered_vendor": discovered_vendor,
            "catalog_name": entry["name"] if entry else "",
            **result,
        })
        counts[result["authorization"].lower()] += 1
        counts[result["compliance"].lower().replace("-", "_")] += 1

    for comp in computers.values():
        comp["software"].sort(key=lambda s: (s["name"].lower(), s["version"]))

    return sorted(computers.values(), key=lambda c: c["hostname"].lower()), unresolved, counts


# ============================================================
# REPORT WRITERS
# ============================================================

DETAIL_HEADERS = [
    "Hostname", "IP Address", "Serial", "Software Name", "Version", "Software Type", "Subscription Type", "Authorization", "Compliance", "Reason",
]
SUMMARY_HEADERS = [
    "Hostname", "IP Address", "Serial", "Total Software", "Authorized", "Unauthorized", "Compliant", "Non-Compliant", "Computer Status",
]


def detail_rows(computers):
    rows = []
    for c in computers:
        for s in c["software"]:
            rows.append([
                c["hostname"], c["ip_address"], c["serial"],
                s["name"], s["version"], s["software_type"], s["subscription_type"],
                s["authorization"], s["compliance"], s["reason"],
            ])
    return rows


def summary_rows(computers):
    rows = []
    for c in computers:
        sw = c["software"]
        auth = sum(1 for s in sw if s["authorization"] == "Authorized")
        comp = sum(1 for s in sw if s["compliance"] == "Compliant")
        unauth, noncomp = len(sw) - auth, len(sw) - comp
        status = "Compliant" if sw and unauth == 0 and noncomp == 0 else (
            "No software" if not sw else "Non-Compliant")
        rows.append([
            c["hostname"], c["ip_address"], c["serial"], len(sw),
            auth, unauth, comp, noncomp, status,
        ])
    return rows


def write_csv(path, headers, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(headers)
        writer.writerows(rows)


def write_excel(path, computers, unresolved, stamp):
    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="1F3864")
    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    good = {"Authorized", "Compliant"}
    bad = {"Unauthorized", "Non-Compliant"}

    def sheet(ws, headers, rows, color_cols=()):
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center")
        for row in rows:
            ws.append(row)
        for col_idx in color_cols:
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                if cell.value in good:
                    cell.fill = green
                elif cell.value in bad:
                    cell.fill = red
        ws.freeze_panes = "A2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
        for idx, header in enumerate(headers, 1):
            longest = max(
                [len(str(header))] +
                [len(str(ws.cell(row=r, column=idx).value or ""))
                 for r in range(2, min(ws.max_row, 500) + 1)]
            )
            ws.column_dimensions[get_column_letter(idx)].width = min(longest + 3, 60)

    ws = wb.active
    ws.title = "Software by Computer"
    sheet(ws, DETAIL_HEADERS, detail_rows(computers), color_cols=(9, 10))

    sheet(wb.create_sheet("Computer Summary"), SUMMARY_HEADERS,
          summary_rows(computers), color_cols=(9,))

    flat = [r for r in detail_rows(computers)]
    sheet(wb.create_sheet("Unauthorized"), DETAIL_HEADERS,
          [r for r in flat if r[8] == "Unauthorized"], color_cols=(9, 10))
    sheet(wb.create_sheet("Non-Compliant"), DETAIL_HEADERS,
          [r for r in flat if r[9] == "Non-Compliant"], color_cols=(9, 10))

    sheet(wb.create_sheet("Catalog"),
          ["Name", "Aliases", "Approved Versions", "Min Version",
           "Software Type", "Subscription Type"],
          [[e["name"], ", ".join(e.get("aliases", [])),
            ", ".join(e.get("approved_versions", [])), e.get("min_version", ""),
            e.get("software_type", ""), e.get("subscription_type", "")]
           for e in SOFTWARE_CATALOG])

    sheet(wb.create_sheet("Unresolved"),
          ["ucmdbId", "Name", "Version", "Reason"],
          [[u.get("ucmdbId"), u.get("name", ""), u.get("version", ""),
            u.get("reason", "")] for u in unresolved])

    wb.save(path)


def write_reports(computers, unresolved, stamp):
    outputs = []

    detail = detail_rows(computers)
    summary = summary_rows(computers)

    p = OUTPUT_DIR / f"software_compliance_detail_{stamp}.csv"
    write_csv(p, DETAIL_HEADERS, detail)
    outputs.append(p)

    p = OUTPUT_DIR / f"computer_summary_{stamp}.csv"
    write_csv(p, SUMMARY_HEADERS, summary)
    outputs.append(p)

    outputs.append(save_json(f"software_compliance_report_{stamp}.json", computers))
    outputs.append(save_json(f"software_unresolved_{stamp}.json", unresolved))

    if HAVE_OPENPYXL:
        p = OUTPUT_DIR / f"software_compliance_report_{stamp}.xlsx"
        write_excel(p, computers, unresolved, stamp)
        outputs.append(p)
    else:
        print("\nNOTE: openpyxl not installed - .xlsx skipped. "
              "Install with: pip install openpyxl")

    return outputs


# ============================================================
# MAIN: READ ONLY
# ============================================================

def main():
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    print("=" * 70)
    print("UCMDB SOFTWARE AUTHORIZATION & COMPLIANCE REPORT")
    print("=" * 70)
    print(f"TQL: {TQL_NAME} | Pagination root: {PAGINATION_ROOT} | "
          f"Chunk: {CHUNK_SIZE}")
    print("Mode: READ ONLY - NO CI UPDATES")

    print("\nApproved catalog:")
    for e in SOFTWARE_CATALOG:
        print(f"  {e['name']} | approved: {', '.join(e['approved_versions'])} "
              f"| min: {e.get('min_version', '-')} | "
              f"{e.get('software_type')} / {e.get('subscription_type')}")

    save_json("software_run_status.json", {"status": "running", "mode": "read_only"})

    with requests.Session() as session:
        session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
        })

        authenticate(session)
        cis, relations = fetch_topology(session)

    types = Counter(str(c.get("type")) for c in cis.values())
    computers, unresolved, counts = build_inventory(cis, relations)
    outputs = write_reports(computers, unresolved, stamp)

    summary = {
        "mode": "read_only",
        "generated": stamp,
        "total_unique_cis": len(cis),
        "total_relations": len(relations),
        "ci_types": dict(types),
        "computers_with_software": len(computers),
        "software_records_reported": sum(len(c["software"]) for c in computers),
        "unresolved": len(unresolved),
        "counts": dict(counts),
    }
    save_json("software_summary.json", summary)

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(json.dumps(summary, indent=2))

    print("\nReports written:")
    for p in outputs:
        print(f"  {p}")

    if not relations:
        print(
            "\nWARNING: the query returned NO relations. Computer/IP/serial "
            "cannot be linked. Make sure the TQL has Node -> InstalledSoftware "
            "(and Node -> IpAddress) with the relationships included in the layout."
        )
    if unresolved:
        print(
            f"\nATTENTION: {len(unresolved)} software record(s) could not be "
            "attributed (see software_unresolved_*.json). They were NOT marked "
            "Unauthorized."
        )

    save_json("software_run_status.json", {
        "status": "completed" if relations and not unresolved else "completed_with_warnings",
        "mode": "read_only",
    })
    print("\nNo CIs were modified.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        save_json("software_run_status.json",
                  {"status": "interrupted", "mode": "read_only"})
        print("\nInterrupted. No CI updates were submitted.")
        sys.exit(130)
    except Exception as exc:
        save_json("software_run_status.json",
                  {"status": "failed", "mode": "read_only", "error": str(exc)})
        print(f"\nERROR: {exc}")
        sys.exit(1)
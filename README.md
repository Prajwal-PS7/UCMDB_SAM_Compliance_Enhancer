# UCMDB SAM Compliance Enhancer (REST API)
<img width="2752" height="1536" alt="Gemini_Generated_Image_i5ey1ei5ey1ei5ey" src="https://github.com/user-attachments/assets/543bef4b-0e7c-4f11-bcd5-75217c8d0a4e" />




## UCMDB Software Authorization & Compliance Report

🔹 Script : ucmdb_software_compliance.py
🔹 Target : Micro Focus / OpenText UCMDB v25.4 (REST API)
🔹 Mode   : READ ONLY - the script never creates, updates or deletes any CI

--------------------------------------------------------------------------------
1. PURPOSE
--------------------------------------------------------------------------------
UCMDB / SAM cannot natively classify software by version-based authorization and compliance, or hold the extra classifications the customer needs. This script fills that gap. It pulls computers and their installed software from UCMDB through a TQL, compares each software item against a maintained catalog, and produces reports with these columns:

  Computer  : Hostname, IP Address, Serial
  Software  : Name, Version, Publisher, Software Type, Subscription Type, Authorization, Compliance, Reason

--------------------------------------------------------------------------------
2. HOW THE DECISIONS ARE MADE (VERSION BASED)
--------------------------------------------------------------------------------
  AUTHORIZED     Software is in the catalog AND the installed version exactly matches one of its "approved_versions".

  UNAUTHORIZED   Software is not in the catalog, OR its version is not in the approved list, OR the version is missing.

  COMPLIANT      Software is in the catalog AND installed version is greater than or equal to "min_version" (updated version).

  NON-COMPLIANT  Installed version is lower than "min_version" (outdated), OR software is not in the catalog, OR version is missing.

Notes
  - Authorization uses the EXACT version text (26.01 stays 26.01).
  - Compliance compares versions number by number, e.g.
        10,2,0,7549  ->  10.2.0.7549
        25.4.000.98  ->  25.4.0.98
    so 25.10 is correctly newer than 25.4.
  - Name matching is case-insensitive. An exact name matches, and so does a prefix: "7-Zip 26.01 (x64)" matches the catalog entry "7-Zip".
  - Software that cannot be linked to a computer is NOT marked Unauthorized. It is listed separately as "unresolved".

--------------------------------------------------------------------------------
3. REQUIREMENTS
--------------------------------------------------------------------------------
  - Python 3.8 or newer
  - Python packages:      pip install requests urllib3
  - Optional (Excel):     pip install openpyxl (without openpyxl the script still writes CSV and JSON, and prints a note)
  - Network access from the machine running the script to the UCMDB server (default https://localhost.localdomain.com:8443)
  - A UCMDB user with permission to call the REST API and run the TQL

--------------------------------------------------------------------------------
4. UCMDB-SIDE SETUP (ONE TIME)
--------------------------------------------------------------------------------
Create (or verify) a TQL named:   software_version_compliance

Required structure
    Node  --(Composition)-->  InstalledSoftware         (required)
    Node  --(Containment)-->  IpAddress                 (optional, for IPs)

Query layout - attributes that must be returned
    Node               name, primary_dns_name, serial_number
    InstalledSoftware  name, display_label, version, publisher (publisher is optional)
    IpAddress          name

Other settings
  - Query node label of the pagination root must be exactly "InstalledSoftware" (this is the TQL node label, not the CI type name).
  - Make sure relationships are included in the query result. Without relations the script cannot link software to computers.

--------------------------------------------------------------------------------
5. CONFIGURATION (TOP OF THE SCRIPT)
--------------------------------------------------------------------------------
  SERVER_URL                 UCMDB server URL and port
  API_PATH                   REST path, normally /rest-api
  USERNAME                   UCMDB user (default: admin)
  PASSWORD_ENV_VAR           Name of environment variable holding the password (default: UCMDB_PASSWORD)
  TQL_NAME                   Name of the TQL
  PAGINATION_ROOT            TQL node label used for chunking
  CHUNK_SIZE                 Records per server chunk (default 100)
  VERIFY_SSL                 False = testing only. Production: True, or the path to a trusted CA bundle file
  SOFTWARE_CI_TYPE           CI type of installed software (installed_software)
  IP_CI_TYPE                 CI type of IP addresses (ip_address)
  *_ATTRS tuples             Attribute names tried (in order) for software name, version, publisher, hostname, serial and IP. Edit if your model uses different names.
  EXCLUDE_IPV6               True = ignore IPv6 addresses
  EXCLUDE_IP_PREFIXES        IPs starting with these are ignored (127., 169.254.)
  REPORT_UNLISTED_SOFTWARE   True  = software not in catalog is reported as Unauthorized / Non-Compliant
                             False = only catalog software is reported
  IGNORE_NAME_PATTERNS       List of text fragments; software whose name contains any of them is skipped, e.g.
                             ["Security Update", "Redistributable"]
  SOFTWARE_CATALOG           The approved software list (see section 6)

--------------------------------------------------------------------------------
6. MAINTAINING THE SOFTWARE CATALOG
--------------------------------------------------------------------------------
Each catalog entry looks like this:

    {
        "name": "7-Zip",
        "aliases": [],
        "approved_versions": ["26.01"],
        "min_version": "26.01",
        "software_type": "Freeware",
        "subscription_type": "Freeware",
    },

  name               Official product name (also used for matching)
  aliases            Other display names / prefixes seen in UCMDB
  approved_versions  Exact version strings that count as AUTHORIZED
  min_version        Lowest version that counts as COMPLIANT
  software_type      Freeware / Commercial / Licensed
  subscription_type  Freeware / Perpetual / Annual Subscription / etc.

To approve a new version      add it to approved_versions
To make an old version stale  raise min_version
To add new software           copy an entry and change the values

The catalog is the single source of truth. Review it whenever a new version is approved or a contract changes.

--------------------------------------------------------------------------------
7. HOW TO RUN
--------------------------------------------------------------------------------
Interactive (asks for the password; nothing is echoed):

    python ucmdb_software_compliance.py

Unattended / scheduled run - set the password in the environment first.

  Windows (Command Prompt):
    set UCMDB_PASSWORD=your_password
    python ucmdb_software_compliance.py

  Windows (PowerShell):
    $env:UCMDB_PASSWORD = "your_password"
    python ucmdb_software_compliance.py

  Linux:
    export UCMDB_PASSWORD='your_password'
    python3 ucmdb_software_compliance.py

For scheduling use Windows Task Scheduler or cron. Prefer a secured secret
store over a plain-text password where possible.

Exit codes
    0    completed
    1    failed (see the ERROR line and software_run_status.json)
    130  interrupted by the user (Ctrl+C)


--------------------------------------------------------------------------------
8. OUTPUT FILES
--------------------------------------------------------------------------------
All files are written next to the script. <stamp> = YYYYMMDD_HHMMSS.

  software_compliance_report_<stamp>.xlsx   Main report (needs openpyxl)
      Sheets:
        Software by Computer   one row per computer + software
        Computer Summary       one row per computer with counts and status
        Unauthorized           only Unauthorized rows
        Non-Compliant          only Non-Compliant rows
        Catalog                catalog used for this run
        Unresolved             records that could not be attributed
      Authorized / Compliant cells are green; Unauthorized / Non-Compliant cells are red. Header row is frozen and filterable.

  software_compliance_detail_<stamp>.csv    Same as "Software by Computer"
  computer_summary_<stamp>.csv              Same as "Computer Summary"
  software_compliance_report_<stamp>.json   Full structured data per computer
  software_unresolved_<stamp>.json          Unattributed records and reasons
  software_summary.json                     Run totals and counts
  software_run_status.json                  running / completed / completed_with_warnings / failed / interrupted
  raw_chunks\chunk_NNN.json                 Raw server responses, kept for checking the TQL layout

"Computer Status" in the summary is Compliant only when every software item on that computer is both Authorized and Compliant.

--------------------------------------------------------------------------------
9. TROUBLESHOOTING / HELP
--------------------------------------------------------------------------------
Problem: "TLS certificate verification failed"
  Fix:   Set VERIFY_SSL to True and install the server certificate chain, or set VERIFY_SSL to the path of a CA bundle file.

Problem: HTTP 401 on authenticate
  Fix:   Wrong username/password, or the account is locked. Check USERNAME and the password / UCMDB_PASSWORD value.

Problem: HTTP 403 on topology
  Fix:   The user lacks permission to run the TQL or use the REST API.

Problem: HTTP 400 / 404 on topology
  Fix:   Check TQL_NAME and PAGINATION_ROOT. PAGINATION_ROOT must be the TQL node label ("InstalledSoftware"), not the CI type.

Problem: "WARNING: the query returned NO relations"
  Fix:   The TQL has no Node -> InstalledSoftware link, or relations are not included in the layout. Fix the TQL; hostnames/IP/serial cannot be filled until relations are returned.

Problem: Hostname, IP or serial is blank
  Fix:   The attribute is not in the query layout, or has a different name. Open raw_chunks\chunk_001.json, find the real attribute name and add it to HOST_NAME_ATTRS / SERIAL_ATTRS / IP_ATTRS_* in the script.

Problem: Software shows as Unauthorized but should be approved
  Fix:   Compare the name and version in the report with the catalog.
         - Add an alias if UCMDB shows a different name.
         - The version must match the approved_versions text exactly.

Problem: Many records in the Unresolved sheet
  Fix:   "software name missing" -> software name attribute is not in the          layout. "no computer relation found" -> the TQL does not return the Node for those records.

Problem: No .xlsx file produced
  Fix:   Run  pip install openpyxl  and run the script again. CSV and JSON files are still created without it.

Problem: Report shows noise (updates, runtimes)
  Fix:   Add text fragments to IGNORE_NAME_PATTERNS.

--------------------------------------------------------------------------------
10. SECURITY NOTES
--------------------------------------------------------------------------------
  - The script is read only; it never writes to UCMDB.
  - The password is never stored or printed. The session token and query result IDs are not displayed.
  - Do not leave VERIFY_SSL = False in production.
  - Do not hard-code the password in the script. Use the prompt or the environment variable.
  - Reports contain hostnames, IPs and serial numbers. Store and share them according to your data-handling rules.

--------------------------------------------------------------------------------
11. KNOWN LIMITATIONS
--------------------------------------------------------------------------------
  - This is a custom script and is not part of the standard UCMDB / SAM product.
  - Software type and subscription type come from the catalog, not from UCMDB.
  - Compliance here means version compliance (installed version vs minimum version). It is NOT SAM license-entitlement compliance.
  - Results are produced as reports only; nothing is written back to UCMDB.
  - Only software that the TQL returns is evaluated.

--------------------------------------------------------------------------------
12. QUICK START CHECKLIST
--------------------------------------------------------------------------------
  [ ] Python 3.8+ installed; pip install requests urllib3 openpyxl
  [ ] TQL software_version_compliance created with relations and attributes
  [ ] SERVER_URL, USERNAME, TQL_NAME set in the script
  [ ] SOFTWARE_CATALOG reviewed and approved by the customer
  [ ] Test run; check Unresolved sheet and raw_chunks if anything is blank
  [ ] VERIFY_SSL set properly for production
  [ ] Schedule the run if needed


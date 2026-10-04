#!/usr/bin/env python3
"""
Standalone verification for the generated Fly91 Weekly Status Report (.docx).

Loads a generated .docx with python-docx and asserts that the key structural
elements required by the REAL Fly91 reference layout (see
REFERENCE_FLY91_GOOGLEDOC.txt) are present, in particular:

  * the "Cost Summary Differences" per-service table (with Inspector,
    GuardDuty and a Total Cost row),
  * the Disclaimer and Contents sections,
  * the "BAU Matrix Overview of the resources." section with the EC2 / RDS /
    ELB / WAF utilisation tables (populated with the reference resources),
  * the Amazon Inspector and Guard Duty findings summaries, and
  * the "-- End of Document --" trailer.

Usage:
    python3 scripts/verify_word_report.py <path-to.docx> [--min-images N]
        [--contains "substr" ...]
"""

import argparse
import sys

from docx import Document


def collect_text(doc):
    paras = [p.text for p in doc.paragraphs]
    all_para = "\n".join(paras)
    cells = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cells.append(cell.text)
    all_cell = "\n".join(cells)
    return paras, all_para, all_cell


# Reference EC2 instance IDs that must appear in the EC2 utilisation table
# (8 running + 1 stopped = 9 rows).
REQUIRED_EC2_IDS = [
    "i-06a53ce120fca7a28",   # Fortinet-FW
    "i-041cbf2e1ea9a05ee",   # FLY91- Staging Website
    "i-056b66d8b569a4a26",   # FLY91- Production Website
    "i-0b108151d07fe2bf8",   # Fly91-SFTP
    "i-0e24da06be3f86a40",   # Fly91-Desk
    "i-0d0653d04b9306f42",   # FLY91-Website-testing (stopped)
    "i-00a0097da058e39a1",   # Fly91-Development-Server-Windows
    "i-0233bf5f2a51870df",   # Fly91-Development-Server-Ubuntu
    "i-08b66727cdf8d5896",   # FLY91 - Data Science Production
]

# Reference RDS DB names (6).
REQUIRED_RDS = [
    "fly91-data-science-production-db",
    "fly91-db",
    "fly91-dev-db",
    "fly91-development-psql",
    "fly91-prod-psql-db",
    "fly91-uat-psql-db",
]

# Reference ELB ALB names (3).
REQUIRED_ELB = ["Stage-Fly91-ALB", "Prod-Fly91-ALB", "Fly91-Development-ALB"]

# Reference WAF web ACL names (2).
REQUIRED_WAF = ["Prod-WAF-Fly-91", "Stag-WAF-Fly-91"]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", help="Path to the generated .docx")
    parser.add_argument("--contains", nargs="*", default=[],
                        help="Extra substrings that must appear anywhere in the document")
    parser.add_argument("--min-images", type=int, default=0,
                        help="Minimum number of embedded images (inline shapes) that must "
                             "be present. Defaults to 0 so image-less reports still pass.")
    args = parser.parse_args()

    doc = Document(args.path)
    paras, all_para, all_cell = collect_text(doc)
    full = all_para + "\n" + all_cell

    checks = [
        # Branding / identity.
        ("cover report title", "Weekly Status Report" in all_para),
        ("Fly91 client name present",
         "JUST UDO AVIATION PRIVATE LIMITED (Fly91)" in full),
        ("preparer org (Greatworx) present", "Greatworx" in full),
        ("account id 674351849978 present", "674351849978" in full),

        # Section 1: Cost Summary Differences.
        ("Cost Summary Differences heading",
         any("Cost Summary Differences" in p for p in paras)),
        ("cost table has Inspector row", "Inspector" in all_cell),
        ("cost table has GuardDuty row", "GuardDuty" in all_cell),
        ("cost table has Total Cost row", "Total Cost" in all_cell),
        ("cost table has Tax-Excluded total",
         "Tax Excluded Cost" in all_cell),
        ("cost-analysis decrease bullet present",
         any("decrease" in p and "compared to the previous week" in p
             for p in paras)),

        # Section 2: Disclaimer.
        ("Disclaimer heading", any(p.strip() == "Disclaimer" for p in paras)),
        ("disclaimer confidentiality text",
         "confidential between" in all_para),

        # Section 3: Contents.
        ("Contents heading", any(p.strip() == "Contents" for p in paras)),
        ("Contents lists Amazon Inspector",
         "Amazon Inspector" in all_para),

        # Section 4: BAU Matrix Overview.
        ("BAU Matrix Overview heading",
         any("BAU Matrix Overview of the resources" in p for p in paras)),
        ("Uptime subtitle",
         "Uptime of the servers and applications running on the infrastructure"
         in all_para),
        ("EC2 subsection heading",
         any("Elastic Compute Cloud (EC2)" in p for p in paras)),
        ("EC2 Server Utilization table caption",
         "All EC2 Server Utilization" in all_para),
        ("EC2 Bandwidth and Network stats table",
         "Bandwidth and Network stats (Max)" in all_para),
        ("RDS subsection heading",
         any("Relational Database Service (RDS)" in p for p in paras)),
        ("ELB subsection heading",
         any("Elastic Load Balancing (ELB)" in p for p in paras)),
        ("WAF subsection heading", any("AWS WAF" in p for p in paras)),
        ("Amazon Inspector subsection heading",
         any("Amazon Inspector" in p and p.strip().startswith("5")
             for p in paras)),
        ("Inspector findings summary",
         "findings for the last week" in all_para),
        # Exact reference phrasing for the Inspector summary (fidelity): the
        # grand total reads "400+" (not the raw "464+"), and the High/Medium
        # buckets carry the reference's "+" qualifier.
        ("Inspector total reads '400+'", "400+ findings" in all_para),
        ("Inspector reads '200+ High'", "200+ High" in all_para),
        ("Inspector reads '200+ medium'", "200+ medium" in all_para),
        ("Inspector total NOT the raw '464+'", "464+" not in all_para),
        ("Guard Duty subsection heading",
         any("Guard Duty" in p and p.strip().startswith("6") for p in paras)),
        ("GuardDuty findings summary",
         "new findings this week" in all_para),

        # Count formatting fidelity: EC2 packet / ELB / WAF counts render with
        # an uppercase thousands suffix ("952.35K"), matching the reference and
        # the uppercase "M" already used, with no lowercase "k" leaking in.
        ("uppercase K thousands suffix present (952.35K)",
         "952.35K" in all_cell),
        ("WAF uppercase K suffix present (88.72K)", "88.72K" in all_cell),
        ("no lowercase 'k' thousands suffix in tables",
         "952.35k" not in all_cell and "88.72k" not in all_cell),

        # EC2 stopped-instance bullet carries the reference stop timestamp.
        ("stopped EC2 bullet includes stop time",
         "has been put to \"Stopped\" state from" in all_para),

        # Trailer.
        ("End of Document trailer", "-- End of Document --" in all_para),

        # NEGATIVE: the obsolete Aptech-template elements must be gone.
        ("no multi-account fleet table",
         "Cost Summary Difference of All AWS Accounts" not in full),
        ("no security best practices table",
         "security best practices" not in all_para.lower()),
        ("no Resource Utilization & Alarms table",
         "Resource Utilization & Alarms" not in all_para),
    ]

    for iid in REQUIRED_EC2_IDS:
        checks.append(("EC2 instance id %s present" % iid, iid in all_cell))
    for db in REQUIRED_RDS:
        checks.append(("RDS db %s present" % db, db in all_cell))
    for alb in REQUIRED_ELB:
        checks.append(("ELB %s present" % alb, alb in all_cell))
    for acl in REQUIRED_WAF:
        checks.append(("WAF %s present" % acl, acl in all_cell))

    if args.min_images > 0:
        image_count = len(doc.inline_shapes)
        checks.append(
            ("embedded images >= %d (found %d)" % (args.min_images, image_count),
             image_count >= args.min_images))

    for extra in args.contains:
        checks.append(("contains '%s'" % extra, extra in full))

    failed = [name for name, ok in checks if not ok]
    for name, ok in checks:
        print(("PASS" if ok else "FAIL"), name)

    if failed:
        print("\nFAILED CHECKS: %d" % len(failed), file=sys.stderr)
        return 1

    print("\nALL CHECKS PASSED  tables=%d paras=%d images=%d"
          % (len(doc.tables), len(doc.paragraphs), len(doc.inline_shapes)))
    return 0


if __name__ == "__main__":
    sys.exit(main())

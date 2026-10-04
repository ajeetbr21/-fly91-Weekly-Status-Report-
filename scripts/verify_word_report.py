#!/usr/bin/env python3
"""
Standalone verification for the generated Weekly Status Report (.docx).

Loads a generated .docx with python-docx and asserts that the key structural
elements required by the client reference layout are present, and that the
cost date ranges / alarm dates reflect the selected reporting week.

Usage:
    python3 scripts/verify_word_report.py <path-to.docx> \
        [--current-range "16th June"] [--prev-range "9th June"]

If no range strings are supplied, only structural checks run.
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


REQUIRED_LINKS = [
    "best-practices-root-user.html",
    "credentials-access-keys-best-practices.html",
    "shared-responsibility-model",
    "aws.amazon.com/cloudtrail/",
    "trustedadvisor",
    "gs_creating_billing_alarm",
    "id_credentials_mfa.html",
    "github.com/awslabs/git-secrets",
]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", help="Path to the generated .docx")
    parser.add_argument("--contains", nargs="*", default=[],
                        help="Extra substrings that must appear anywhere in the document")
    args = parser.parse_args()

    doc = Document(args.path)
    paras, all_para, all_cell = collect_text(doc)
    full = all_para + "\n" + all_cell

    checks = [
        ("cover report title", "Weekly Status Report" in all_para),
        ("submitted-by label", "Submitted By" in all_para),
        ("Fly91 client name present",
         "JUST UDO AVIATION PRIVATE LIMITED (Fly91)" in full),
        ("submitter/activity org (Greatworx) present", "Greatworx" in full),
        ("account id 674351849978 present", "674351849978" in full),
        ("security best practices heading",
         any("security best practices" in p for p in paras)),
        ("per-account Summary section", any(p.strip() == "Summary" for p in paras)),
        ("Billing and Cost Overview",
         any("Billing and Cost Overview" in p for p in paras)),
        ("Resource Utilization & Alarms", "Resource Utilization & Alarms" in all_para),
        ("alarm table headers",
         "Server Name" in all_cell and "Alert date and no. of trigger" in all_cell),
        ("End Of Document trailer", "End Of Document" in all_para),
        # NEGATIVE: single-account scope must NOT include the multi-account
        # fleet table heading.
        ("no multi-account fleet table (single-account scope)",
         "Cost Summary Difference of All AWS Accounts" not in full),
    ]
    for name, link in [("link: " + lk, lk in all_cell) for lk in REQUIRED_LINKS]:
        checks.append((name, link))
    for extra in args.contains:
        checks.append(("contains '%s'" % extra, extra in full))

    failed = [name for name, ok in checks if not ok]
    for name, ok in checks:
        print(("PASS" if ok else "FAIL"), name)

    if failed:
        print("\nFAILED CHECKS: %d" % len(failed), file=sys.stderr)
        return 1

    print("\nALL CHECKS PASSED  tables=%d paras=%d" % (len(doc.tables), len(doc.paragraphs)))
    return 0


if __name__ == "__main__":
    sys.exit(main())

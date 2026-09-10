#!/usr/bin/env python3
"""Verify a generated RFQ against the configurator export it came from.

A mis-pasted configuration is the failure that actually reaches customers, and
it is invisible in a 70-row table. This checks the things a human skims past:
every module code and quantity present and in agreement with the export, no
placeholder left unfilled, and the letterhead still intact.

Exit code is 0 when everything agrees, 1 otherwise.

Usage:
    python3 check_rfq.py RFQ.docx --config CONFIG.xlsx
"""
from __future__ import annotations

import argparse
import os
import re
import sys

try:
    from docx import Document
except ImportError:
    sys.exit("python-docx is required:  pip install python-docx")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_rfq import read_config, find_config_table  # noqa: E402


def collect_doc_rows(table) -> tuple[list[tuple[str, str, str]], list[str]]:
    """Split the Configuration table into item rows and subheading labels.

    A subheading is a row whose cells were merged into one, so it reports a
    single distinct cell rather than four.
    """
    items, headings = [], []
    for row in table.rows[1:]:
        distinct = []
        for cell in row.cells:
            if not distinct or cell._tc is not distinct[-1]._tc:
                distinct.append(cell)
        texts = [c.text.strip() for c in distinct]
        if len(distinct) == 1:
            if texts[0]:
                headings.append(texts[0])
        elif any(texts):
            qty, incl, module, desc = (texts + ["", "", "", ""])[:4]
            items.append((qty, incl, module, desc))
    return items, headings


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("rfq", help="the generated .docx")
    ap.add_argument("--config", help="the configurator export it was built from")
    args = ap.parse_args(argv)

    doc = Document(args.rfq)
    problems, notes = [], []

    # --- placeholders ------------------------------------------------------
    body = doc.element.body.xml
    leftover = sorted(set(re.findall(r"\{\{(\w+)\}\}", body)))
    if leftover:
        problems.append("unfilled placeholder(s) still in the document: "
                        + ", ".join("{{%s}}" % t for t in leftover))

    # --- letterhead --------------------------------------------------------
    section = doc.sections[0]
    header_imgs = (section.first_page_header._element.xml.count("<a:blip")
                   + section.header._element.xml.count("<a:blip"))
    if header_imgs == 0:
        problems.append("no image in the page header - the letterhead/logo was lost")
    else:
        notes.append(f"letterhead images: {header_imgs}")

    for label, table_index in (("Pricing", 5), ("Terms & conditions", 7)):
        if len(doc.tables) <= table_index:
            problems.append(f"the {label} table is missing from the document")

    # --- configuration table ----------------------------------------------
    table = find_config_table(doc)
    doc_items, headings = collect_doc_rows(table)
    notes.append(f"configuration rows: {len(doc_items)}")
    if headings:
        notes.append(f"class subheadings: {len(headings)} ({', '.join(headings)})")

    empty_state = [m for q, i, m, _ in doc_items if not q and not i]
    if empty_state:
        problems.append("row(s) with neither a quantity nor an Incl. mark: "
                        + ", ".join(empty_state[:8]))
    both_state = [m for q, i, m, _ in doc_items if q and i]
    if both_state:
        problems.append("row(s) marked both charged and included: "
                        + ", ".join(both_state[:8]))

    # --- compare against the export ---------------------------------------
    if args.config:
        cfg = read_config(args.config)
        src = [(it["qty"], it["incl"], it["module"], it["description"])
               for it in cfg["items"]]

        if len(src) != len(doc_items):
            problems.append(f"row count mismatch: export has {len(src)} modules, "
                            f"document has {len(doc_items)}")

        # Grouping reorders rows, so compare as multisets rather than pairwise.
        from collections import Counter
        src_c, doc_c = Counter(src), Counter(doc_items)
        for row in sorted(set(src_c) - set(doc_c) | {r for r in src_c if src_c[r] != doc_c.get(r)}):
            problems.append(f"missing from the document: {row[2] or row[3][:40]!r} "
                            f"(qty {row[0] or '-'}, incl {row[1] or '-'})")
        for row in sorted(set(doc_c) - set(src_c)):
            problems.append(f"in the document but not the export: {row[2] or row[3][:40]!r}")

        notes.append(f"export rows: {len(src)}")

    # --- report ------------------------------------------------------------
    for note in notes:
        print(f"  {note}")
    if problems:
        print(f"\nFAILED - {len(problems)} problem(s):", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        return 1
    print("\nOK - the document matches the export.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

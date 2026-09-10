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
from build_rfq import (read_config, find_config_table, find_pricing_table,  # noqa: E402
                       find_picture_anchor, distinct_cells, PICTURE_BOX_IN,
                       _is_training_row)


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
    ap.add_argument("--expect-picture", action="store_true",
                    help="treat a missing system picture as a failure")
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

    # --- system picture ----------------------------------------------------
    anchor = find_picture_anchor(doc)
    if anchor is None:
        problems.append("the template's picture anchor paragraph is gone")
    else:
        extents = re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', anchor._p.xml)
        if not extents:
            message = "no system picture in the document"
            (problems if args.expect_picture else notes).append(message)
        else:
            width, height = (int(v) / 914400 for v in extents[0])
            notes.append(f"system picture: {width:.2f}in x {height:.2f}in")
            box_w, box_h = PICTURE_BOX_IN
            # A picture wider or taller than the layout box pushes the page
            # around, which is exactly what the fit-to-box sizing prevents.
            if width > box_w + 0.02 or height > box_h + 0.02:
                problems.append(
                    f"the picture ({width:.2f} x {height:.2f}in) overflows the "
                    f"{box_w} x {box_h}in layout box")

    # --- training line items ------------------------------------------------
    pricing = find_pricing_table(doc)
    if pricing is None:
        problems.append("no pricing table found in the document")
    else:
        trainings = [distinct_cells(r)[0].text.strip() for r in pricing.rows
                     if _is_training_row(distinct_cells(r)[0].text)]
        if trainings:
            notes.append(f"training rows: {', '.join(trainings)}")
        if len(set(trainings)) != len(trainings):
            duplicates = sorted({t for t in trainings if trainings.count(t) > 1})
            problems.append("duplicate training line(s): " + ", ".join(duplicates))
        numbers = [int(m.group(1)) for t in trainings
                   if (m := re.search(r"#(\d+)$", t))]
        if numbers and numbers != list(range(1, len(numbers) + 1)):
            problems.append(f"training numbering is not sequential: {numbers}")
        for row in pricing.rows:
            cells = distinct_cells(row)
            if _is_training_row(cells[0].text) and len(cells) > 1 \
                    and not cells[1].text.strip():
                problems.append(f"training row {cells[0].text.strip()!r} has no "
                                f"description")

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

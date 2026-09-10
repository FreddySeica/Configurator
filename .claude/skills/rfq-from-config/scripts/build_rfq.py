#!/usr/bin/env python3
"""Turn a machine-configurator export (.xls/.xlsx/.csv) into a customer-ready RFQ .docx.

The configurator writes one row per module. This script reads those rows, groups
them by their Class column, and rebuilds the "Configuration" table inside the
company RFQ template so the result can be sent to the customer as-is.

Everything the script does not understand it leaves alone: the pricing table,
terms, service options and spare-parts appendix all come through from the
template untouched, because those are commercial decisions a person makes.

Usage:
    python3 build_rfq.py CONFIG_EXPORT [-o OUT.docx] [--customer "..."] [--date ...]
                         [--subject "..."] [--description "..."] [--protocol "..."]
                         [--template PATH] [--no-group] [--json]
"""
from __future__ import annotations

import argparse
import copy
import datetime as _dt
import json
import os
import re
import sys

try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("python-docx is required:  pip install python-docx")

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, "..", "assets", "rfq_template.docx")

# Shading used by the template's own header rows - reused so subheadings look native.
SUBHEAD_FILL = "E0E0E0"

# Column headers the configurator emits. Matching is case/punctuation-insensitive.
COLUMN_ALIASES = {
    "qty":   ("#.", "#", "n.", "no.", "qty", "quantity"),
    "incl":  ("incl.", "incl", "included"),
    "module": ("module", "code", "part number", "part no.", "p/n"),
    "desc":  ("description", "descrizione", "desc"),
    "class": ("class", "classe", "category", "group"),
}


# --------------------------------------------------------------------------- #
# Reading the configurator export
# --------------------------------------------------------------------------- #
def _norm(value) -> str:
    """Collapse a cell to a comparable lowercase token."""
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def _clean(value) -> str:
    """Render a cell as the text a human would expect to see.

    Spreadsheets store quantities as floats, so a quantity of 1 arrives as 1.0.
    Printing "1.0" in a customer quote looks like a mistake, so whole numbers
    lose their decimal part while genuine fractions keep theirs.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else f"{value:g}"
    if isinstance(value, int):
        return str(value)
    return re.sub(r"[ \t]+", " ", str(value)).strip()


def read_grid(path: str) -> list[list]:
    """Return the first populated sheet as a list of rows, whatever the format."""
    ext = os.path.splitext(path)[1].lower()

    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            sys.exit("Reading .xls needs xlrd:  pip install 'xlrd>=2.0'")
        book = xlrd.open_workbook(path)
        for sheet in book.sheets():
            if sheet.nrows:
                return [[sheet.cell_value(r, c) for c in range(sheet.ncols)]
                        for r in range(sheet.nrows)]
        sys.exit(f"{path} has no populated sheet.")

    if ext in (".xlsx", ".xlsm", ".xltx"):
        try:
            import openpyxl
        except ImportError:
            sys.exit("Reading .xlsx needs openpyxl:  pip install openpyxl")
        book = openpyxl.load_workbook(path, data_only=True)
        for sheet in book.worksheets:
            if sheet.max_row and sheet.max_row > 1:
                return [list(row) for row in sheet.iter_rows(values_only=True)]
        return [list(row) for row in book.worksheets[0].iter_rows(values_only=True)]

    if ext in (".csv", ".tsv", ".txt"):
        import csv
        delim = "\t" if ext == ".tsv" else ","
        with open(path, newline="", encoding="utf-8-sig") as fh:
            return [row for row in csv.reader(fh, delimiter=delim)]

    sys.exit(f"Unsupported input format: {ext}")


def locate_header(grid: list[list]) -> tuple[int, dict]:
    """Find the header row and map our field names onto its column indexes.

    The configurator has changed its layout before (extra price columns, a title
    row above the header), so the header is located by content rather than by a
    fixed row number. That way a new export column does not silently shift the
    descriptions into the wrong place.
    """
    for idx, row in enumerate(grid[:30]):
        cells = [_norm(c) for c in row]
        mapping = {}
        for field, aliases in COLUMN_ALIASES.items():
            for col, text in enumerate(cells):
                if text in aliases and field not in mapping:
                    mapping[field] = col
        # Module + Description are the two columns the RFQ cannot be built without.
        if "module" in mapping and "desc" in mapping:
            return idx, mapping
    sys.exit(
        "Could not find the header row. Expected a row containing 'Module' and "
        "'Description' within the first 30 rows of the export."
    )


def _clean_qty(value) -> str:
    """Same as _clean, but also rescues whole numbers that arrived as text.

    CSV exports hand every cell over as a string, so a quantity of 1 shows up
    as the literal "1.0". Printing that in a customer quote reads as a defect,
    so numeric-looking quantities are normalised here. Only the quantity and
    Incl. columns go through this - module codes such as "4.0 READY" must keep
    their text exactly as the configurator wrote them.
    """
    text = _clean(value)
    if re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", text):
        number = float(text.replace(",", "."))
        return str(int(number)) if number.is_integer() else f"{number:g}"
    return text


def read_config(path: str) -> dict:
    """Parse an export into {'title': str, 'items': [...]}"""
    grid = read_grid(path)
    header_idx, cols = locate_header(grid)

    # Anything above the header that looks like a sentence is the config title,
    # e.g. "Maors Config - - Test configuration".
    title = ""
    for row in grid[:header_idx]:
        for cell in row:
            text = _clean(cell)
            if len(text) > 3:
                title = text
                break
        if title:
            break

    items = []
    for row in grid[header_idx + 1:]:
        def get(field, numeric=False):
            if field not in cols or cols[field] >= len(row):
                return ""
            raw = row[cols[field]]
            return _clean_qty(raw) if numeric else _clean(raw)

        module, desc = get("module"), get("desc")
        if not module and not desc:
            continue
        # The export ends with a totals strip ("TOT: ... EUR") that is not a module.
        if not module and re.match(r"^(tot|total)\b", _norm(desc)):
            break
        items.append({
            "qty": get("qty", numeric=True),
            "incl": get("incl", numeric=True),
            "module": module,
            "description": desc,
            "class": get("class") or "Other",
        })

    if not items:
        sys.exit(f"No configuration rows found in {path}.")
    return {"title": title, "items": items}


# --------------------------------------------------------------------------- #
# Writing into the template
# --------------------------------------------------------------------------- #
def fill_placeholders(doc, values: dict) -> list[str]:
    """Replace {{TOKEN}} runs. Returns the tokens still left unfilled.

    The template keeps each token inside a single run, so swapping the run text
    preserves the letterhead's fonts and spacing exactly.
    """
    filled, remaining = set(), set()
    for para in _all_paragraphs(doc):
        for run in para.runs:
            found = re.findall(r"\{\{(\w+)\}\}", run.text)
            if not found:
                continue
            for token in found:
                value = values.get(token)
                if value is None:
                    remaining.add(token)
                else:
                    run.text = run.text.replace("{{%s}}" % token, value)
                    filled.add(token)
    return sorted(remaining - filled)


def _all_paragraphs(doc):
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def find_config_table(doc):
    """The configuration table is the one whose header names the module column."""
    for table in doc.tables:
        header = [c.text.strip().lower() for c in table.rows[0].cells]
        if "module" in header and "description" in header:
            return table
    sys.exit(
        "The template has no Configuration table. Expected a table whose first "
        "row reads '#. | Incl. | Module | Description'."
    )


def _set_cell(cell, text: str, bold: bool | None = None):
    """Write text into a cell, keeping the prototype run's font."""
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    para = cell.paragraphs[0]
    runs = para.runs
    if not runs:
        para.add_run(text)
    else:
        runs[0].text = text
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    if bold is not None:
        for run in para.runs:
            run.bold = bold


def _shade(cell, fill: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def build_config_table(table, items: list[dict], group: bool = True) -> dict:
    """Rebuild the configuration table from the export rows.

    Row 1 of the template is a blank prototype: cloning it for every line keeps
    the borders, cell widths and font that the template already defines, which
    is far more reliable than trying to re-declare that formatting here.
    """
    prototype = copy.deepcopy(table.rows[1]._tr)
    body = table._tbl

    for row in list(table.rows)[1:]:
        body.remove(row._tr)

    def new_row():
        tr = copy.deepcopy(prototype)
        body.append(tr)
        return table.rows[-1]

    def add_item(item):
        cells = new_row().cells
        for cell, key in zip(cells, ("qty", "incl", "module", "description")):
            _set_cell(cell, item[key] if key != "description" else item["description"])

    def add_subheading(label):
        row = new_row()
        merged = row.cells[0].merge(row.cells[len(row.cells) - 1])
        _set_cell(merged, label, bold=True)
        _shade(merged, SUBHEAD_FILL)

    if not group:
        for item in items:
            add_item(item)
        return {"groups": [], "rows": len(items)}

    # Preserve the configurator's own ordering: a class appears where it first
    # appears in the export, so the machine's base unit stays at the top.
    order, buckets = [], {}
    for item in items:
        key = item["class"]
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(item)

    for key in order:
        add_subheading(key)
        for item in buckets[key]:
            add_item(item)

    return {"groups": [(k, len(buckets[k])) for k in order], "rows": len(items)}


def ordinal(day: int) -> str:
    if 11 <= day <= 13:
        return f"{day}th"
    return f"{day}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th') }"


def default_date() -> str:
    today = _dt.date.today()
    return f"{today.strftime('%B')} {ordinal(today.day)}, {today.year}"


# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("config", help="configurator export (.xls/.xlsx/.csv)")
    ap.add_argument("-o", "--output", help="output .docx (default: RFQ_<config title>.docx)")
    ap.add_argument("-t", "--template", default=DEFAULT_TEMPLATE)
    ap.add_argument("--customer", default="", help="fills the 'To:' block")
    ap.add_argument("--date", default=None, help="e.g. 'September 7th, 2026' (default: today)")
    ap.add_argument("--subject", default=None, help="text after 'Subject:'")
    ap.add_argument("--description", default=None,
                    help="machine description; use ' | ' to split across the two lines")
    ap.add_argument("--protocol", default="", help="e.g. 'PRV 260230/V_IL rev.01'")
    ap.add_argument("--no-group", action="store_true",
                    help="keep raw export order instead of grouping by Class")
    ap.add_argument("--json", action="store_true", help="print a machine-readable summary")
    args = ap.parse_args(argv)

    cfg = read_config(args.config)
    doc = Document(args.template)

    desc = args.description or cfg["title"] or ""
    desc_1, _, desc_2 = desc.partition(" | ")

    missing = fill_placeholders(doc, {
        "CUSTOMER": args.customer,
        "PROTOCOL": args.protocol,
        "DATE": args.date or default_date(),
        "SUBJECT": args.subject or "Offer as per your requirements",
        "DESCRIPTION_1": desc_1.strip(),
        "DESCRIPTION_2": desc_2.strip(),
    })

    stats = build_config_table(find_config_table(doc), cfg["items"], group=not args.no_group)

    out = args.output
    if not out:
        stem = re.sub(r"[^\w\-. ]+", "_", cfg["title"] or "configuration").strip() or "configuration"
        out = f"RFQ_{stem}.docx"
    doc.save(out)

    summary = {
        "output": os.path.abspath(out),
        "title": cfg["title"],
        "item_count": stats["rows"],
        "groups": stats["groups"],
        "unfilled_placeholders": missing,
    }
    if args.json:
        print(json.dumps(summary, indent=2))
    else:
        print(f"Wrote {summary['output']}")
        print(f"  configuration title : {cfg['title'] or '(none found)'}")
        print(f"  modules written     : {stats['rows']}")
        if stats["groups"]:
            print("  grouped by class    : "
                  + ", ".join(f"{k} ({n})" for k, n in stats["groups"]))
        if missing:
            print("  still to fill in Word: " + ", ".join(missing))
    return 0


if __name__ == "__main__":
    sys.exit(main())

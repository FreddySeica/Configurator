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
DEFAULT_SYSTEMS = os.path.join(HERE, "..", "assets", "systems")
DEFAULT_TRAININGS = os.path.join(HERE, "..", "assets", "trainings.json")
DEFAULT_THEME = os.path.join(HERE, "..", "assets", "theme.json")

# The example offer sits the machine render in a 4.91 x 3.93 in box. New photos
# are fitted inside it rather than forced to its width, so an unusually tall or
# wide picture cannot push the rest of the page around.
PICTURE_BOX_IN = (4.91, 3.93)
PICTURE_SUFFIXES = (".jpg", ".jpeg", ".png")

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
# System picture library
# --------------------------------------------------------------------------- #
def _tokens(text: str) -> list[str]:
    return [t for t in re.split(r"[^0-9a-z]+", text.lower()) if t]


def load_aliases(directory: str) -> dict:
    """Module code -> picture name overrides, keyed by a normalised code."""
    path = os.path.join(directory, "aliases.json")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            raw = json.load(fh).get("aliases", {})
    except (json.JSONDecodeError, OSError) as exc:
        print(f"warning: could not read {path}: {exc}", file=sys.stderr)
        return {}
    return {re.sub(r"[^0-9a-z]", "", str(k).lower()): v for k, v in raw.items()}


def list_systems(directory: str) -> list[str]:
    """Every picture available, by system name (the filename without extension)."""
    if not os.path.isdir(directory):
        return []
    return sorted(
        os.path.splitext(f)[0] for f in os.listdir(directory)
        if os.path.splitext(f)[1].lower() in PICTURE_SUFFIXES and not f.startswith(".")
    )


def resolve_picture(directory: str, wanted: str | None, cfg: dict) -> tuple[str | None, str]:
    """Pick the picture for this machine. Returns (path or None, explanation).

    The system name is not a field in the export, so it is inferred from the
    Base-class module code - the machine itself, e.g. VX - and from the
    configuration title. Matching is deliberately conservative: a wrong machine
    photo on a quote is worse than no photo, so anything less than a clear
    single winner is reported back rather than guessed at.
    """
    available = list_systems(directory)
    if not available:
        return None, f"no pictures in {directory}"

    by_norm = {}
    for name in available:
        by_norm.setdefault(re.sub(r"[^0-9a-z]", "", name.lower()), name)

    if wanted:
        key = re.sub(r"[^0-9a-z]", "", wanted.lower())
        if key in by_norm:
            return os.path.join(directory, _picture_file(directory, by_norm[key])), \
                   f"matched '{wanted}'"
        near = [n for n in available if key and key in re.sub(r"[^0-9a-z]", "", n.lower())]
        if len(near) == 1:
            return os.path.join(directory, _picture_file(directory, near[0])), \
                   f"'{wanted}' matched '{near[0]}'"
        return None, (f"'{wanted}' does not match any picture. Available: "
                      + ", ".join(available))

    base = next((i for i in cfg["items"] if i["class"].lower() == "base"), None)

    # An alias is a decision someone already made about this machine, so it
    # outranks anything inferred from the filenames.
    aliases = load_aliases(directory)
    if base and aliases:
        code = re.sub(r"[^0-9a-z]", "", base["module"].lower())
        if code in aliases:
            target = aliases[code]
            key = re.sub(r"[^0-9a-z]", "", str(target).lower())
            if key in by_norm:
                return os.path.join(directory, _picture_file(directory, by_norm[key])), \
                       f"alias {base['module']} -> {target}"
            return None, (f"aliases.json maps '{base['module']}' to '{target}', "
                          f"which is not in {directory}")

    # Candidates, strongest signal first: the base machine's module code, then
    # its description, then the configuration title.
    candidates = []
    if base:
        candidates.append(base["module"])
        candidates.append(base["description"])
    candidates.append(cfg.get("title", ""))

    # Words that name the product line rather than the machine. Ignoring them
    # keeps the comparison on the part that actually distinguishes a system.
    generic = {"pilot", "seica", "series", "opera"}

    scored = {}
    for weight, candidate in enumerate(reversed(candidates)):
        cand_tokens = set(_tokens(candidate))
        if not cand_tokens:
            continue
        for name in available:
            name_tokens = {t for t in _tokens(name) if t not in generic}
            if not name_tokens:
                continue
            matched = cand_tokens & name_tokens
            if not matched:
                continue
            # A filename with no words left over is describing this machine and
            # not a variant of it: for the code "VX", "Pilot VX" is exact while
            # "Pilot VX AUT" still has "aut" unaccounted for, so the plain VX
            # wins. When nothing is exact the scores stay level and the caller
            # is asked, which is what should happen for a bare "Compact" that
            # fits three different machines equally badly.
            exact = name_tokens <= cand_tokens
            score = (weight + 1, int(exact), len(matched))
            if score > scored.get(name, (0, 0, 0)):
                scored[name] = score

    if not scored:
        return None, ("could not tell which system this is. Available: "
                      + ", ".join(available))
    best = max(scored.values())
    winners = [n for n, v in scored.items() if v == best]
    if len(winners) > 1:
        return None, ("several pictures match equally (" + ", ".join(winners)
                      + "); pass --system to choose")
    return os.path.join(directory, _picture_file(directory, winners[0])), \
           f"auto-matched '{winners[0]}'"


def _picture_file(directory: str, stem: str) -> str:
    for suffix in PICTURE_SUFFIXES:
        if os.path.exists(os.path.join(directory, stem + suffix)):
            return stem + suffix
    raise FileNotFoundError(stem)


def find_picture_anchor(doc):
    """The centred paragraph under the description block holds the machine photo."""
    for table in doc.tables[:6]:
        for row in table.rows:
            for cell in row.cells:
                text = "\n".join(p.text for p in cell.paragraphs)
                if "{{DESCRIPTION" in text or "Description:" in text:
                    continue
                for para in cell.paragraphs:
                    if para.alignment is not None and int(para.alignment) == 1:
                        return para
    return None


def insert_picture(doc, path: str) -> tuple[float, float]:
    """Put the picture in the anchor paragraph, scaled to fit the layout box."""
    from docx.image.image import Image as DocxImage
    from docx.shared import Inches

    para = find_picture_anchor(doc)
    if para is None:
        raise RuntimeError("the template has no centred paragraph to hold the picture")

    image = DocxImage.from_file(path)
    box_w, box_h = PICTURE_BOX_IN
    native_w, native_h = image.width.inches, image.height.inches
    scale = min(box_w / native_w, box_h / native_h)
    width, height = native_w * scale, native_h * scale

    for run in list(para.runs):
        run._r.getparent().remove(run._r)
    para.add_run().add_picture(path, width=Inches(width), height=Inches(height))
    return round(width, 2), round(height, 2)


# --------------------------------------------------------------------------- #
# Training line items
# --------------------------------------------------------------------------- #
def load_trainings(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)["catalog"]


def parse_training_args(requested: list[str], catalog: list[dict]) -> list[dict]:
    """Turn --training arguments into the rows to write.

    Accepts a key ("adv1w") or a key with a count ("adv1w:2"). Numbered types
    are given running numbers across the whole selection, so asking for one
    2-week and one 1-week training yields ADV_TRAIN#1 and ADV_TRAIN#2 - the
    ordering a reader expects, and the one the example offer uses.
    """
    by_key = {entry["key"]: entry for entry in catalog}
    chosen, counter = [], 0
    for item in requested:
        key, _, count = item.partition(":")
        key = key.strip()
        if key not in by_key:
            sys.exit(f"Unknown training '{key}'. Available: "
                     + ", ".join(f"{e['key']} ({e['label']})" for e in catalog))
        try:
            repeat = int(count) if count else 1
        except ValueError:
            sys.exit(f"'{item}': the count after ':' must be a whole number.")
        if repeat < 1:
            sys.exit(f"'{item}': the count must be at least 1.")
        entry = by_key[key]
        for _ in range(repeat):
            part = entry["part_number"]
            if entry.get("numbered"):
                counter += 1
                part = f"{part}#{counter}"
            chosen.append({"part_number": part,
                           "description": entry["description"],
                           "quantity": "1"})
    return chosen


def distinct_cells(row) -> list:
    """The cells of a row, with merged spans collapsed to one entry.

    Word reports a merged cell once per grid column it covers. The pricing
    table spans Description across two columns, so writing by raw index puts
    the quantity on top of the description.
    """
    out = []
    for cell in row.cells:
        if not out or cell._tc is not out[-1]._tc:
            out.append(cell)
    return out


def find_pricing_table(doc):
    for table in doc.tables:
        header = [c.text.strip().lower() for c in table.rows[0].cells]
        if any(h.startswith("part number") for h in header) and \
           any(h.startswith("price") for h in header):
            return table
    return None


def _is_training_row(part_number: str) -> bool:
    return bool(re.match(r"^(install\+train|adv_train|train)", part_number.strip().lower()))


def apply_trainings(doc, rows: list[dict]) -> dict:
    """Replace the pricing table's training lines with the requested ones.

    Only the training rows are touched. The machine, NRE and freight lines are
    commercial content that belongs to whoever priced the deal, and the price
    column is left empty throughout - this decides what is offered, never what
    it costs.
    """
    table = find_pricing_table(doc)
    if table is None:
        return {"written": 0, "note": "no pricing table found - trainings skipped"}

    existing = [r for r in table.rows if _is_training_row(r.cells[0].text)]
    if not existing:
        return {"written": 0, "note": "no training rows in the template to replace"}

    from docx.table import _Row

    # The last existing training row stays put as the insertion point and is
    # removed once the new rows are in, so the replacements land exactly where
    # the old ones were - above the freight line, below NRE.
    prototype = copy.deepcopy(existing[0]._tr)
    anchor = existing[-1]._tr
    for row in existing[:-1]:
        row._tr.getparent().remove(row._tr)

    written = []
    for row in rows:
        tr = copy.deepcopy(prototype)
        anchor.addprevious(tr)
        cells = distinct_cells(_Row(tr, table))
        values = [row["part_number"], row["description"], row["quantity"], ""]
        for cell, value in zip(cells, values):
            _set_cell(cell, value)
        for extra in cells[len(values):]:
            _set_cell(extra, "")
        written.append(row["part_number"])

    anchor.getparent().remove(anchor)
    return {"written": len(written), "part_numbers": written}


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


def attach_included(items: list[dict]) -> list[dict]:
    """Bind each included row to the charged row it came with.

    The configurator lists a chargeable module and then, beneath it, whatever
    that module brings with it - rows carrying a quantity in `Incl.` rather than
    `#.`. So the twelve included rows under the base machine are what is inside
    the machine, while DONGLEPST further down is what comes with the
    programming station, not with the machine.

    That relationship is positional and would be lost by sorting the rows on
    their own Class: the customer would read that a dongle for the programming
    station is a separate Programming/Repair Stations line item, and that the
    base machine ships without the software it actually includes. Keeping each
    included row with its parent is what the hand-written offers do.

    Returns one block per charged row: {"lead": item, "included": [items]}.
    """
    blocks: list[dict] = []
    for item in items:
        is_included = bool(item["incl"]) and not item["qty"]
        if is_included and blocks:
            blocks[-1]["included"].append(item)
        else:
            # A leading orphan - an included row with nothing above it - becomes
            # its own block so it is still shown rather than silently dropped.
            blocks.append({"lead": item, "included": []})
    return blocks


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

    # Group by the Class of the charged row; anything it includes travels with
    # it. Preserve the configurator's own ordering, so a class appears where it
    # first appears in the export and the base machine stays at the top.
    order, buckets = [], {}
    for block in attach_included(items):
        key = block["lead"]["class"]
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(block)

    counts = {}
    for key in order:
        add_subheading(key)
        counts[key] = 0
        for block in buckets[key]:
            add_item(block["lead"])
            counts[key] += 1
            for included in block["included"]:
                add_item(included)
                counts[key] += 1

    return {"groups": [(k, counts[k]) for k in order], "rows": len(items)}


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
    ap.add_argument("config", nargs="?",
                    help="configurator export (.xls/.xlsx/.csv)")
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
    ap.add_argument("--system", default=None,
                    help="system name to pick a picture for, e.g. 'Pilot VX' "
                         "(default: inferred from the export)")
    ap.add_argument("--picture", default=None,
                    help="use this image file directly, ignoring the library")
    ap.add_argument("--systems-dir", default=DEFAULT_SYSTEMS,
                    help="directory of system pictures, named <system>.jpg")
    ap.add_argument("--no-picture", action="store_true",
                    help="leave the picture space empty")
    ap.add_argument("--training", action="append", default=[], metavar="KEY[:N]",
                    help="training line item to offer; repeatable, e.g. "
                         "--training install --training adv2w --training adv1w:2")
    ap.add_argument("--trainings-file", default=DEFAULT_TRAININGS)
    ap.add_argument("--theme", default=DEFAULT_THEME,
                    help="design tokens applied to the finished document")
    ap.add_argument("--no-theme", action="store_true",
                    help="leave the template's own styling untouched")
    ap.add_argument("--list-trainings", action="store_true",
                    help="print the training catalogue and exit")
    ap.add_argument("--list-systems", action="store_true",
                    help="print the available system pictures and exit")
    ap.add_argument("--json", action="store_true", help="print a machine-readable summary")
    args = ap.parse_args(argv)

    if args.list_trainings:
        for entry in load_trainings(args.trainings_file):
            number = " (numbered #1, #2, ...)" if entry.get("numbered") else ""
            print(f"{entry['key']:10} {entry['part_number']:15} {entry['label']}{number}")
        return 0
    if args.list_systems:
        names = list_systems(args.systems_dir)
        print("\n".join(names) if names else f"(no pictures in {args.systems_dir})")
        return 0

    if not args.config:
        ap.error("a configurator export is required (or use --list-trainings / "
                 "--list-systems)")

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

    # --- system picture ----------------------------------------------------
    picture = {"status": "skipped"}
    if not args.no_picture:
        if args.picture:
            path, why = args.picture, f"given explicitly ({os.path.basename(args.picture)})"
            if not os.path.exists(path):
                sys.exit(f"Picture not found: {path}")
        else:
            path, why = resolve_picture(args.systems_dir, args.system, cfg)
        if path:
            width, height = insert_picture(doc, path)
            picture = {"status": "inserted", "file": os.path.basename(path),
                       "why": why, "size_in": [width, height]}
        else:
            picture = {"status": "not inserted", "why": why}

    # --- training line items ----------------------------------------------
    trainings = {"written": 0, "note": "left as in the template"}
    if args.training:
        rows = parse_training_args(args.training, load_trainings(args.trainings_file))
        trainings = apply_trainings(doc, rows)

    # --- design tokens ------------------------------------------------------
    theming = {"applied": False, "note": "not applied"}
    if not args.no_theme and os.path.exists(args.theme):
        from apply_theme import apply_theme, load_theme
        tokens = load_theme(args.theme)
        counts = apply_theme(doc, tokens)
        theming = {"applied": True, "source": tokens.get("source", "unspecified"),
                   **counts}
    elif not args.no_theme:
        theming = {"applied": False, "note": f"no theme file at {args.theme}"}

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
        "picture": picture,
        "trainings": trainings,
        "theme": theming,
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
        if picture["status"] == "inserted":
            w, h = picture["size_in"]
            print(f"  system picture      : {picture['file']} ({picture['why']}), "
                  f"{w}in x {h}in")
        elif picture["status"] == "not inserted":
            print(f"  system picture      : NONE - {picture['why']}")
        if trainings.get("part_numbers"):
            print("  trainings offered   : " + ", ".join(trainings["part_numbers"]))
        elif trainings.get("note"):
            print(f"  trainings           : {trainings['note']}")
        if theming["applied"]:
            print(f"  theme               : {theming['headings']} headings, "
                  f"{theming['header_rows']} table headers, "
                  f"{theming['subheadings']} class subheadings")
            print(f"    tokens from       : {theming['source']}")
        if missing:
            print("  still to fill in Word: " + ", ".join(missing))
    return 0


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""Turn a machine-configurator export (.xls/.xlsx/.csv) into a customer-ready RFQ .docx.

The configurator writes one row per module. This script reads those rows and
fills the SEICA Israel commercial-offer template: the Configuration table
grouped by Class, the "What the configuration includes" capability summary, the
letterhead block, the system photo, and the training lines in Pricing.

Everything else it leaves alone. Prices, part numbers, freight terms and
service-contract figures are commitments a person makes, and the template marks
them with a placeholder colour that this script is careful not to overwrite -
so a built offer shows at a glance what still needs a human.

Usage:
    python3 build_rfq.py CONFIG_EXPORT [-o OUT.docx] [--customer "..."]
                         [--attention "..."] [--protocol "..."] [--date ...]
                         [--description "Title | one-line summary"]
                         [--training KEY[:N]] [--system NAME] [--no-group] [--json]
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

# The photo sits under the title on an A4 page with a 7.09 in text column.
# Half the column keeps it clearly subordinate to the offer itself, and fitting
# inside the box rather than forcing a width means an unusually tall or wide
# photo cannot push the letterhead down the page.
PICTURE_BOX_IN = (3.60, 2.60)
PICTURE_SUFFIXES = (".jpg", ".jpeg", ".png")

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


def _is_training_row(part_number: str, description: str = "") -> bool:
    """Whether a pricing row is a training line.

    A filled offer names them by part number (INSTALL+TRAIN, ADV_TRAIN#1). A
    blank template has not been given codes yet and identifies them by what the
    line is for, so both are accepted and a template stays usable before anyone
    has decided on numbering.
    """
    if re.match(r"^(install\+train|adv_train|train)", part_number.strip().lower()):
        return True
    text = description.strip().lower()
    return bool(re.search(r"\btraining\b|installation and production start-up", text))


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

    def row_is_training(row):
        cells = distinct_cells(row)
        return _is_training_row(cells[0].text,
                                cells[1].text if len(cells) > 1 else "")

    existing = [r for r in table.rows if row_is_training(r)]
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
        # Part number, description and quantity only. Whatever the prototype
        # row holds in the price column stays - in a blank template that is the
        # "[ 0,00 ]" placeholder, and blanking it would hide the fact that a
        # price is still owed on this line.
        for cell, value in zip(cells, [row["part_number"], row["description"],
                                       row["quantity"]]):
            set_cell(cell, value)
        written.append(row["part_number"])

    anchor.getparent().remove(anchor)
    return {"written": len(written), "part_numbers": written}


# --------------------------------------------------------------------------- #
# Writing into the template
# --------------------------------------------------------------------------- #
# The design system writes unfilled placeholders in a lighter slate than body
# text: "[ Code ]", "__ days from PO", "Recommended spare parts for ______" all
# carry it, while the service and terms tables - which ship already written -
# carry no colour at all and inherit the near-black default. The colour is
# therefore load-bearing: it tells a reader what still needs filling in.
PLACEHOLDER_COLOR = "5B6A82"


def _promote_from_placeholder(run) -> None:
    """Let a filled-in run take the body colour instead of the placeholder one.

    Only the placeholder slate is cleared. Any other colour on the run was a
    deliberate choice by whoever designed the template - a navy label, a white
    header - and is left alone.
    """
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return
    for color in rPr.findall(qn("w:color")):
        if (color.get(qn("w:val")) or "").upper() == PLACEHOLDER_COLOR:
            rPr.remove(color)


def set_text(paragraph, text: str, promote: bool = True) -> None:
    """Put `text` into a paragraph, keeping the first run's formatting.

    The template's placeholders are split across runs - Word breaks "[ Code ]"
    into "[ " and "Code ]" as it is typed and spell-checked - so a search and
    replace over run text misses them. Rewriting the paragraph from its first
    run sidesteps that entirely and keeps the design system's font and size.

    `promote` then drops the placeholder colour, because what is being written
    is real content. Without it a finished offer still reads as a form with
    blanks in it, which is worse than wrong - it looks unfinished to a customer.
    """
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for run in runs[1:]:
        run._r.getparent().remove(run._r)
    if promote and text:
        _promote_from_placeholder(runs[0])


def set_cell(cell, text: str, prototype=None, promote: bool = True) -> None:
    """Write text into a table cell, keeping its styling.

    Some letterhead value cells ship empty, with no run to inherit from. When
    that happens a run is cloned from `prototype` - the one cell that does have
    styled placeholder text - so a filled value looks like the design intended
    rather than falling back to Word's defaults.
    """
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    paragraph = cell.paragraphs[0]
    if not paragraph.runs and prototype is not None:
        paragraph._p.append(copy.deepcopy(prototype))
        for node in paragraph.runs[0]._r.findall(qn("w:t")):
            paragraph.runs[0]._r.remove(node)
    set_text(paragraph, text, promote=promote)


def _all_paragraphs(doc):
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                yield from cell.paragraphs


def distinct_cells(row) -> list:
    """Row cells with merged spans collapsed to a single entry."""
    out = []
    for cell in row.cells:
        if not out or cell._tc is not out[-1]._tc:
            out.append(cell)
    return out


def find_config_table(doc):
    """The table whose header row names the module column."""
    for table in doc.tables:
        header = [c.text.strip().lower() for c in distinct_cells(table.rows[0])]
        if "module" in header and "description" in header:
            return table
    sys.exit(
        "The template has no Configuration table. Expected a table whose first "
        "row reads '#. | Incl. | Module | Description'."
    )


def find_pricing_table(doc):
    for table in doc.tables:
        header = [c.text.strip().lower() for c in distinct_cells(table.rows[0])]
        if any(h.startswith("part number") for h in header) and \
           any(h.startswith("price") for h in header):
            return table
    return None


# --------------------------------------------------------------------------- #
# Letterhead and title
# --------------------------------------------------------------------------- #
LETTERHEAD_LABELS = {
    "to": "customer",
    "attn.": "attention",
    "attn": "attention",
    "protocol no.": "protocol",
    "protocol number": "protocol",
    "date": "date",
}


def fill_letterhead(doc, values: dict) -> list[str]:
    """Fill the To / Attn. / Protocol no. / Date block.

    The cells are found by their printed label rather than by position, so the
    block can be rearranged in Word without breaking this.
    """
    filled = []
    table = None
    for candidate in doc.tables:
        labels = {c.text.strip().lower() for r in candidate.rows
                  for c in distinct_cells(r)}
        if "to" in labels and "date" in labels:
            table = candidate
            break
    if table is None:
        return filled

    prototype = None
    for row in table.rows:
        for cell in distinct_cells(row):
            if cell.paragraphs and cell.paragraphs[0].runs and \
                    "[" in cell.text:
                prototype = copy.deepcopy(cell.paragraphs[0].runs[0]._r)
                break
        if prototype is not None:
            break

    for row in table.rows:
        cells = distinct_cells(row)
        for index, cell in enumerate(cells):
            key = LETTERHEAD_LABELS.get(cell.text.strip().lower())
            if key is None or index + 1 >= len(cells):
                continue
            value = values.get(key)
            if value:
                set_cell(cells[index + 1], value, prototype)
                filled.append(key)
    return filled


def fill_title(doc, solution: str, summary: str) -> None:
    """The first two non-empty body paragraphs are the title and its one-liner."""
    leading = [p for p in doc.paragraphs if p.text.strip()][:2]
    if leading and solution:
        set_text(leading[0], solution)
    if len(leading) > 1 and summary:
        set_text(leading[1], summary)


def strip_author_notes(doc) -> int:
    """Remove the template's instructions to whoever is filling it in.

    Lines like "Duplicate a class row and its items for each further group"
    are addressed to the person authoring an offer, not to the customer
    receiving one, and they read as a mistake in a document that has been sent.
    """
    removed = 0
    for paragraph in list(doc.paragraphs):
        text = paragraph.text.strip().lower()
        if text.startswith(("duplicate a class row",
                            "the capability summary the customer reads")):
            paragraph._p.getparent().remove(paragraph._p)
            removed += 1
    return removed


# --------------------------------------------------------------------------- #
# The Configuration table
# --------------------------------------------------------------------------- #
def attach_included(items: list[dict]) -> list[dict]:
    """Bind each included row to the charged row it came with.

    The configurator lists a chargeable module and then, beneath it, whatever
    that module brings with it - rows carrying a quantity in `Incl.` rather than
    `#.`. So the rows under the base machine are what is inside the machine,
    while a dongle further down belongs to the programming station, not to the
    machine.

    That relationship is positional and would be lost by sorting the rows on
    their own Class: the customer would read that a dongle for the programming
    station is a separate line item, and that the base machine ships without
    the software it actually includes.

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
    """Rebuild the Configuration table from the export.

    The template ships three example rows that double as prototypes: a shaded
    class row, a charged line with a quantity under "#.", and an included line
    with its quantity under "Incl.". Cloning those keeps the design system's
    shading, borders and type without this code restating any of it.
    """
    rows = table.rows
    class_proto = charged_proto = included_proto = None
    for row in rows[1:]:
        cells = distinct_cells(row)
        if len(cells) == 1:
            if class_proto is None:
                class_proto = copy.deepcopy(row._tr)
        elif len(cells) >= 2:
            if cells[0].text.strip() and charged_proto is None:
                charged_proto = copy.deepcopy(row._tr)
            elif cells[1].text.strip() and included_proto is None:
                included_proto = copy.deepcopy(row._tr)
    if charged_proto is None:
        sys.exit("The Configuration table has no example item row to copy.")
    if included_proto is None:
        included_proto = charged_proto
    if class_proto is None:
        class_proto = charged_proto

    body = table._tbl
    for row in list(rows)[1:]:
        body.remove(row._tr)

    from docx.table import _Row

    def add(prototype, values):
        tr = copy.deepcopy(prototype)
        body.append(tr)
        cells = distinct_cells(_Row(tr, table))
        for cell, value in zip(cells, values):
            set_cell(cell, value)
        return cells

    def add_item(item):
        add(charged_proto if item["qty"] else included_proto,
            [item["qty"], item["incl"], item["module"], item["description"]])

    if not group:
        for item in items:
            add_item(item)
        return {"groups": [], "rows": len(items)}

    order, buckets = [], {}
    for block in attach_included(items):
        key = block["lead"]["class"]
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(block)

    counts = {}
    for key in order:
        add(class_proto, [key])
        counts[key] = 0
        for block in buckets[key]:
            add_item(block["lead"])
            counts[key] += 1
            for included in block["included"]:
                add_item(included)
                counts[key] += 1

    return {"groups": [(k, counts[k]) for k in order], "rows": len(items)}


# --------------------------------------------------------------------------- #
# "What the configuration includes"
# --------------------------------------------------------------------------- #
# Classes that describe what the machine can do, as opposed to what it is
# packed in or plugged into. The summary is meant to be read before the line
# items, so a wooden box and a keyboard earn their place in the table above
# but not in a list of capabilities.
CAPABILITY_CLASSES = {
    "base", "hardware", "software", "boundary scan", "matrix",
    "open fix", "power supply", "cabinet/sorters",
}


def capability_lines(items: list[dict]) -> list[str]:
    """One line per quoted capability, led by its module code."""
    lines = []
    for item in items:
        if not item["qty"]:
            continue                       # included in the line above it
        if item["class"].strip().lower() not in CAPABILITY_CLASSES:
            continue
        code = item["module"] or item["description"][:20]
        lines.append(f"{code} \u2014 {item['description']}")
    return lines


def fill_capabilities(doc, lines: list[str]) -> int:
    """Replace the template's example capability lines with the real ones."""
    template_lines = [p for p in doc.paragraphs
                      if p.text.strip().startswith("[") and "CODE" in p.text]
    if not template_lines:
        return 0
    prototype = template_lines[0]
    anchor = prototype._p
    for extra in template_lines[1:]:
        extra._p.getparent().remove(extra._p)

    written = 0
    for line in lines:
        new_p = copy.deepcopy(prototype._p)
        anchor.addprevious(new_p)
        from docx.text.paragraph import Paragraph
        set_text(Paragraph(new_p, prototype._parent), line)
        written += 1
    anchor.getparent().remove(anchor)
    return written


# --------------------------------------------------------------------------- #
# Machine picture
# --------------------------------------------------------------------------- #
def insert_picture(doc, path: str) -> tuple[float, float]:
    """Place the machine photo directly under the title block.

    This design ships without a picture frame, so one is created rather than
    filled: a new centred paragraph after the one-line summary, before the
    letterhead block.
    """
    from docx.image.image import Image as DocxImage
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.text.paragraph import Paragraph

    leading = [p for p in doc.paragraphs if p.text.strip()][:2]
    if not leading:
        raise RuntimeError("the template has no title paragraph to sit under")
    after = leading[-1]

    holder = OxmlElement("w:p")
    after._p.addnext(holder)
    paragraph = Paragraph(holder, after._parent)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    image = DocxImage.from_file(path)
    box_w, box_h = PICTURE_BOX_IN
    scale = min(box_w / image.width.inches, box_h / image.height.inches)
    width = image.width.inches * scale
    height = image.height.inches * scale
    paragraph.add_run().add_picture(path, width=Inches(width), height=Inches(height))
    return round(width, 2), round(height, 2)


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
    ap.add_argument("--customer", default="", help="fills the 'To' cell")
    ap.add_argument("--date", default=None, help="e.g. 'September 7th, 2026' (default: today)")
    ap.add_argument("--description", default=None,
                    help="the title, or 'Solution name | one-line summary' "
                         "to set both lines of the title block at once")
    ap.add_argument("--protocol", default="", help="e.g. 'PRV 260230/V_IL rev.01'")
    ap.add_argument("--attention", default="", help="fills the 'Attn.' cell")
    ap.add_argument("--summary", default="",
                    help="the one-line summary under the title")
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
    ap.add_argument("--theme", default=None,
                    help="re-assert design tokens over the finished document; "
                         "not needed - the template already carries them")
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
    solution, _, one_liner = desc.partition(" | ")

    fill_title(doc, solution.strip(), one_liner.strip() or args.summary or "")
    letterhead = fill_letterhead(doc, {
        "customer": args.customer,
        "attention": args.attention,
        "protocol": args.protocol,
        "date": args.date or default_date(),
    })

    stats = build_config_table(find_config_table(doc), cfg["items"], group=not args.no_group)

    # --- what the configuration includes -----------------------------------
    lines = capability_lines(cfg["items"])
    capabilities = {"written": fill_capabilities(doc, lines), "available": len(lines)}

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
    # The template is itself the design system's output, so generated rows
    # inherit their styling by being cloned from it. Theming is only here for
    # re-asserting tokens over a template that has drifted.
    theming = {"applied": False, "note": "template carries its own styling"}
    if args.theme:
        from apply_theme import apply_theme, load_theme
        tokens = load_theme(args.theme)
        counts = apply_theme(doc, tokens)
        theming = {"applied": True, "source": tokens.get("source", "unspecified"),
                   **counts}

    notes_removed = strip_author_notes(doc)

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
        "letterhead_filled": letterhead,
        "capabilities": capabilities,
        "author_notes_removed": notes_removed,
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
        print(f"  capability lines    : {capabilities['written']}")
        if letterhead:
            print("  letterhead filled   : " + ", ".join(letterhead))
    return 0


if __name__ == "__main__":
    sys.exit(main())

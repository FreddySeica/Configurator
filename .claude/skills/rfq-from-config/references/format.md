# Export layout and template anatomy

Read this before editing `assets/rfq_template.docx` by hand or when the
configurator changes its export format. For everyday use, `SKILL.md` is enough.

## The configurator export

One sheet, one row per module. In the reference export
(`Maors Config - - Test configuration`) the sheet is named `Foglio1`:

| Row | Contents |
|---|---|
| 0 | Configuration title, merged across the sheet — e.g. `Maors Config - - Test configuration` |
| 1 | Header: `#.` · `Incl.` · `Module` · `Description` · `Class`, then price columns |
| 2…n | One module per row |
| last | Totals strip: `TOT:` and a currency code, no module |

### Columns

| Field | Header text accepted | Meaning |
|---|---|---|
| qty | `#.` `#` `n.` `no.` `qty` `quantity` | quantity for a separately charged item |
| incl | `Incl.` `incl` `included` | quantity for an item already included in the base machine |
| module | `Module` `code` `part number` `p/n` | module code, printed verbatim |
| desc | `Description` `descrizione` | printed verbatim |
| class | `Class` `classe` `category` `group` | becomes the subheading row |

`qty` and `incl` are mutually exclusive in a healthy export: a module is either
charged or included. A row with both filled, or neither, is worth querying with
the sales manager before the offer goes out.

Price columns exist in the export but the configurator leaves them empty, which
is why the skill does not attempt to price anything.

### Parsing rules that matter

- **The header is found by content, not position.** `locate_header` scans the
  first 30 rows for one containing both `Module` and `Description`. A title
  row, a blank row, or a new leading column will not shift the data.
- **Reading stops at the totals strip** — a row with no module code whose
  description starts with `TOT` or `TOTAL`.
- **Quantities are normalised, module codes are not.** Spreadsheets store 1 as
  `1.0` and CSV hands over the string `"1.0"`; both render as `1`. Only the
  quantity and Incl. columns get this treatment, so codes like `4.0 READY` and
  `M22-16/9` survive intact.
- **Rows with no class fall into `Other`**, which is also where the
  configurator puts hand-typed customer-specific items.
- **An `Incl.` row belongs to the charged row above it.** The export's order
  carries that meaning, so `attach_included` binds each included row to the
  preceding charged row and the pair moves through grouping together. The
  included row is placed by its parent's Class, never its own.

  In the reference export this puts twelve rows — everything from
  `8Z-HEADS-HR` to `CAD` — in the `Base` section alongside the VX, because they
  are what the machine includes, while `DONGLEPST` stays beside `PSTDI` under
  Programming/Repair Stations. Sorting those rows on their own Class instead
  would tell the customer the machine ships without its own software.

  An included row with no charged row above it becomes its own block rather
  than being dropped, and `check_rfq.py` reports any included row left opening
  a section.

## The RFQ template

`assets/rfq_template.docx` is the example offer with the customer's data
removed. Everything else — logos, page headers and footers, styles, the
pricing and terms boilerplate, the spare-parts appendix — is carried over
untouched, which is why the template is regenerated from a real offer rather
than rebuilt from scratch.

### Tables, in document order

| # | Table | Skill writes it? |
|---|---|---|
| 0 | `To:` block | yes — `{{CUSTOMER}}` |
| 1 | Protocol number / Date | yes — `{{PROTOCOL}}`, `{{DATE}}` |
| 2 | Subject | yes — `{{SUBJECT}}` |
| 3 | Description lines | yes — `{{DESCRIPTION_1}}`, `{{DESCRIPTION_2}}` |
| 4 | **Configuration** (`#.` · `Incl.` · `Module` · `Description`) | yes — rebuilt from the export |
| 5 | Pricing | no |
| 6 | Service Contract Options | no |
| 7 | Terms & conditions | no |
| 8 | Spare parts appendix | no |

Column widths in the Configuration table are 426 / 522 / 2210 / 7599 twips.

### Two things the template must keep

**Each `{{TOKEN}}` lives in a single run.** Word splits text across runs as
people type and spell-check, and a token broken into `{{CUST` + `OMER}}` is
invisible to the replacement pass — the placeholder would ship to the customer.
If you retype a placeholder in Word, regenerate the template with
`make_template.py` instead of trusting the edit.

**The Configuration table keeps exactly one blank body row.** `build_rfq.py`
deep-copies that row for every module, which is how borders, cell widths and
fonts stay consistent without the script re-declaring any of it. Delete the
prototype and the script has nothing to clone.

Row 0 carries `w:tblHeader` so the column titles repeat when the table breaks
across pages — a 70-row table always does.

Class subheadings are generated: the prototype row is cloned, its four cells
merged, the text set bold, and the cell shaded `E0E0E0` — the same grey the
template already uses for its other header rows, so the result looks native
rather than bolted on.

## The system picture

`assets/systems/` is the picture library. One image per machine, and the
filename without its extension *is* the system name:

```
assets/systems/
    Pilot VX.jpg
    Pilot V8.jpg
    Pilot SL.jpg
```

`.jpg`, `.jpeg` and `.png` are accepted. Adding a machine means dropping a file
in — there is no index to update and no code to change.

Two things in that directory are deliberately not library entries:
`aliases.json` (not an image) and `alternates/` (a directory). `list_systems`
filters on the image extensions, so alternate views of a machine can be kept
alongside without competing for a match. Reach one with `--picture`.

Where a machine has both a plain and an Opera-series photo, the library entry
is the Opera one and the other view sits in `alternates/`.

### How a picture is chosen

Three mechanisms, in descending order of authority:

1. **`--system` / `--picture`** — an explicit instruction for this one build.
2. **`aliases.json`** — a mapping someone already decided on, so it outranks
   anything inferred. Keys are Base-class module codes as the export writes
   them, compared without regard to case, spaces or punctuation; values are
   picture names from the directory, without the extension. A key pointing at a
   picture that is not there is reported rather than silently ignored, and a
   malformed file produces a warning and falls back to inference.
3. **Filename inference** — described below.

Reach for an alias when the configurator's code does not resemble the filename,
or when inference is ambiguous and the same question would otherwise come up on
every quote for that machine.

The export never names the system, so inference gathers candidates in order of
how much they can be trusted:

1. the module code of the `Base`-class row — the machine itself, e.g. `VX`
2. that row's description
3. the configuration title

Each candidate is split into words and compared against the words of every
filename. Product-line words (`pilot`, `seica`, `series`, `opera`) are dropped,
so the comparison turns on the part that distinguishes a machine.

A filename whose remaining words are *all* accounted for by the candidate beats
one that still has words left over. If a variant such as `Pilot VX AUT` is ever
added to the library, the bare code `VX` still resolves to `Pilot VX`: both
match on `vx`, but the variant leaves `aut` unexplained. Giving the code as
`VX AUT` picks the variant instead, since then nothing is left over on either
side and it matches more words.

When nothing is exact the scores stay level and the caller is asked — a bare
`Compact` fits `Compact TK`, `Compact SL SC` and `Compact Digital XL` equally
badly, and picking one would be a guess.

A module code cannot distinguish a machine from a variant the code does not
name. Where that matters, record it in `aliases.json` or pass `--system`.

Matching is intentionally strict. If nothing matches, or two pictures match
equally well, the script inserts nothing and says why. A quote carrying a photo
of the wrong machine is worse than one carrying no photo, so an ambiguous case
becomes a question for the sales manager rather than a guess.

Overrides: `--system "Pilot VX"` names a library entry, `--picture PATH` uses a
file directly (this is how to reach `alternates/`), `--no-picture` leaves the
space empty.

### Sizing

The example offer sits the render in a **4.91 × 3.93 in** box (`PICTURE_BOX_IN`
in `build_rfq.py`). New pictures are scaled to fit *inside* that box with their
proportions intact, so a 4:1 panorama comes out 4.91 × 1.23 in and a tall
portrait 0.98 × 3.93 in. Neither can push the rest of the page around, and
`check_rfq.py` fails if a picture ever exceeds the box.

The picture lives in the centred paragraph directly beneath the description
block. `find_picture_anchor` locates it by that centring rather than by index,
and the paragraph stays in the template even when empty — it is the anchor the
script writes into.

## The training catalogue

`assets/trainings.json` lists what can be offered. Adding a training type is a
data edit, not a code change:

```json
{
  "key": "adv1w",
  "part_number": "ADV_TRAIN",
  "label": "Advanced hardware & programming - 1 week onsite",
  "description": "One week onsite of one Seica engineer for advanced hardware and programming training session",
  "numbered": true
}
```

| Field | Purpose |
|---|---|
| `key` | what `--training` takes on the command line |
| `part_number` | the Part number column in the pricing table |
| `label` | shown by `--list-trainings`; for the sales manager, never printed to the customer |
| `description` | printed verbatim in the offer — word it the way it should read |
| `numbered` | append `#1`, `#2`, … when true |

Numbering runs across the whole selection in the order the trainings are
passed, which is why `--training install --training adv2w --training adv1w`
reproduces the example offer's `INSTALL+TRAIN`, `ADV_TRAIN#1` (two weeks),
`ADV_TRAIN#2` (one week). `--training adv1w:2` adds two of the same type.

### How the rows are replaced

`apply_trainings` finds the pricing table, identifies its existing training
rows by part number (`INSTALL+TRAIN`, `ADV_TRAIN…`, `TRAIN…`), and swaps them
for the requested set. The first of those rows is the formatting prototype; the
last stays in place as the insertion point so the new rows land exactly where
the old ones were — under NRE, above freight.

Everything else in the pricing table is untouched, and the price column is left
empty on every row the skill writes. This decides what is *offered*, never what
it costs.

Passing no `--training` at all leaves the template's own rows alone, which is
the right result when training has not been discussed.

### Merged cells

The pricing table merges Description across two grid columns, so Word reports
that cell twice when you iterate `row.cells`. Writing by raw index puts the
quantity on top of the description. Use `distinct_cells(row)` — it collapses
merged spans — for any row in that table.

## Regenerating the template from a new letterhead

When the company letterhead changes, take a real offer that already uses it:

```bash
python3 scripts/make_template.py NEW_OFFER.docx -o assets/rfq_template.docx
```

The script also lifts the machine photo out of the offer and writes it beside
the output as `extracted_system_picture.png` — rename it after the system and
move it into `assets/systems/`, since the photo belongs to the machine being
quoted rather than to the letterhead.

The script reports which placeholders it found. Anything listed under
*NOT found in the source* has to be placed by hand — open the template in
Word, type the token, and confirm it survived as one run:

```bash
python3 -c "import zipfile,re; \
print(re.findall(r'\{\{\w+\}\}', zipfile.ZipFile('assets/rfq_template.docx') \
.read('word/document.xml').decode()))"
```

Then prove the new template works end to end before anyone quotes from it:

```bash
python3 scripts/build_rfq.py SOME_EXPORT.xls -o /tmp/check.docx --customer "Test" \
  --training install --training adv2w
python3 scripts/check_rfq.py /tmp/check.docx --config SOME_EXPORT.xls --expect-picture
```

## If the Configuration table is renamed

`find_config_table` matches a table whose first row contains both `module` and
`description` (case-insensitive). If the company renames those columns, add the
new wording to `COLUMN_ALIASES` in `build_rfq.py` and to that match — keeping
the old spellings alongside, so archived exports still build.

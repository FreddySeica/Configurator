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

`assets/rfq_template.docx` is the SEICA Israel commercial-offer template, A4,
Arial, carrying the letterhead logo in the page header and the company details
in the footer. It is the design system's own output, so anything the script
generates is cloned from a row or paragraph already in it and inherits the
styling rather than restating it.

### Blocks, in document order

| Block | Skill writes it? |
|---|---|
| Title paragraph | yes — from `--description` before the `\|` |
| One-line summary | yes — after the `\|`, or `--summary` |
| *(photo)* | yes — a centred paragraph it creates here |
| `To` | **never** — there is no flag for it and no label mapping |
| `Attn.` | blank unless `--attention` is passed |
| `Protocol no.` | blank unless `--protocol` is passed |
| `Date` | yes — today, unless `--date` overrides |
| **Configuration** table | yes — rebuilt from the export |
| Pricing table | training rows only; part numbers and prices left alone |
| **What the configuration includes** | yes — one line per quoted capability |
| Service Contract Options | no |
| Terms & Conditions | no |
| Signature block | no |

`To` is absent from `LETTERHEAD_LABELS` rather than merely left unsupplied:
with no route from that label to a value, the cell cannot be written even by
mistake. The customer's name is written by whoever sends the offer.

Of the rest, only `Date` is filled without being asked for; `Attn.` and
`Protocol no.` stay as the template left them unless their flag is passed.

The Configuration grid is 567 / 624 / 1985 / 7029 twips.

### The placeholder colour is load-bearing

The design writes unfilled values in slate `5B6A82` and real content with no
colour override, inheriting the near-black default. Compare the Terms table
(`__ days from PO`, in slate) with the Service Contract table (already written,
no override).

`set_text` therefore *promotes* anything it fills: it drops a `w:color` of
exactly `5B6A82` and leaves every other colour alone, since those were
deliberate — a navy class row, a white header. The result is that a built offer
is grey exactly where a person still owes something, most visibly every
`[ 0,00 ]`. `PLACEHOLDER_COLOR` in `build_rfq.py` holds the value.

### Prototype rows

The Configuration table ships three example rows that double as prototypes, and
`build_config_table` picks them out by shape rather than by index:

- **class row** — a row merged into one cell, shaded `EAF2FA`
- **charged row** — a quantity under `#.`
- **included row** — a quantity under `Incl.`

Deleting all three leaves nothing to clone and the build stops with a message.
The four `[ CODE ] — [ … ]` lines under *What the configuration includes* work
the same way: the first is the prototype, the rest are removed.

### Author notes are stripped

Two lines in the template address whoever is filling it in — "Duplicate a class
row and its items…" and "The capability summary the customer reads…". They are
instructions to the author, not to the customer, and read as a mistake in a
document that has been sent, so `strip_author_notes` removes them from the
output while leaving them in the template.

### Placeholders are split across runs

Word breaks `[ Code ]` into `[ ` and `Code ]` as it is typed and spell-checked,
so a search-and-replace over run text misses them. `set_text` rewrites the
paragraph from its first run instead, which sidesteps the splitting and keeps
the font and size that run carries.

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

### Sizing and placement

This design ships **no picture frame**, so the script creates one: a centred
paragraph inserted directly after the title block, before the letterhead table.

Photos are fitted inside a **3.60 × 2.60 in** box (`PICTURE_BOX_IN` in
`build_rfq.py`) with proportions intact, so a 4:1 panorama comes out shorter and
a tall portrait narrower. Half the 7.09 in text column keeps the photo clearly
subordinate to the offer, and fitting rather than forcing a width means an
odd-shaped image cannot push the letterhead down the page. `check_rfq.py` fails
if a picture ever exceeds the box.

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

## Design tokens

`assets/theme.json` records the Seica Israel palette and type scale:

```json
"color": { "brand_dark": "1A2B5A", "brand": "0082B5", "brand_light": "EAF2FA",
           "neutral": "5B6A82", "heading_text": "0082B5",
           "body_text": "0B1220", "placeholder": "5B6A82" },
"table": { "header_fill": "1A2B5A", "header_text": "FFFFFF",
           "subheading_fill": "EAF2FA", "subheading_text": "1A2B5A" }
```

These come from the design system itself, read out of the template it produced.
**Nothing has to apply them for an offer to look right** — the template carries
its own styling and generated rows inherit it by being cloned. The file exists
so the palette is readable without opening Word, and so `apply_theme.py` can
re-assert it over a template that has drifted. `--theme PATH` runs that pass;
it is off by default.

`placeholder` is the exception to "reference only": `build_rfq.py` reads that
colour to know what counts as unfilled. See **The placeholder colour is
load-bearing** above.

`apply_theme.py` restyles the body typeface, section headings, table header rows
and class subheadings. It never touches page headers and footers (fixed logo
artwork) or font sizes (the title is larger than the body on purpose, and
flattening sizes would destroy the hierarchy while claiming to be a restyle).

## The capability copy library

`assets/capabilities.json` holds the customer-facing copy for *What the
configuration includes*. The section explains what the machine does for the
customer; the Configuration table already lists what is in it, so this file is
what keeps the two from saying the same thing twice.

```json
{
  "lead_in": "The solution offered is N°1 unit {machine}, equipped with:",
  "furthermore_lead_in": "Furthermore the offer includes:",
  "capabilities": {
    "OBPFPDVC": {
      "text": "1 Seica general purpose programmer option for On-board programming ...",
      "aliases": ["OBPDVC4"]
    }
  }
}
```

| Field | Purpose |
|---|---|
| `text` | the line the customer reads; lead with what they get |
| `bullets` | sub-points, indented one level under the line |
| `group` | `main` (default) or `furthermore` |
| `aliases` | other codes reusing this copy; the first match wins and later aliases are skipped |
| `hidden` | keep the copy, leave it out of the section |

`{machine}` in `lead_in` is replaced with the solution name — the part of
`--description` before the `|`.

### Two gates, and an override

```json
"include_classes": ["Hardware", "Programming/Repair Stations"],
"class_overrides": { "PRBVIEW": "Hardware", "V8OFX": "Hardware" }
```

`capability_entries` walks the configuration in export order and writes a module
only if **both** gates pass:

1. its class is in `include_classes` — the section is about what the machine
   does, so software licences, packaging, peripherals and automation options are
   left to the table;
2. the library has copy for its code.

An entry with `group: "furthermore"` skips gate 1. Those are offer extras, not
machine features, and excluding them by class would be wrong.

`hidden: true` is the third way out: the copy exists and stays on file, but the
line is not written. It is for items that are genuinely in the offer without
being capabilities - `KEL` (a keyboard) and `M22-16/9` (a monitor) both carry
it, since listing them beside the laser sensors weakens the section. They are
reported as `held_back` rather than dropped silently, so the decision stays
visible in the build output:

```
held back from it   : KEL, M22-16/9
```

A module that passes gate 1 but fails gate 2 is reported in `uncovered` and
printed after the build — that report is how the library grows, naming exactly
what a customer would have read about and did not. There is deliberately **no**
fallback to the export description: filling a gap that way restores the mirror
this section exists to avoid.

`class_overrides` is applied by `apply_class_overrides` immediately after the
export is read, before anything else looks at a class. So a refiled module moves
in the Configuration table's class rows *and* in this section — one correction
rather than two that can drift. The build prints each move:

```
reclassified        : V8OFX: Open fix -> Hardware; PRBVIEW: Mech.Tools -> Hardware
```

Note the side effect, which is usually wanted: moving the only member of a class
removes that class row from the table entirely. `Open fix` disappears once
`V8OFX` becomes Hardware.

### Sub-bullets

The template defines a single bullet level (`numId` 1, `ilvl` 0, en-dash).
Sub-points reuse that list and add a 1134-twip left indent rather than a second
level being added to `numbering.xml` — editing the design system's own numbering
for the sake of one section is a bigger change than the visual difference
justifies.

The *Furthermore the offer includes:* lead-in is the same paragraph with its
`numPr` removed from its `pPr`, since it introduces the group rather than being
an item in it.

### Writing the copy

Every figure should trace to the export description or a previous offer. The
file's header says so, and it matters: this text goes to a customer, and an
invented rating is a commitment the machine may not meet. Entries not yet
reviewed by Sales are marked as drafts there.

## Replacing the template

`assets/rfq_template.docx` comes from the SEICA Israel design system in Claude
Design. To adopt a new revision, export it from there and drop it in — there is
no conversion step, because the script reads the template by structure rather
than by fixed positions.

What a replacement must keep for the build to work:

- a table whose header row reads `#.` / `Incl.` / `Module` / `Description`
- inside it, one shaded merged **class row**, one **charged** example row
  (quantity under `#.`) and one **included** example row (quantity under
  `Incl.`) — these are the prototypes that get cloned
- a letterhead table containing the labels `To` and `Date` (and optionally
  `Attn.`, `Protocol no.`), each with its value cell to the right
- at least one `[ CODE ] — [ … ]` line under a *What the configuration includes*
  heading
- a pricing table with `Part number` and `Price` in its header row, and a row
  whose description mentions training

Then prove it before anyone quotes from it:

```bash
python3 scripts/build_rfq.py SOME_EXPORT.xls -o /tmp/check.docx \
  --attention "Test" --training install --training adv2w
python3 scripts/check_rfq.py /tmp/check.docx --config SOME_EXPORT.xls --expect-picture
```

A previous `make_template.py` blanked a filled offer back into a template. It
was written for the older template's `{{TOKEN}}` placeholders and is gone: the
design system is the source of truth now, so a new template is exported rather
than derived.

## If the Configuration table is renamed

`find_config_table` matches a table whose first row contains both `module` and
`description` (case-insensitive). If the company renames those columns, add the
new wording to `COLUMN_ALIASES` in `build_rfq.py` and to that match — keeping
the old spellings alongside, so archived exports still build.

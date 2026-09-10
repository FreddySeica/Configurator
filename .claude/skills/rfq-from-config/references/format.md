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

## Regenerating the template from a new letterhead

When the company letterhead changes, take a real offer that already uses it:

```bash
python3 scripts/make_template.py NEW_OFFER.docx -o assets/rfq_template.docx
```

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
python3 scripts/build_rfq.py SOME_EXPORT.xls -o /tmp/check.docx --customer "Test"
python3 scripts/check_rfq.py /tmp/check.docx --config SOME_EXPORT.xls
```

## If the Configuration table is renamed

`find_config_table` matches a table whose first row contains both `module` and
`description` (case-insensitive). If the company renames those columns, add the
new wording to `COLUMN_ALIASES` in `build_rfq.py` and to that match — keeping
the old spellings alongside, so archived exports still build.

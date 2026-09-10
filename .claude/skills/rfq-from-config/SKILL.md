---
name: rfq-from-config
description: Turn a machine-configurator export (.xlsx/.xls/.csv listing modules with #., Incl., Module, Description, Class columns) into a customer-ready RFQ / commercial-offer Word document on the company letterhead template. Use this whenever the user mentions building a quote, an RFQ, a commercial offer, a customer proposal, or "sending a configuration to a customer" — and also whenever they hand over a configurator or config export and ask to put it into a template, into the RFQ, into Word, or "make it ready to send". Trigger even if they only say something like "here's the config for Valid, prep the offer" without naming the template or the file format.
---

# RFQ from configurator export

A sales manager configures a machine in the configurator tool, exports the
result as a spreadsheet, and needs to send the customer a formal offer. Doing
that by hand means pasting 50–70 rows into Word and repairing the formatting
every time. This skill does the paste, and deliberately stops short of the
commercial decisions.

**What the skill fills in:** the `To:` block, protocol number, date, subject,
description lines, and the whole Configuration table (grouped by Class).

**What the skill leaves alone — on purpose:** the Pricing table, Service
Contract Options, Terms & conditions, and the spare-parts appendix. Prices,
delivery dates and payment terms are commitments to a customer; a wrong number
there is far more costly than a missing one, so a person fills those in Word
before sending. Say so when handing back the file rather than leaving the user
to discover it.

## Workflow

### 1. Locate the export and read it before asking anything

Run the script with `--json` against the export to see what is actually in it:

```bash
python3 scripts/build_rfq.py CONFIG.xlsx --json -o /tmp/probe.docx
```

The summary reports the configuration title, the module count, and the Class
groups. Knowing the machine is a Pilot VX with 56 modules lets you propose a
sensible subject line instead of asking the user to invent one, which is the
difference between one question and four.

Dependencies, if the script complains: `pip install python-docx openpyxl xlrd`
(`xlrd` only matters for the older `.xls` exports).

### 2. Ask only for what the export cannot tell you

The customer name is the one field never present in the export — always ask
for it. For everything else, propose a default and let the user correct it:

| Field | Sensible default |
|---|---|
| `--customer` | *ask* — company name, optionally `/ Attn: ...` |
| `--date` | today, already formatted as `September 10th, 2026` |
| `--subject` | `Offer for <machine> as per your requirements` |
| `--description` | the configuration title from the export |
| `--protocol` | ask, or leave blank for the user to type in Word |

Ask for these in a single round, not one at a time. If the user has already
supplied a value in their message, do not ask again.

### 3. Build the document

```bash
python3 scripts/build_rfq.py CONFIG.xlsx \
  -o "RFQ_<Customer>_<Machine>.docx" \
  --customer "Valid Ltd. / Attn: Mr. Cohen" \
  --protocol "PRV 260230/V_IL rev.01" \
  --subject "Offer for Custom Functional Test Solution as per your requirements" \
  --description "Functional Flying Probe Test Solution: Pilot VX (Opera) | Vertical double sided Flying Probe Tester"
```

`--description` splits on ` | ` across the template's two description lines.
Pass `--no-group` to keep the raw export order instead of grouping by Class.

Name the output after the customer and machine. `RFQ_Valid_PilotVX.docx` is
findable six months later; `RFQ_Maors Config - - Test configuration.docx`
(the default, derived from the export title) is not.

### 4. Check the result before handing it over

The script reports its own row count, but it cannot tell whether the numbers
are *right*. Confirm that the modules written match the export's row count and
that no placeholder survived:

```bash
python3 scripts/check_rfq.py OUT.docx --config CONFIG.xlsx
```

This compares every module code and quantity in the document against the
export and fails loudly on a mismatch, an unfilled `{{TOKEN}}`, or a lost
letterhead image. A silent mis-paste is the failure mode that actually reaches
customers, so run it rather than eyeballing the row count.

### 5. Hand back the file and say what is still open

Deliver the `.docx` and state plainly which sections still need the user:
prices in the Pricing table, service-contract figures, and the protocol number
if it was left blank. Do not describe the offer as ready to send.

## Working with the Class grouping

The export's Class column (Base, Hardware, Software, PC/Peripherals,
Mech.Tools, Packaging, Other…) becomes a shaded subheading row spanning the
table. Groups appear in the order they first occur in the export, which keeps
the base machine at the top where a reader expects it, rather than imposing an
alphabetical order that would bury it.

`Other` is where the configurator puts customer-specific line items typed in
by hand — special power supplies, extra magazines, third-party instruments.
These often have informal module codes (`two PCW11TOW`, `One Doungle`). They
are legitimate quote lines, so pass them through as written; tidy the wording
only if the user asks, and never silently drop a row you do not recognise.

If a row has neither `#.` nor `Incl.` filled, or has both, that is usually a
configurator glitch. Mention it — it is cheap to check before the offer goes
out and awkward afterwards.

## When the template or export changes

Both are read by content, not by fixed positions: the script finds the header
row by looking for `Module` + `Description`, and finds the Configuration table
by its header text. An added column or a shifted row will not break it.

Two changes do require attention, and `references/format.md` covers both:
replacing the template with a newer company letterhead, and what to do if the
Configuration table's columns are renamed. Read it before editing
`assets/rfq_template.docx` by hand — the template carries single-run
`{{TOKEN}}` placeholders and one blank prototype row that the script clones,
and Word will happily split those tokens across runs if they are retyped.

## Files

- `scripts/build_rfq.py` — parse export, fill template, write the RFQ
- `scripts/check_rfq.py` — verify the output against the export
- `assets/rfq_template.docx` — letterhead, empty Configuration table, pricing
  and terms boilerplate
- `references/format.md` — export layout, template anatomy, how to re-derive
  the template from a new company letterhead

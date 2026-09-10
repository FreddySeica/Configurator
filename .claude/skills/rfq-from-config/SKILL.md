---
name: rfq-from-config
description: Turn a machine-configurator export (.xlsx/.xls/.csv listing modules with #., Incl., Module, Description, Class columns) into a customer-ready RFQ / commercial-offer Word document on the company letterhead template, including the system photo and the training line items. Use this whenever the user mentions building a quote, an RFQ, a commercial offer, a customer proposal, or "sending a configuration to a customer" — and also whenever they hand over a configurator or config export and ask to put it into a template, into the RFQ, into Word, or "make it ready to send". Trigger even if they only say something like "here's the config for Valid, prep the offer" without naming the template or the file format.
---

# RFQ from configurator export

A sales manager configures a machine in the configurator tool, exports the
result as a spreadsheet, and needs to send the customer a formal offer. Doing
that by hand means pasting 50–70 rows into Word and repairing the formatting
every time. This skill does the paste, and deliberately stops short of the
commercial decisions.

**What the skill fills in:** the `To:` block, protocol number, date, subject,
description lines, the Configuration table (grouped by Class), the system
photo, and the training line items in the Pricing table.

**What the skill leaves alone — on purpose:** every price, the NRE and freight
lines, Service Contract Options, Terms & conditions, and the spare-parts
appendix. Prices, delivery dates and payment terms are commitments to a
customer; a wrong number there is far more costly than a missing one, so a
person fills those in Word before sending. Say so when handing back the file
rather than leaving the user to discover it.

## Workflow

### 1. Read the export before asking anything

```bash
python3 scripts/build_rfq.py CONFIG.xlsx --json -o /tmp/probe.docx
```

The summary reports the configuration title, the module count, the Class
groups, and which system picture was matched. Knowing the machine is a Pilot VX
with 56 modules lets you propose a sensible subject line instead of asking the
user to invent one, which is the difference between one question and four.

Dependencies, if the script complains: `pip install python-docx openpyxl xlrd`
(`xlrd` only matters for the older `.xls` exports).

### 2. Ask for what the export cannot tell you — in one round

Three things are genuinely unknowable from the spreadsheet. Ask them together,
not one at a time, and skip any the user already gave you.

**The customer.** Never inferable. Always ask.

**The trainings.** The export says nothing about them, and the number varies
per deal. Show the catalogue and ask how many of each:

```bash
python3 scripts/build_rfq.py --list-trainings
```

```
install    INSTALL+TRAIN   Installation + basic training - 1 week onsite
adv2w      ADV_TRAIN       Advanced hardware & programming - 2 weeks onsite
adv1w      ADV_TRAIN       Advanced hardware & programming - 1 week onsite
```

Ask it the way a colleague would — "installation week plus how many advanced
training weeks?" — rather than making the user learn the keys. Then translate
their answer into `--training` flags. Advanced trainings get numbered
`ADV_TRAIN#1`, `#2`, `#3`… automatically in the order you pass them, so pass
the longer sessions first if the user lists them that way.

**The system picture, but only if the match is unsure.** The script infers the
machine from the Base-class module code and reports what it picked. If it says
`NONE`, show the user what is available and ask:

```bash
python3 scripts/build_rfq.py --list-systems
```

For everything else, propose a default rather than asking:

| Field | Sensible default |
|---|---|
| `--date` | today, already formatted as `September 10th, 2026` |
| `--subject` | `Offer for <machine> as per your requirements` |
| `--description` | the configuration title from the export |
| `--protocol` | ask, or leave blank for the user to type in Word |

### 3. Build the document

```bash
python3 scripts/build_rfq.py CONFIG.xlsx \
  -o "RFQ_<Customer>_<Machine>.docx" \
  --customer "Valid Ltd. / Attn: Mr. Cohen" \
  --protocol "PRV 260230/V_IL rev.01" \
  --subject "Offer for Custom Functional Test Solution as per your requirements" \
  --description "Functional Flying Probe Test Solution: Pilot VX (Opera) | Vertical double sided Flying Probe Tester" \
  --training install --training adv2w --training adv1w
```

`--description` splits on ` | ` across the template's two description lines.
`--training KEY:N` adds N of the same type (`--training adv1w:2`). Other flags
worth knowing: `--system "Pilot VX"` to override the picture match, `--picture
PATH` for a one-off image outside the library, `--no-picture` to leave the space
empty, and `--no-group` to keep the raw export order instead of grouping by
Class.

Omitting `--training` entirely leaves whatever the template already has, which
is the right behaviour when the user has not mentioned training at all.

Name the output after the customer and machine. `RFQ_Valid_PilotVX.docx` is
findable six months later; `RFQ_Maors Config - - Test configuration.docx`
(the default, derived from the export title) is not.

### 4. Check the result before handing it over

The script reports its own row count, but it cannot tell whether the numbers
are *right*.

```bash
python3 scripts/check_rfq.py OUT.docx --config CONFIG.xlsx --expect-picture
```

This compares every module code and quantity against the export and fails
loudly on a mismatch, an unfilled `{{TOKEN}}`, a lost letterhead, a picture
that overflows its layout box, a duplicated training line, or training numbers
that skip. A silent mis-paste is the failure mode that actually reaches
customers, so run it rather than eyeballing the row count. Drop
`--expect-picture` if the offer is deliberately going out without a photo.

### 5. Hand back the file and say what is still open

Deliver the `.docx` and state plainly which sections still need the user: every
price in the Pricing table, service-contract figures, and the protocol number
if it was left blank. Do not describe the offer as ready to send.

## The system picture

`assets/systems/` holds one image per machine, named after the system. The
filename *is* the system name, so adding a machine means dropping a file in and
nothing else. `.jpg`, `.jpeg` and `.png` all work. Currently stocked:

```
Compact Digital XL   Compact SL SC   Compact TK
Pilot VX             Valid LR        Valid SL
```

Where a machine has both a plain and an Opera-series photo, the library holds
the Opera one; the other views live in `assets/systems/alternates/`, which is
out of the matching pool so it cannot make a match ambiguous. To use one, pass
`--picture "assets/systems/alternates/VALID-LR-02.jpg"`.

Matching looks for the Base module code as a whole word in the filename, so
`VX` finds `Pilot VX` and `TK` finds `Compact TK`. It is deliberately
conservative: when several pictures match equally well — a bare `Compact` fits
three of them — it inserts nothing and says so. A quote carrying a photo of the
wrong machine is worse than one carrying no photo at all.

When the configurator's code does not resemble the filename, record it once in
`assets/systems/aliases.json` instead of passing `--system` forever:

```json
"aliases": { "TK": "Compact TK", "VALID-LR": "Valid LR" }
```

That file starts empty. If a build reports `NONE` for the picture and the user
tells you which machine it is, offer to add the alias — it is the difference
between fixing it once and answering the same question every quote.

Pictures are scaled to fit inside the template's 4.91 × 3.93 in box without
distortion, so a wide or tall photo comes out smaller rather than shoving the
page around.

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

## When the template, catalogue or export changes

Both files are read by content, not by fixed positions: the script finds the
export header by looking for `Module` + `Description`, and finds the tables by
their header text. An added column or a shifted row will not break it.

Adding a training type is a data edit — a new entry in `assets/trainings.json`,
no code change. `references/format.md` covers that, the picture library, and
how to re-derive the template from a newer company letterhead. Read it before
editing `assets/rfq_template.docx` by hand — the template carries single-run
`{{TOKEN}}` placeholders and prototype rows that the script clones, and Word
will happily split those tokens across runs if they are retyped.

## Files

- `scripts/build_rfq.py` — parse export, fill template, write the RFQ
- `scripts/check_rfq.py` — verify the output against the export
- `scripts/make_template.py` — re-derive the template from a new letterhead
- `assets/rfq_template.docx` — letterhead, empty Configuration table, pricing
  and terms boilerplate
- `assets/trainings.json` — the training catalogue
- `assets/systems/` — one picture per machine, filename = system name;
  `aliases.json` for module codes that do not resemble it, `alternates/` for
  other views
- `references/format.md` — export layout, template anatomy, catalogue format

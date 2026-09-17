---
name: rfq-from-config
description: Turn a machine-configurator export (.xlsx/.xls/.csv listing modules with #., Incl., Module, Description, Class columns) into a customer-ready RFQ / commercial offer on the SEICA Israel Word template, filling the Configuration table, the "What the configuration includes" summary, the letterhead and the system photo. Use this whenever the user mentions building a quote, an RFQ, a commercial offer, a customer proposal, or "sending a configuration to a customer" — and also whenever they hand over a configurator or config export and ask to put it into a template, into the offer, into Word, or "make it ready to send". Trigger even if they only say something like "here's the config for Valid, prep the offer" without naming the template or the file format.
---

# RFQ from configurator export

A sales manager configures a machine in the configurator tool, exports a
spreadsheet, and needs a formal offer to send. Doing that by hand means pasting
50–70 rows into Word and repairing the formatting every time. This skill does
the paste, and deliberately stops short of the commercial decisions.

**What the skill fills:** the title and its one-line summary, the To / Attn. /
Protocol no. / Date block, the system photo, the **Configuration** table, the
**What the configuration includes** summary, and the training lines in Pricing.

**What it leaves for a person — on purpose:** every price, the part numbers in
the Pricing table, freight and incoterms, service-contract figures, and the
blanks in Terms & Conditions. Prices and delivery terms are commitments to a
customer; a wrong number there costs far more than a missing one. Say so when
handing the file back rather than letting the user discover it.

## The placeholder convention

The template writes anything still to be filled in a lighter slate
(`5B6A82`) and real content in the default near-black. That colour is
load-bearing: it is how someone scanning the offer sees what is still owed.

So everything the skill writes is promoted out of that slate, and everything it
deliberately does not fill keeps it. A finished build therefore shows grey
exactly where a person still has work to do — most visibly every `[ 0,00 ]` in
the Pricing table. Do not "tidy up" leftover brackets; they are the worklist.

## Workflow

### 1. Read the export before asking anything

```bash
python3 scripts/build_rfq.py CONFIG.xlsx --json -o /tmp/probe.docx
```

The summary reports the configuration title, module count, Class groups, which
system picture matched, and how many capability lines it will write. Knowing
the machine is a Pilot VX with 56 modules lets you propose a title instead of
asking the user to invent one — the difference between one question and four.

Dependencies, if the script complains: `pip install python-docx openpyxl xlrd`
(`xlrd` only matters for the older `.xls` exports).

### 2. Ask once, for what the spreadsheet cannot say

**The customer** — never inferable. Always ask, along with the `Attn.` contact.

**The trainings** — the export says nothing about them and the count varies per
deal. Show the catalogue and ask in plain terms ("installation week plus how
many advanced training weeks?"):

```bash
python3 scripts/build_rfq.py --list-trainings
```

Then translate the answer into `--training` flags. Advanced trainings are
numbered `ADV_TRAIN#1`, `#2`, … in the order passed, so pass the longer sessions
first if the user lists them that way. Omitting `--training` entirely leaves the
template's own placeholder row, which is right when training has not come up.

**The picture, only if the match is unsure.** The script infers the machine from
the Base module code and reports what it picked; if it says `NONE`, show
`--list-systems` and ask.

Everything else gets a proposed default rather than a question: today's date,
the configuration title as the solution name, and a blank protocol number for
the user to type in Word.

### 3. Build

```bash
python3 scripts/build_rfq.py CONFIG.xlsx \
  -o "Offer_<Customer>_<Machine>.docx" \
  --customer "Valid Ltd." --attention "Mr. Cohen" \
  --protocol "PRV 260230/V_IL rev.01" \
  --description "Flying Probe Test Solution: Pilot VX | Dual-side functional and in-circuit test for vertical double-sided boards" \
  --training install --training adv2w --training adv1w
```

`--description` splits on ` | ` into the title and the line under it.
`--training KEY:N` adds N of one type. `--system` overrides the picture match,
`--picture PATH` uses a file outside the library, `--no-picture` omits it, and
`--no-group` keeps raw export order instead of grouping by Class.

Name the output after the customer and machine. `Offer_Valid_PilotVX.docx` is
findable six months later; the default, derived from the export title, is not.

### 4. Check before handing it over

```bash
python3 scripts/check_rfq.py OUT.docx --config CONFIG.xlsx --expect-picture
```

This compares every module code and quantity against the export and fails on a
mismatch, a lost letterhead, a photo that overflows its box, a duplicated or
misnumbered training line, or an included row separated from its parent. It
also lists the placeholders still awaiting a person, which is the handover note
you want. A silent mis-paste is the failure that actually reaches customers, so
run it rather than eyeballing the table.

### 5. Hand back and say what is open

Deliver the `.docx` and name what still needs the user: the prices, the part
numbers in Pricing, freight and incoterm, the service-contract figures, and the
blanks in Terms. Do not describe the offer as ready to send.

## Included modules travel with their parent

The configurator lists a chargeable module and then, beneath it, whatever that
module brings along — rows with a quantity under `Incl.` rather than `#.`. That
relationship is positional, so those rows are placed under their parent and take
the parent's Class, never their own.

The twelve included rows under the base machine therefore sit in the **Base**
section with it, even though their own Class says Hardware or Software: they are
what is inside the machine. `DONGLEPST`, listed under the programming station,
stays with that station instead. Get this wrong and the offer tells the customer
the machine ships without the software it includes.

`Other` is where the configurator puts hand-typed customer-specific lines
(`two PCW11TOW`, `One Doungle`). Pass them through as written; tidy the wording
only if asked, and never drop a row you do not recognise. A row with neither
`#.` nor `Incl.` filled, or with both, is usually a configurator glitch — worth
mentioning before the offer goes out.

## What the configuration includes

The summary takes one line per **quoted capability**: charged (`#.`) modules
whose Class is Base, Hardware, Software, Boundary scan, Matrix, Open fix, Power
Supply or Cabinet/sorters. Included rows are skipped — they are inside the lines
above them — and so are Packaging, PC/Peripherals, Mech.Tools and Programming
stations, because a wooden box and a keyboard earn a place in the table but not
in a list of what the machine can do. That filter is `CAPABILITY_CLASSES` in
`build_rfq.py`; widen it there if the user wants more.

## The system picture

`assets/systems/` holds one image per machine, named after the system — the
filename *is* the system name, so adding a machine means dropping a file in:

```
Compact Digital XL   Compact SL SC   Compact TK   Pilot BT
Pilot BTP            Pilot BTV       Pilot V8     Pilot VX
Valid LR             Valid SL
```

Other views live in `assets/systems/alternates/`, outside the matching pool so
they cannot make a match ambiguous; reach one with `--picture`.

Matching compares the Base module code against the words of each filename,
ignoring product-line words (`pilot`, `seica`, `series`, `opera`), and prefers a
filename with nothing left over. When several fit equally badly — a bare
`Compact` matches three — it inserts nothing and says so, because a photo of the
wrong machine is worse than no photo. Record a stubborn code once in
`assets/systems/aliases.json` rather than passing `--system` forever.

This design ships no picture frame, so the skill **creates** one: a centred
paragraph under the title block, with the photo fitted inside 3.60 × 2.60 in so
it stays subordinate to the offer and cannot push the letterhead down the page.

## Files

- `scripts/build_rfq.py` — parse the export, fill the template, write the offer
- `scripts/check_rfq.py` — verify the result against the export
- `scripts/apply_theme.py` — re-assert design tokens over a drifted template
- `assets/rfq_template.docx` — the SEICA Israel commercial-offer template
- `assets/trainings.json` — the training catalogue
- `assets/theme.json` — the design system's tokens, for reference
- `assets/systems/` — machine photos, `aliases.json`, `alternates/`
- `references/format.md` — export layout, template anatomy, catalogue format

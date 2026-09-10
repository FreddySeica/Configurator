# Configurator → RFQ

Tooling for turning machine-configurator exports into customer-ready
commercial offers.

## `rfq-from-config` skill

`.claude/skills/rfq-from-config/` takes the spreadsheet the configurator
produces for a machine build and writes it into the company RFQ letterhead as
a formatted Configuration table, grouped by module class.

It fills the letterhead fields, the configuration list, the system photo and
the training line items. It deliberately does not touch prices,
service-contract figures or terms — those are commercial commitments, so a
person completes them in Word before the offer is sent.

The system photo comes from `assets/systems/`, where the filename is the
machine name (`Pilot VX.jpg`); the right one is matched from the export and
scaled to fit the template's layout box. Stocked today: Compact Digital XL,
Compact SL SC, Compact TK, Pilot BT, Pilot BTP, Pilot BTV, Pilot V8, Pilot VX,
Valid LR, Valid SL — with other views in `alternates/` and
`aliases.json` for module codes that do not resemble the filename. Training line items come from
`assets/trainings.json` and are numbered `ADV_TRAIN#1`, `#2`, … in the order
requested.

Colours and typeface come from `assets/theme.json`. The values shipped today
were sampled from the letterhead artwork, not from a design system — swap that
one file to adopt the real tokens. `--no-theme` builds without them.

```bash
pip install python-docx openpyxl xlrd

S=.claude/skills/rfq-from-config/scripts

python3 $S/build_rfq.py --list-trainings      # what can be offered
python3 $S/build_rfq.py --list-systems        # which machine photos exist

python3 $S/build_rfq.py CONFIG.xlsx \
  -o "RFQ_Customer_Machine.docx" \
  --customer "Customer Ltd. / Attn: ..." \
  --protocol "PRV 000000/V_IL rev.01" \
  --training install --training adv2w --training adv1w

python3 $S/check_rfq.py "RFQ_Customer_Machine.docx" \
  --config CONFIG.xlsx --expect-picture
```

Accepts `.xls`, `.xlsx` and `.csv` exports. See the skill's `SKILL.md` for the
full workflow and `references/format.md` for the export and template layout.

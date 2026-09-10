# Configurator → RFQ

Tooling for turning machine-configurator exports into customer-ready
commercial offers.

## `rfq-from-config` skill

`.claude/skills/rfq-from-config/` takes the spreadsheet the configurator
produces for a machine build and writes it into the company RFQ letterhead as
a formatted Configuration table, grouped by module class.

It fills the letterhead fields and the configuration list. It deliberately
does not touch pricing, service-contract figures or terms — those are
commercial commitments, so a person completes them in Word before the offer
is sent.

```bash
pip install python-docx openpyxl xlrd

python3 .claude/skills/rfq-from-config/scripts/build_rfq.py CONFIG.xlsx \
  -o "RFQ_Customer_Machine.docx" \
  --customer "Customer Ltd. / Attn: ..." \
  --protocol "PRV 000000/V_IL rev.01"

python3 .claude/skills/rfq-from-config/scripts/check_rfq.py \
  "RFQ_Customer_Machine.docx" --config CONFIG.xlsx
```

Accepts `.xls`, `.xlsx` and `.csv` exports. See the skill's `SKILL.md` for the
full workflow and `references/format.md` for the export and template layout.

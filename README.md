# Risk Register Skill

A standalone Claude Skill (and standalone GitHub project) for creating and maintaining a GRC
risk register. Pulled out of the larger `grc-toolkit` project so it can be used, explained,
or shared on its own.

## Files

```
risk-register-skill/
├── SKILL.md                              # Instructions Claude reads to use this
├── README.md                             # This file
├── assets/
│   ├── risk_register_template.xlsx       # Blank, ready to fill in
│   └── risk_register_starter_kit.xlsx    # Pre-filled with 15 realistic, scored example risks
└── scripts/
    └── populate_risk_register.py         # Appends new risk rows programmatically
```

## Quick use (in Claude)

1. Upload this as a zip: claude.ai → Settings → Customize → Skills → Upload skill
2. Make sure the toggle is ON
3. In a new chat: *"Add a risk to my risk register: [describe it]"*

## Quick use (standalone, no Claude)

Open `assets/risk_register_starter_kit.xlsx` directly in Excel — it already has 15 example
risks filled in and scored. Read the "Legend & Instructions" tab first.

To bulk-add risks via script:
```bash
pip install openpyxl
python scripts/populate_risk_register.py assets/risk_register_template.xlsx your_risks.json
```
See the docstring at the top of `populate_risk_register.py` for the expected JSON format.

## What makes this "automated"

- **Risk Score / Risk Level are Excel formulas** (Likelihood × Impact, banded into
  Low/Medium/High/Critical) — not values typed in, so the sheet recalculates itself if inputs
  change.
- **Inherent risk and residual risk are tracked separately** — the score before a control vs.
  after — so a control's actual effect is provable (e.g. Critical → Medium), not just claimed.
- **The script does the repetitive part** (adding rows, applying formulas consistently) so
  bulk-adding risks doesn't mean retyping formulas by hand.
